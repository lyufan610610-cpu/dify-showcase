"""PowerPoint (.pptx) document extractor used for RAG ingestion.

Load ``.pptx`` files into Markdown-ish text with embedded image links. Each
slide becomes its own ``Document``: the body keeps the ``## Slide N`` heading
and carries a 0-based ``page`` metadata entry, aligned with ``PdfExtractor`` so
the Track A transcription service can emit ``<!-- page: N -->`` markers and keep
引用溯源 correct. Tables become Markdown tables, and pictures are persisted
through the storage layer so the frontend can still render them after the file
has been vectorized.

视觉模型图片清洗
----------------
业务反馈 PPT 里存在大量"无关图片"（公司 Logo、水印、页眉页脚装饰、分隔线、
通用图标），它们逐页重复出现，既不承载业务信息，又会污染向量库与前端渲染。

本模块在解析阶段直接用视觉模型做一次二分类，取代原先"按跨页重复"的纯规则判定。
判定不只看图片本身：第一遍解析会同时记录每张图出现在哪几页、以及这些页面的正文，
送模型时把"出现位置 + 所在页面正文节选"作为上下文一起给出，由模型判断图片内容与
该页文字是否严重无关——最典型的就是逐页重复的公司 Logo、水印与页眉页脚装饰。
解析因此分两遍执行：

1. 第一遍只登记图片（内容哈希 -> 二进制）、出现页数与页面正文，正文写占位符，**不落盘**；
2. 第二遍把候选图片连同各自上下文分批交给视觉模型判定 ``decorative`` / ``meaningful``，
   被判为装饰性的图片既不落盘也不写入正文，其余图片落盘后把占位符替换成
   Markdown 链接，再组装 ``Document``。

模型接入复用 ``TRANSCRIBE_*`` 配置（``TRANSCRIBE_ENABLED`` /
``TRANSCRIBE_MODEL_PROVIDER`` / ``TRANSCRIBE_MODEL_NAME`` / ``TRANSCRIBE_IMAGE_DETAIL``），
不新增任何配置项。所有环节 **fail-open**：开关关闭、模型不可用、调用失败、
输出无法解析时一律保留全部图片，绝不让图片清洗影响文档解析结果。
"""

import base64
import hashlib
import logging
import mimetypes
import os
import re
import uuid
from collections.abc import Iterator, Mapping, Sequence
from typing import Any, cast, override

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from sqlalchemy.orm import Session

from configs import dify_config
from core.rag.extractor.extractor_base import BaseExtractor
from core.rag.models.document import Document
from extensions.ext_database import db
from extensions.ext_storage import storage
from extensions.storage.storage_type import StorageType
from libs.datetime_utils import naive_utc_now
from models.enums import CreatorUserRole
from models.model import UploadFile

logger = logging.getLogger(__name__)

# --------------------------------------------------------------------------------------
# 图片清洗常量
# --------------------------------------------------------------------------------------
# 正文里图片先写成占位符，等模型判定完再决定"落盘并替换成链接"还是"整条丢弃"。
# \x00 是不可能出现在正常文本里的哨兵，``str.splitlines()`` 也不会在它上面断行。
_IMAGE_PLACEHOLDER = "\x00pptx-image\x00"

# 送模型判定的单张图片字节上限：过大的图又贵又容易超时，直接保留、不判定。
_IMAGE_MAX_BYTES = 1_500_000
# 一次解析最多判定多少张图，避免超大 deck 打爆预算。
_IMAGE_MAX_CANDIDATES = 60
# 单批图片的张数上限与字节上限。
_IMAGE_BATCH_MAX_COUNT = 4
_IMAGE_BATCH_MAX_BYTES = 4_000_000

# 上下文预算：一张图列几个出现页、取几页正文、每页正文截多少字。
# 上下文是"图片与页面文字是否相关"的判断依据，太小会误判、太大会拖慢解析。
_CONTEXT_SLIDE_LIST_MAX = 6
_CONTEXT_SLIDE_TEXT_MAX = 3
_CONTEXT_CHARS_PER_SLIDE = 120

# 可交给视觉模型判定的 MIME 白名单：跳过 emf / wmf 等矢量格式（模型多不支持）。
_JUDGEABLE_MIME_TYPES: frozenset[str] = frozenset(
    {"image/png", "image/jpeg", "image/gif", "image/bmp", "image/tiff", "image/webp"}
)

# 模型若把结果包在 ``` 代码块里，需要剥掉最外层围栏。
_CODE_FENCE_PATTERN = re.compile(r"^\s*```[a-zA-Z0-9_+\-]*\s*\n(.*?)\n\s*```\s*$", re.DOTALL)

# 丢弃装饰性图片后可能留下连续空行，收敛为最多一个空行，保持段落结构。
_BLANK_LINE_RUN_PATTERN = re.compile(r"\n{3,}")

# 图片清单的开场白：明确"图片前的文字 = 该图的上下文"，避免模型把上下文串台。
_IMAGE_VERDICT_INTRO = (
    "以下按顺序给出同一份 PPT 里抽出的图片。每张图片前面紧跟的文字是它自己的上下文"
    "（出现位置与所在页面正文节选），请逐张独立判断，不要把不同图片的上下文混淆。"
)

# 判定提示词：只要一个 JSON，键是图片序号，值是 decorative / meaningful。
_IMAGE_VERDICT_PROMPT = (
    "请结合每张图片的上下文，判断这张图是否值得保留进检索用的文档正文，"
    "并只输出一个 JSON 对象。\n"
    "判 decorative（丢弃）的典型情况：\n"
    "1. 公司、机构、品牌 Logo、徽标、校徽、二维码、印章、水印；\n"
    "2. 页眉页脚装饰、页码装饰、背景花纹、装饰性分隔线、纯色色块、空白占位图；\n"
    "3. 与所在页面文字在内容上严重无关的配图，例如随手贴的通用插画、表情包、"
    "无信息量的装饰图标。\n"
    "判 meaningful（保留）的情况：与页面文字相关的图表、数据截图、流程图、架构图、"
    "界面截图、产品照片、示意图，任何看懂它能帮助理解页面内容的图片。\n"
    "特别注意：同一张图在多页重复出现，而所在页面文字又从未提到它，基本可以断定是 "
    "Logo 或装饰元素，应判 decorative。\n"
    "拿不准时判 meaningful，宁可保留，不要误删信息。\n"
    '键是图片序号（字符串，从 "1" 开始），值是 "decorative" 或 "meaningful"。\n'
    '严格只输出 JSON，不要解释，不要代码块。示例：{"1": "decorative", "2": "meaningful"}'
)

# 容错解析：兼容 JSON（"1": "decorative"）、行文本（1: decorative）与中文写法（1：装饰）。
_IMAGE_VERDICT_PATTERN = re.compile(
    r"(?<!\d)(\d{1,3})(?!\d)\s*[\"']?\s*[:：=\-]?\s*[\"']?\s*(decorative|meaningful|装饰|信息)",
    re.IGNORECASE,
)
_VERDICT_ALIASES: dict[str, str] = {
    "decorative": "decorative",
    "装饰": "decorative",
    "meaningful": "meaningful",
    "信息": "meaningful",
}

# 占位符后面紧跟 40 位 sha1；给模型看页面正文时要把它整块抹掉，只留文字。
_IMAGE_PLACEHOLDER_PATTERN = re.compile(re.escape(_IMAGE_PLACEHOLDER) + r"[0-9a-f]{0,40}")


# --------------------------------------------------------------------------------------
# 模块级工具
# --------------------------------------------------------------------------------------
def _mime_type_for_ext(ext: str) -> str:
    """按扩展名猜 MIME（``jpeg`` 也要能命中 ``image/jpeg``）。"""
    mime_type, _ = mimetypes.guess_type("image." + (ext or "png"))
    return mime_type or ""


def _supports_vision(model_instance: Any, model_name: str) -> bool:
    """判断模型是否具备视觉能力。schema 查询失败时 fail-open（返回 True）。"""
    try:
        from graphon.model_runtime.entities.model_entities import ModelFeature

        model_schema = model_instance.model_type_instance.get_model_schema(model_name, model_instance.credentials)
    except Exception:
        logger.warning("PptxExtractor 获取模型 schema 失败，按支持视觉处理。", exc_info=True)
        return True

    if model_schema is None:
        return True

    features = getattr(model_schema, "features", None)
    if not features:
        return True

    return ModelFeature.VISION in features


def _select_candidates(
    image_bank: Mapping[str, tuple[bytes, str]],
    occurrences: Mapping[str, int],
) -> list[str]:
    """挑出可送模型判定的图片：体积与格式合格，跨页重复的优先排前面。"""
    usable: list[str] = []
    for key, (blob, ext) in image_bank.items():
        if not blob or len(blob) > _IMAGE_MAX_BYTES:
            continue
        if _mime_type_for_ext(ext) not in _JUDGEABLE_MIME_TYPES:
            continue
        usable.append(key)

    if not usable:
        return []

    # 跨页重复的图片更可能是装饰性元素（Logo / 水印），优先送模型判定。
    usable.sort(key=lambda key: (-occurrences.get(key, 0), len(image_bank[key][0])))

    if len(usable) > _IMAGE_MAX_CANDIDATES:
        logger.info(
            "PptxExtractor 候选图片 %s 张，超过单次上限 %s，只判定前 %s 张。",
            len(usable),
            _IMAGE_MAX_CANDIDATES,
            _IMAGE_MAX_CANDIDATES,
        )
        usable = usable[:_IMAGE_MAX_CANDIDATES]
    return usable


def _iter_image_batches(
    image_bank: Mapping[str, tuple[bytes, str]],
    candidates: Sequence[str],
) -> Iterator[list[str]]:
    """按张数与总字节双上限把候选列表切成若干批。"""
    batch: list[str] = []
    batch_bytes = 0
    for key in candidates:
        size = len(image_bank[key][0])
        if batch and (len(batch) >= _IMAGE_BATCH_MAX_COUNT or batch_bytes + size > _IMAGE_BATCH_MAX_BYTES):
            yield batch
            batch = []
            batch_bytes = 0
        batch.append(key)
        batch_bytes += size
    if batch:
        yield batch


def _parse_image_verdicts(text: str) -> dict[int, str]:
    """从模型输出里抽出 ``{图片序号: 判定}``；解析不出来就返回空字典（fail-open）。"""
    if not text:
        return {}

    stripped = text.strip()
    fence = _CODE_FENCE_PATTERN.match(stripped)
    if fence:
        stripped = fence.group(1).strip()

    verdicts: dict[int, str] = {}
    for match in _IMAGE_VERDICT_PATTERN.finditer(stripped):
        try:
            index = int(match.group(1))
        except (TypeError, ValueError):
            continue
        if index <= 0:
            continue
        alias = _VERDICT_ALIASES.get(match.group(2).lower())
        if alias is None:
            continue
        verdicts[index] = alias
    return verdicts


def _slide_text(body: str) -> str:
    """把页面正文里的图片占位符与多余空白清掉，得到纯文字。

    这段文字会作为"该页在讲什么"的上下文送给模型，占位符对它没有意义。
    """
    if _IMAGE_PLACEHOLDER not in body:
        flattened = body
    else:
        flattened = _IMAGE_PLACEHOLDER_PATTERN.sub(" ", body)
    return " ".join(flattened.split())


def _build_image_context(
    key: str,
    occurrence_pages: Mapping[str, Sequence[int]],
    slide_texts: Mapping[int, str],
) -> str:
    """拼出一张图片的上下文：出现在哪几页，以及这些页面在讲什么。

    "图片内容与该页文字是否严重无关"是判定装饰图的主要依据，所以上下文按
    ``_CONTEXT_SLIDE_LIST_MAX`` 个页码、``_CONTEXT_SLIDE_TEXT_MAX`` 页正文、
    每页 ``_CONTEXT_CHARS_PER_SLIDE`` 字做三重截断，避免超长 deck 把提示词撑爆。
    """
    pages = list(occurrence_pages.get(key, ()))
    if not pages:
        return "出现位置：未知。"

    shown = "、".join(str(page) for page in pages[:_CONTEXT_SLIDE_LIST_MAX])
    shown += " 等页" if len(pages) > _CONTEXT_SLIDE_LIST_MAX else " 页"
    lines = [f"出现位置：第 {shown}，共 {len(pages)} 页出现。"]

    # 节选优先取最初出现的几页：逐页重复的 Logo 通常首页就带着线索。
    excerpts: list[str] = []
    for page in pages[:_CONTEXT_SLIDE_TEXT_MAX]:
        text = (slide_texts.get(page) or "").strip()
        if not text:
            continue
        if len(text) > _CONTEXT_CHARS_PER_SLIDE:
            text = text[:_CONTEXT_CHARS_PER_SLIDE] + "…"
        excerpts.append(f"第 {page} 页文字：{text}")

    if excerpts:
        lines.append("所在页面正文节选：" + " / ".join(excerpts))
    else:
        lines.append("所在页面正文节选：（这些页面没有文字）")
    return "\n".join(lines)


class PptxExtractor(BaseExtractor):
    """Load ``.pptx`` files into Markdown-ish text with embedded image links.

    Args:
        file_path: Path to the file to load.
        tenant_id: Tenant that owns the extracted images.
        user_id: Account used as the creator of the extracted images.
        session: Session used to persist extracted images.
    """

    _closed: bool
    _session: Session | None

    def __init__(self, file_path: str, tenant_id: str, user_id: str, *, session: Session | None = None):
        """Initialize with file path."""
        self._closed = False
        self.file_path = file_path
        self.tenant_id = tenant_id
        self.user_id = user_id
        self._session = session

        if "~" in self.file_path:
            self.file_path = os.path.expanduser(self.file_path)

        if not os.path.isfile(self.file_path):
            raise ValueError(f"File path {self.file_path} is not a valid file")

    @override
    def extract(self) -> list[Document]:
        return self.parse_pptx_documents(self.file_path)

    # ----------------------------------------------------------------------------------
    # 基础工具
    # ----------------------------------------------------------------------------------
    @staticmethod
    def _image_key(blob: bytes) -> str:
        """Content hash used to deduplicate identical images."""
        return hashlib.sha1(blob).hexdigest()

    @staticmethod
    def _shape_sort_key(shape) -> tuple[float, float]:
        """Order shapes by reading position: top first, then left."""
        try:
            return (shape.top or 0, shape.left or 0)
        except Exception:
            return (0, 0)

    @staticmethod
    def _normalize_ext(ext: str) -> str:
        """归一化图片扩展名（``jpg`` -> ``jpeg``，去掉前导点）。"""
        ext = (ext or "png").lstrip(".").lower()
        if ext == "jpg":
            ext = "jpeg"
        return ext

    def _save_image(self, blob: bytes, ext: str) -> str:
        """Persist one image and return its Markdown link."""
        ext = self._normalize_ext(ext)

        file_uuid = str(uuid.uuid4())
        file_key = "image_files/" + self.tenant_id + "/" + file_uuid + "." + ext
        mime_type, _ = mimetypes.guess_type(file_key)

        storage.save(file_key, blob)

        upload_file = UploadFile(
            tenant_id=self.tenant_id,
            storage_type=StorageType(dify_config.STORAGE_TYPE),
            key=file_key,
            name=file_key,
            size=len(blob),
            extension=ext,
            mime_type=mime_type or "",
            created_by=self.user_id,
            created_by_role=CreatorUserRole.ACCOUNT,
            created_at=naive_utc_now(),
            used=True,
            used_by=self.user_id,
            used_at=naive_utc_now(),
        )
        session = self._session or db.session
        session.add(upload_file)

        base_url = dify_config.FILES_URL
        link = f"![image]({base_url}/files/{upload_file.id}/file-preview)"
        logger.info("PptxExtractor saved image key=%s id=%s", file_key, upload_file.id)
        return link

    def _table_to_markdown(self, table) -> str:
        """Render a PowerPoint table as a Markdown table."""
        lines: list[str] = []
        for row_idx, row in enumerate(table.rows):
            cells = [cell.text.strip().replace("\n", " ") for cell in row.cells]
            lines.append("| " + " | ".join(cells) + " |")
            if row_idx == 0:
                lines.append("| " + " | ".join(["---"] * len(cells)) + " |")
        return "\n".join(lines)

    # ----------------------------------------------------------------------------------
    # 形状遍历（第一遍：只登记图片，不落盘）
    # ----------------------------------------------------------------------------------
    def _process_shape(
        self,
        shape,
        out: list[str],
        image_bank: dict[str, tuple[bytes, str]],
        seen: set[str],
    ) -> None:
        """Append the Markdown representation of one shape to ``out``.

        图片不再立即落盘，只在 ``image_bank`` 里登记"内容哈希 -> (二进制, 扩展名)"，
        正文写占位符；等视觉模型判定完成后由 :meth:`_persist_images` 统一落盘。
        ``seen`` 记录本页出现过的图片哈希，用于统计跨页重复次数。
        """
        try:
            shape_type = shape.shape_type
        except Exception:
            shape_type = None

        # 1) picture -> 登记到 image_bank 并写占位符（此处不落盘）
        if shape_type == MSO_SHAPE_TYPE.PICTURE:
            try:
                image = shape.image
                blob = image.blob
            except Exception:
                logger.warning("PptxExtractor skipped an unreadable picture")
                return

            key = self._image_key(blob)
            seen.add(key)
            image_bank.setdefault(key, (blob, self._normalize_ext(getattr(image, "ext", "png"))))
            out.append(_IMAGE_PLACEHOLDER + key)
            return

        # 2) table -> Markdown table
        if getattr(shape, "has_table", False):
            try:
                markdown = self._table_to_markdown(shape.table)
            except Exception:
                logger.warning("PptxExtractor failed to render a table")
                markdown = ""
            if markdown:
                out.append(markdown)
            return

        # 3) grouped shapes -> walk children
        if shape_type == MSO_SHAPE_TYPE.GROUP:
            try:
                for child in sorted(shape.shapes, key=self._shape_sort_key):
                    self._process_shape(child, out, image_bank, seen)
            except Exception:
                logger.warning("PptxExtractor failed to walk a group shape")
            return

        # 4) text frame -> raw text
        if getattr(shape, "has_text_frame", False):
            try:
                text = (shape.text_frame.text or "").strip()
            except Exception:
                text = ""
            if text:
                out.append(text)
            return

        # 5) charts, embedded objects and other non-textual elements
        out.append("<!-- non-textual element -->")

    # ----------------------------------------------------------------------------------
    # 视觉模型判定与落盘（第二遍）
    # ----------------------------------------------------------------------------------
    def _resolve_vision_model(self) -> Any | None:
        """按 ``TRANSCRIBE_*`` 配置解析出可用的视觉模型实例；不可用时返回 None。"""
        if not bool(getattr(dify_config, "TRANSCRIBE_ENABLED", False)):
            return None
        if not self.tenant_id:
            return None

        provider = (getattr(dify_config, "TRANSCRIBE_MODEL_PROVIDER", "") or "").strip()
        model_name = (getattr(dify_config, "TRANSCRIBE_MODEL_NAME", "") or "").strip()
        if not provider or not model_name:
            return None

        try:
            # 延迟导入：解析器位于 extract_processor 的模块级导入链上，
            # 在模块顶层引入模型运行时容易形成循环导入，故放到函数内。
            from core.model_manager import ModelManager
            from graphon.model_runtime.entities.model_entities import ModelType

            model_instance = ModelManager.for_tenant(tenant_id=self.tenant_id).get_model_instance(
                tenant_id=self.tenant_id,
                provider=provider,
                model_type=ModelType.LLM,
                model=model_name,
            )
        except Exception:
            logger.warning("PptxExtractor 无法获取视觉模型实例，跳过图片清洗。", exc_info=True)
            return None

        if not _supports_vision(model_instance, model_name):
            logger.warning("模型 %s/%s 不支持视觉（VISION）能力，跳过图片清洗。", provider, model_name)
            return None
        return model_instance

    def _judge_image_batch(
        self,
        model_instance: Any,
        image_bank: Mapping[str, tuple[bytes, str]],
        batch: Sequence[str],
        contexts: Mapping[str, str],
    ) -> dict[int, str]:
        """把一批图片连同各自上下文交给视觉模型二分类，返回 ``{批内序号: 判定}``。

        每张图前面先放一段它自己的上下文文本（出现位置 + 所在页面正文节选），
        再放图片本身；模型据此判断图片内容与页面文字是否严重无关。
        """
        from core.credit_usage import CreditUsageCreatedBy
        from core.model_context import use_credit_usage_metadata
        from graphon.model_runtime.entities.llm_entities import LLMResult
        from graphon.model_runtime.entities.message_entities import (
            ImagePromptMessageContent,
            PromptMessage,
            PromptMessageContentUnionTypes,
            TextPromptMessageContent,
            UserPromptMessage,
        )

        # ``detail`` 只在拿到枚举时才传，避免把 None 塞进字段导致下游渲染异常。
        image_kwargs: dict[str, Any] = {}
        detail_enum = getattr(ImagePromptMessageContent, "DETAIL", None)
        if detail_enum is not None:
            requested = (getattr(dify_config, "TRANSCRIBE_IMAGE_DETAIL", "") or "").strip().lower()
            image_kwargs["detail"] = detail_enum.HIGH if requested == "high" else detail_enum.LOW

        contents: list[PromptMessageContentUnionTypes] = [
            TextPromptMessageContent(data=_IMAGE_VERDICT_INTRO),
        ]
        for index, key in enumerate(batch, start=1):
            blob, ext = image_bank[key]
            context = contexts.get(key) or "出现位置：未知。"
            contents.append(
                TextPromptMessageContent(data=f"第 {index} 张图的上下文：\n{context}\n这张图是：")
            )
            contents.append(
                ImagePromptMessageContent(
                    format=ext,
                    base64_data=base64.b64encode(blob).decode(),
                    mime_type=_mime_type_for_ext(ext),
                    filename=f"slide-image-{index}.{ext}",
                    **image_kwargs,
                )
            )
        contents.append(TextPromptMessageContent(data=_IMAGE_VERDICT_PROMPT))

        with use_credit_usage_metadata({"created_by": CreditUsageCreatedBy.KNOWLEDGE_INDEXING}):
            result = model_instance.invoke_llm(
                prompt_messages=cast(list[PromptMessage], [UserPromptMessage(content=contents)]),
                model_parameters={"temperature": 0},
                stream=False,
            )

        # stream=False 时必然返回 LLMResult，而不是 Generator。
        if not isinstance(result, LLMResult):
            raise ValueError("PptxExtractor 图片判定：stream=False 时预期返回 LLMResult")
        return _parse_image_verdicts(result.message.get_text_content() or "")

    def _decide_decorative_images(
        self,
        image_bank: Mapping[str, tuple[bytes, str]],
        occurrences: Mapping[str, int],
        occurrence_pages: Mapping[str, Sequence[int]],
        slide_texts: Mapping[int, str],
        slide_count: int,
    ) -> set[str]:
        """用视觉模型判定装饰性图片集合；任何环节失败都返回空集合（保留全部图片）。"""
        if not image_bank:
            return set()

        model_instance = self._resolve_vision_model()
        if model_instance is None:
            return set()

        candidates = _select_candidates(image_bank, occurrences)
        if not candidates:
            return set()

        # 判定依据是"图片内容 vs 所在页面文字"，所以先给每张候选图备好上下文。
        contexts = {key: _build_image_context(key, occurrence_pages, slide_texts) for key in candidates}

        decorative: set[str] = set()
        for batch in _iter_image_batches(image_bank, candidates):
            try:
                verdicts = self._judge_image_batch(model_instance, image_bank, batch, contexts)
            except Exception:
                logger.warning("PptxExtractor 图片判定调用失败，保留该批图片。", exc_info=True)
                continue
            for index, key in enumerate(batch, start=1):
                if verdicts.get(index) == "decorative":
                    decorative.add(key)

        if decorative:
            logger.info(
                "PptxExtractor 视觉模型判定：%s 页共 %s 张候选图片，识别出 %s 张装饰性图片并丢弃。",
                slide_count,
                len(candidates),
                len(decorative),
            )
        return decorative

    def _persist_images(
        self,
        image_bank: Mapping[str, tuple[bytes, str]],
        decorative_keys: set[str],
    ) -> dict[str, str]:
        """把保留的图片写入存储，返回 ``hash -> Markdown 链接``；单个失败只记日志。"""
        links: dict[str, str] = {}
        for key, (blob, ext) in image_bank.items():
            if key in decorative_keys:
                continue
            try:
                links[key] = self._save_image(blob, ext)
            except Exception:
                logger.exception("PptxExtractor failed to persist an image")
        return links

    def _render_body(self, body: str, links: Mapping[str, str], decorative_keys: set[str]) -> str:
        """把占位符替换成图片链接；装饰性图片所在行整行丢弃。"""
        lines: list[str] = []
        for line in body.splitlines():
            index = line.find(_IMAGE_PLACEHOLDER)
            if index < 0:
                lines.append(line)
                continue

            key = line[index + len(_IMAGE_PLACEHOLDER) :].strip()
            if key in decorative_keys:
                continue
            link = links.get(key)
            if not link:
                # 落盘失败的图片同样丢弃，避免正文里留下裸占位符。
                continue
            prefix = line[:index]
            suffix = line[index + len(_IMAGE_PLACEHOLDER) + len(key) :]
            lines.append(f"{prefix}{link}{suffix}")

        return _BLANK_LINE_RUN_PATTERN.sub("\n\n", "\n".join(lines)).strip()

    # ----------------------------------------------------------------------------------
    # 主流程
    # ----------------------------------------------------------------------------------
    def parse_pptx_documents(self, pptx_path: str) -> list[Document]:
        """按页切分：每一页产出一个 ``Document``，并写入 0 基 ``page`` 元数据。

        与 ``PdfExtractor`` 的约定保持一致（0 基 enumerate），转写侧的
        ``_page_offset()`` 会再 ``+1`` 得到人类可读页码，与正文里的
        ``## Slide N`` 标题对齐。

        两遍解析：第一遍登记图片、出现页数与页面正文，正文写占位符且不落盘；第二遍由
        视觉模型结合"图片内容 + 所在页面文字"挑出装饰性图片（既不落盘也不进正文），
        其余图片落盘后替换占位符再组装文档。
        整页只剩装饰图或空白时不产出 ``Document``，避免生成只有标题的垃圾 chunk。
        """
        prs = Presentation(pptx_path)

        image_bank: dict[str, tuple[bytes, str]] = {}
        occurrences: dict[str, int] = {}
        occurrence_pages: dict[str, list[int]] = {}
        slide_texts: dict[int, str] = {}
        slide_bodies: list[tuple[int, str]] = []
        slide_count = 0

        # 第一遍：只登记图片，不落盘
        for slide_idx, slide in enumerate(prs.slides, start=1):
            slide_count = slide_idx
            parts: list[str] = []
            seen: set[str] = set()
            for shape in sorted(slide.shapes, key=self._shape_sort_key):
                self._process_shape(shape, parts, image_bank, seen)
            for key in seen:
                occurrences[key] = occurrences.get(key, 0) + 1
                occurrence_pages.setdefault(key, []).append(slide_idx)
            body = "\n".join(part for part in parts if part and part.strip())
            if not body:
                continue
            # 页面纯文字留作该页所有图片的判定上下文（占位符对模型没有意义）。
            slide_texts[slide_idx] = _slide_text(body)
            slide_bodies.append((slide_idx, body))

        # 第二遍：模型判定 -> 落盘 -> 渲染
        decorative_keys = self._decide_decorative_images(
            image_bank,
            occurrences,
            occurrence_pages,
            slide_texts,
            slide_count,
        )
        links = self._persist_images(image_bank, decorative_keys)

        documents: list[Document] = []
        for slide_idx, body in slide_bodies:
            text = self._render_body(body, links, decorative_keys)
            if not text.strip():
                continue
            documents.append(
                Document(
                    page_content=f"## Slide {slide_idx}\n{text}",
                    metadata={"source": self.file_path, "page": slide_idx - 1},
                )
            )

        if self._session is None:
            db.session.commit()

        logger.info(
            "PptxExtractor parsed %s: %d slides, %d documents, %d unique images, %d decorative dropped",
            self.file_path,
            slide_count,
            len(documents),
            len(image_bank),
            len(decorative_keys),
        )
        return documents

    def parse_pptx(self, pptx_path: str) -> str:
        """整篇文本视图（向后兼容），内容与 ``parse_pptx_documents`` 拼接一致。"""
        return "\n\n".join(document.page_content for document in self.parse_pptx_documents(pptx_path))

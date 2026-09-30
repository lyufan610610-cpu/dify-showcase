"""多模态文档转写服务（Transcription Service）。

把原生 extractor 产出的 ``list[Document]`` 交给多模态（视觉）大模型转写为统一的
Markdown，再进入 Dify 原有的向量化流程。转写只做增强：任何失败都回退到原生
解析结果，不阻断文档摄入。
"""

from __future__ import annotations

import logging
import math
import re
from collections.abc import Mapping, Sequence
from typing import Any, cast

from sqlalchemy import select
from sqlalchemy.orm import Session

from configs import dify_config
from core.app.file_access import DatabaseFileAccessController
from core.credit_usage import CreditUsageCreatedBy
from core.llm_generator.prompts import DEFAULT_TRANSCRIBE_PROMPT
from core.model_context import with_credit_usage_created_by
from core.model_manager import ModelManager
from core.rag.models.document import Document
from extensions.ext_database import db
from factories.file_factory import build_from_mapping
from graphon.file import FileTransferMethod, FileType, file_manager
from graphon.model_runtime.entities.llm_entities import LLMResult
from graphon.model_runtime.entities.message_entities import (
    ImagePromptMessageContent,
    PromptMessage,
    PromptMessageContentUnionTypes,
    TextPromptMessageContent,
    UserPromptMessage,
)
from graphon.model_runtime.entities.model_entities import ModelFeature, ModelType
from models import UploadFile

logger = logging.getLogger(__name__)

__all__ = [
    "SKIP_EXTENSIONS",
    "SUPPORTED_EXTENSIONS",
    "is_transcription_enabled",
    "transcribe_documents",
]

_file_access_controller = DatabaseFileAccessController()

# 已是结构化 Markdown 的文件无需二次转写。
SKIP_EXTENSIONS: frozenset[str] = frozenset({".md", ".markdown", ".mdx"})

# 需要走视觉模型转写的扩展名白名单（.ppt 为老二进制格式，python-pptx 不支持）。
SUPPORTED_EXTENSIONS: frozenset[str] = frozenset(
    {".pdf", ".docx", ".xlsx", ".xls", ".csv", ".txt", ".htm", ".html", ".epub", ".pptx"}
)

# 插入转写结果中的位置标记，供引用溯源使用（HTML 注释形式，渲染时不可见）。
PAGE_MARKER_TEMPLATE = "<!-- page: {page} -->"
ROW_MARKER_TEMPLATE = "<!-- row: {row} -->"

# 与 ParagraphIndexProcessor._extract_images_from_text 保持完全一致的图片链接解析规则。
_MARKDOWN_IMAGE_PATTERN = re.compile(r"!\[[^\]]*\]\(([^)]+)\)")
_IMAGE_PREVIEW_PATTERN = re.compile(r"/files/([a-f0-9\-]+)/image-preview(?:\?.*?)?")
_FILE_PREVIEW_PATTERN = re.compile(r"/files/([a-f0-9\-]+)/file-preview(?:\?.*?)?")
# 模型常常把整段 Markdown 包在 ``` 代码块里返回，需要剥掉最外层围栏。
_CODE_FENCE_PATTERN = re.compile(r"^\s*```[a-zA-Z0-9_+\-]*\s*\n(.*?)\n\s*```\s*$", re.DOTALL)

_DETAIL_ENUM = getattr(ImagePromptMessageContent, "DETAIL", None)

# "跨页重复的装饰性图片"判定阈值：去 query 后 URL 路径相同的图片，出现页数
# >= _DECORATIVE_IMAGE_MIN_PAGES 且占全篇页数比例 >= _DECORATIVE_IMAGE_MIN_RATIO。
_DECORATIVE_IMAGE_MIN_PAGES = 2
_DECORATIVE_IMAGE_MIN_RATIO = 0.5

# 剥离链接后可能留下连续空行，收敛为最多一个空行。
_BLANK_LINE_RUN_PATTERN = re.compile(r"\n{3,}")


# --------------------------------------------------------------------------------------
# 对外入口
# --------------------------------------------------------------------------------------
def is_transcription_enabled() -> bool:
    """转写总开关。默认为 False，避免未配置时影响既有行为。"""
    return bool(getattr(dify_config, "TRANSCRIBE_ENABLED", False))


def transcribe_documents(
    *,
    documents: Sequence[Document],
    upload_file: UploadFile | None,
    file_extension: str,
    session: Session | None = None,
) -> list[Document] | None:
    """把原生 extractor 的解析结果交给多模态模型转写为统一的 Markdown。

    条件不满足或转写失败时返回 None，调用方据此回退到原生 ``documents``。
    """
    if not documents:
        return None

    if not is_transcription_enabled():
        return None

    extension = (file_extension or "").lower()
    if extension in SKIP_EXTENSIONS or extension not in SUPPORTED_EXTENSIONS:
        return None

    # URL / Notion 等场景没有 UploadFile，无法定位图片，跳过。
    if upload_file is None:
        return None

    tenant_id = getattr(upload_file, "tenant_id", None)
    if not tenant_id:
        return None

    provider = _resolve_provider()
    model_name = _resolve_model_name()
    if not provider or not model_name:
        logger.info("转写已开启，但 TRANSCRIBE_MODEL_PROVIDER / TRANSCRIBE_MODEL_NAME 未配置，跳过转写。")
        return None

    try:
        return _transcribe(
            documents=documents,
            tenant_id=tenant_id,
            provider=provider,
            model_name=model_name,
            session=session,
        )
    except Exception:
        # 不向上抛异常，失败回退原生结果。
        logger.exception("文档转写失败，回退到原生解析结果。tenant_id=%s, model=%s", tenant_id, model_name)
        return None


# --------------------------------------------------------------------------------------
# 主流程
# --------------------------------------------------------------------------------------
@with_credit_usage_created_by(CreditUsageCreatedBy.KNOWLEDGE_INDEXING)
def _transcribe(
    *,
    documents: Sequence[Document],
    tenant_id: str,
    provider: str,
    model_name: str,
    session: Session | None,
) -> list[Document] | None:
    model_manager = ModelManager.for_tenant(tenant_id=tenant_id)
    model_instance = model_manager.get_model_instance(
        tenant_id=tenant_id,
        provider=provider,
        model_type=ModelType.LLM,
        model=model_name,
    )

    if not _supports_vision(model_instance, model_name):
        logger.warning("模型 %s/%s 不支持视觉（VISION）能力，跳过转写。", provider, model_name)
        return None

    prompt = _resolve_prompt()
    model_parameters = _resolve_model_parameters()
    image_detail = _resolve_image_detail()
    max_chars = _max_chars_per_chunk()
    max_chunks = _max_chunks()

    produced: list[Document] = []
    succeeded = 0

    # 业务需求：去掉"无关图片"（公司 Logo、水印、页眉页脚配图）。
    # 这类图跨页重复出现，只有在这里（唯一能看到全部页面的地方）才能判定，
    # 且必须早于 _split_chunks()：剥离后的文本再进切块与约束 2 校验，
    # 因此不会出现"转写结果缺少图片链接而整块被丢弃"。
    # 注意：剥离只作用于喂给模型的文本；转写失败时仍回退到未剥离的原文档，
    # 保证约束 1（绝不破坏原生解析结果）不被削弱。
    stripped_documents = _drop_decorative_images(documents)

    for document, source in zip(documents, stripped_documents, strict=False):
        try:
            transcribed = _transcribe_document(
                document=source,
                tenant_id=tenant_id,
                session=session,
                model_instance=model_instance,
                prompt=prompt,
                model_parameters=model_parameters,
                image_detail=image_detail,
                max_chars=max_chars,
                max_chunks=max_chunks,
            )
        except Exception:
            logger.warning("单篇文档转写失败，保留原生解析结果。", exc_info=True)
            transcribed = None

        if transcribed is None:
            produced.append(document)
        else:
            produced.append(transcribed)
            succeeded += 1

    # 全部失败则整体回退，避免产生"部分转写"的混合结果。
    if succeeded == 0:
        logger.warning("全部文档转写失败，回退到原生解析结果。")
        return None

    return produced


def _transcribe_document(
    *,
    document: Document,
    tenant_id: str,
    session: Session | None,
    model_instance: Any,
    prompt: str,
    model_parameters: Mapping[str, Any],
    image_detail: Any,
    max_chars: int,
    max_chunks: int,
) -> Document | None:
    """转写单个 Document（PDF 场景下即单页）。失败返回 None。"""
    text = document.page_content or ""
    if not text.strip():
        return None

    chunks = _split_chunks(text, max_chars)
    if not chunks:
        return None

    transcribed_chunks: list[str] = []
    succeeded = 0

    for index, chunk in enumerate(chunks):
        # 超过上限的块原样保留，保证内容不丢失。
        if max_chunks > 0 and index >= max_chunks:
            transcribed_chunks.append(chunk)
            continue

        try:
            result = _transcribe_chunk(
                text=chunk,
                image_urls=_image_urls_in_text(chunk),
                tenant_id=tenant_id,
                session=session,
                model_instance=model_instance,
                prompt=prompt,
                model_parameters=model_parameters,
                image_detail=image_detail,
            )
        except Exception:
            logger.warning("文档分块转写失败，保留该块原文。", exc_info=True)
            result = None

        if result is None:
            transcribed_chunks.append(chunk)
        else:
            transcribed_chunks.append(result)
            succeeded += 1

    if succeeded == 0:
        return None

    body = "\n\n".join(part for part in transcribed_chunks if part.strip())
    if not body.strip():
        return None

    marker = _marker_for(document.metadata)
    page_content = f"{marker}\n\n{body}" if marker else body

    return Document(page_content=page_content, metadata=dict(document.metadata or {}))


def _transcribe_chunk(
    *,
    text: str,
    image_urls: list[str],
    tenant_id: str,
    session: Session | None,
    model_instance: Any,
    prompt: str,
    model_parameters: Mapping[str, Any],
    image_detail: Any,
) -> str | None:
    """调用视觉模型转写单个文本块。失败返回 None。"""
    prompt_message_contents: list[PromptMessageContentUnionTypes] = []

    image_files = _build_image_files(tenant_id=tenant_id, image_urls=image_urls, session=session)
    for image_file in image_files:
        try:
            prompt_message_contents.append(
                file_manager.to_prompt_message_content(image_file, image_detail_config=image_detail)
            )
        except Exception:
            logger.warning("图片转换为 prompt 内容失败，跳过该图片。", exc_info=True)
            continue

    prompt_message_contents.append(TextPromptMessageContent(data=f"{prompt}\n\n{text}"))

    result = model_instance.invoke_llm(
        prompt_messages=cast(list[PromptMessage], [UserPromptMessage(content=prompt_message_contents)]),
        model_parameters=dict(model_parameters),
        stream=False,
    )

    # stream=False 时必然返回 LLMResult，而不是 Generator。
    if not isinstance(result, LLMResult):
        raise ValueError("转写：stream=False 时预期返回 LLMResult")

    content = _strip_code_fence((result.message.get_text_content() or "").strip())
    if not content:
        return None

    # 图片链接缺失时丢弃该块转写结果，保留原文。
    missing = _missing_image_urls(image_urls, content)
    if missing:
        logger.warning("转写输出丢失了 %s 个图片链接，丢弃该块转写结果以保留原文。", len(missing))
        return None

    return content


# --------------------------------------------------------------------------------------
# 文本切分与图片链接处理
# --------------------------------------------------------------------------------------
def _split_chunks(text: str, max_chars: int) -> list[str]:
    """按行累积切分，绝不切断单行，尽量保持 Markdown 结构完整。"""
    if max_chars <= 0 or len(text) <= max_chars:
        return [text] if text else []

    chunks: list[str] = []
    buffer: list[str] = []
    buffer_len = 0

    for line in text.splitlines(keepends=True):
        line_len = len(line)
        # 单行本身就超限，独立成块，避免与相邻行混在一起被截断。
        if buffer and buffer_len + line_len > max_chars:
            chunks.append("".join(buffer))
            buffer = []
            buffer_len = 0
        buffer.append(line)
        buffer_len += line_len
        if buffer_len >= max_chars:
            chunks.append("".join(buffer))
            buffer = []
            buffer_len = 0

    if buffer:
        chunks.append("".join(buffer))

    return [chunk for chunk in chunks if chunk.strip()]


def _image_urls_in_text(text: str) -> list[str]:
    """按出现顺序提取去重后的 Markdown 图片 URL。"""
    urls: list[str] = []
    seen: set[str] = set()
    for match in _MARKDOWN_IMAGE_PATTERN.finditer(text or ""):
        url = (match.group(1) or "").strip()
        if not url:
            continue
        key = _url_key(url)
        if key in seen:
            continue
        seen.add(key)
        urls.append(url)
    return urls


def _match_upload_file_id(url: str) -> str | None:
    """从 ``/files/<id>/image-preview`` 或 ``/files/<id>/file-preview`` 中取出 UploadFile ID。"""
    if not url:
        return None
    match = _IMAGE_PREVIEW_PATTERN.search(url)
    if match:
        return match.group(1)
    match = _FILE_PREVIEW_PATTERN.search(url)
    if match:
        return match.group(1)
    return None


def _url_key(url: str) -> str:
    """去掉 query string，仅比较路径部分。"""
    return (url or "").split("?", 1)[0].strip()


def _missing_image_urls(original_urls: Sequence[str], transcribed: str) -> list[str]:
    """返回原文本中有、但转写结果中缺失的图片链接。"""
    if not original_urls:
        return []

    missing: list[str] = []
    for url in original_urls:
        key = _url_key(url)
        if key and key in transcribed:
            continue
        # 回退匹配：只要 upload_file id 还在，就认为图片被保留了。
        upload_file_id = _match_upload_file_id(url)
        if upload_file_id and upload_file_id in transcribed:
            continue
        missing.append(url)
    return missing


def _strip_code_fence(text: str) -> str:
    """剥掉模型可能加上的最外层 ``` 围栏。"""
    if not text:
        return ""
    match = _CODE_FENCE_PATTERN.match(text)
    if match:
        return match.group(1).strip()
    return text


# --------------------------------------------------------------------------------------
# 装饰性图片剥离（去掉"无关图片"）
# --------------------------------------------------------------------------------------
def _drop_decorative_images(documents: Sequence[Document]) -> list[Document]:
    """剥离"跨页重复出现"的装饰性图片链接（公司 Logo、水印、页眉页脚配图）。

    判定依据只有一条：**同一张图在多页重复出现**。PPTX / PDF 每个页面是一个
    Document，Logo 这类装饰性图片天然逐页复现；正文插图极少跨页复用，
    因此该判定能把"无关图片"摘出来而不误伤正文。

    调用时机必须在 :func:`_split_chunks` 之前：剥掉的链接不会进入切块结果，
    也就不会参与约束 2（图片链接必须逐字符保留）的校验，两者不冲突。
    """
    decorative = _decorative_image_keys(documents)
    if not decorative:
        return list(documents)

    stripped: list[Document] = []
    dropped = 0
    for document in documents:
        text = document.page_content or ""
        cleaned, hits = _strip_decorative_image_links(text, decorative)
        dropped += hits
        if hits == 0:
            stripped.append(document)
        else:
            stripped.append(Document(page_content=cleaned, metadata=dict(document.metadata or {})))

    if dropped:
        logger.info(
            "已剥离 %s 处跨页重复的装饰性图片链接（识别出 %s 张无关图片）。", dropped, len(decorative)
        )
    return stripped


def _decorative_image_keys(documents: Sequence[Document]) -> set[str]:
    """按"跨页重复出现"找出装饰性图片的 URL key 集合（key 已去 query）。"""
    page_count = sum(1 for document in documents if (document.page_content or "").strip())
    if page_count < _DECORATIVE_IMAGE_MIN_PAGES:
        # 单页文档（如 docx 只有一个 Document）不存在"跨页复现"，不剥离任何图片。
        return set()

    occurrences: dict[str, int] = {}
    for document in documents:
        keys = {_url_key(url) for url in _image_urls_in_text(document.page_content or "")}
        for key in keys:
            if key:
                occurrences[key] = occurrences.get(key, 0) + 1

    threshold = max(_DECORATIVE_IMAGE_MIN_PAGES, math.ceil(page_count * _DECORATIVE_IMAGE_MIN_RATIO))
    return {key for key, count in occurrences.items() if count >= threshold}


def _strip_decorative_image_links(text: str, decorative_keys: set[str]) -> tuple[str, int]:
    """删除文本中命中装饰性图片集合的 Markdown 图片链接，返回 (新文本, 删除处数)。"""
    if not text or not decorative_keys:
        return text, 0

    hits = 0

    def _replace(match: re.Match[str]) -> str:
        nonlocal hits
        url = (match.group(1) or "").strip()
        if _url_key(url) not in decorative_keys:
            return match.group(0)
        hits += 1
        return ""

    cleaned = _MARKDOWN_IMAGE_PATTERN.sub(_replace, text)
    if hits == 0:
        return text, 0

    # 图片独占一行时，删除链接后会留下空行，收敛连续空行以免把段落结构撑散。
    cleaned = _BLANK_LINE_RUN_PATTERN.sub("\n\n", cleaned).strip("\n")
    return cleaned, hits


# --------------------------------------------------------------------------------------
# 位置标记（引用溯源）
# --------------------------------------------------------------------------------------
def _marker_for(metadata: Mapping[str, Any] | None) -> str:
    """根据原生 extractor 写入的 metadata 生成位置标记。"""
    if not metadata:
        return ""

    page = _as_index(metadata.get("page"))
    if page is not None:
        return PAGE_MARKER_TEMPLATE.format(page=page + _page_offset())

    row = _as_index(metadata.get("row"))
    if row is not None:
        return ROW_MARKER_TEMPLATE.format(row=row + _page_offset())

    return ""


def _as_index(value: Any) -> int | None:
    """把 metadata 中的页码 / 行号安全地转成整数（拒绝 bool）。"""
    if isinstance(value, bool):
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, str) and value.strip().lstrip("-").isdigit():
        return int(value.strip())
    return None


def _page_offset() -> int:
    """页码偏移。pdf_extractor 使用 0 基 enumerate，默认 +1 得到人类可读页码。"""
    value = _as_index(getattr(dify_config, "TRANSCRIBE_PAGE_OFFSET", 1))
    return 1 if value is None else value


# --------------------------------------------------------------------------------------
# 图片 → File
# --------------------------------------------------------------------------------------
def _build_image_files(*, tenant_id: str, image_urls: Sequence[str], session: Session | None) -> list[Any]:
    """把文本中的图片链接还原成可喂给视觉模型的 File 对象列表。"""
    if not image_urls:
        return []

    upload_file_ids: list[str] = []
    for url in image_urls:
        upload_file_id = _match_upload_file_id(url)
        if upload_file_id and upload_file_id not in upload_file_ids:
            upload_file_ids.append(upload_file_id)

    if not upload_file_ids:
        return []

    db_session = session if session is not None else db.session

    upload_files = db_session.scalars(
        select(UploadFile).where(
            UploadFile.id.in_(upload_file_ids),
            UploadFile.tenant_id == tenant_id,
        )
    ).all()

    files: list[Any] = []
    for upload_file in upload_files:
        mime_type = getattr(upload_file, "mime_type", None)
        if not mime_type or "image" not in mime_type:
            continue

        mapping = {
            "upload_file_id": upload_file.id,
            "transfer_method": FileTransferMethod.LOCAL_FILE.value,
            "type": FileType.IMAGE.value,
        }
        try:
            file_obj = build_from_mapping(
                mapping=mapping,
                tenant_id=tenant_id,
                access_controller=_file_access_controller,
            )
            files.append(file_obj)
        except Exception:
            logger.warning("根据 UploadFile %s 构造 File 失败，跳过。", upload_file.id, exc_info=True)
            continue

    return files


# --------------------------------------------------------------------------------------
# 模型能力与配置解析
# --------------------------------------------------------------------------------------
def _supports_vision(model_instance: Any, model_name: str) -> bool:
    """判断模型是否具备视觉能力。schema 查询失败时 fail-open（返回 True）。"""
    try:
        model_schema = model_instance.model_type_instance.get_model_schema(model_name, model_instance.credentials)
    except Exception:
        logger.warning("获取模型 schema 失败，按支持视觉处理。", exc_info=True)
        return True

    if model_schema is None:
        return True

    features = getattr(model_schema, "features", None)
    if not features:
        return True

    return ModelFeature.VISION in features


def _resolve_provider() -> str:
    return (getattr(dify_config, "TRANSCRIBE_MODEL_PROVIDER", "") or "").strip()


def _resolve_model_name() -> str:
    return (getattr(dify_config, "TRANSCRIBE_MODEL_NAME", "") or "").strip()


def _resolve_prompt() -> str:
    prompt = (getattr(dify_config, "TRANSCRIBE_PROMPT", "") or "").strip()
    if not prompt:
        prompt = DEFAULT_TRANSCRIBE_PROMPT

    language = (getattr(dify_config, "TRANSCRIBE_LANGUAGE", "") or "").strip()
    language = language or "the same language as the input content"
    try:
        return prompt.format(language=language)
    except (KeyError, IndexError, ValueError):
        return prompt


def _resolve_model_parameters() -> Mapping[str, Any]:
    temperature = getattr(dify_config, "TRANSCRIBE_TEMPERATURE", None)
    if temperature is None:
        return {}
    try:
        return {"temperature": float(temperature)}
    except (TypeError, ValueError):
        return {}


def _resolve_image_detail() -> Any:
    if _DETAIL_ENUM is None:
        return None
    detail = (getattr(dify_config, "TRANSCRIBE_IMAGE_DETAIL", "") or "").strip().lower()
    if detail == "high":
        return _DETAIL_ENUM.HIGH
    return _DETAIL_ENUM.LOW


def _max_chars_per_chunk() -> int:
    value = _as_index(getattr(dify_config, "TRANSCRIBE_MAX_CHARS_PER_CHUNK", 12000))
    return value if value and value > 0 else 12000


def _max_chunks() -> int:
    value = _as_index(getattr(dify_config, "TRANSCRIBE_MAX_CHUNKS", 0))
    return value if value and value > 0 else 0

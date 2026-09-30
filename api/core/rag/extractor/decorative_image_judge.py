"""装饰性图片的视觉判定。

把图片本身连同"它出现在哪里、附近的正文写了什么"一起交给视觉模型做
``decorative`` / ``meaningful`` 二分类，供转写链路在切块前摘除公司 Logo、水印、
页眉页脚配图等无关图片。

全链路 fail-open：单批调用失败只保留该批图片，解析不出判定等同于全部保留，
模型实例不可用或全部批次失败时 :func:`judge_images` 返回 ``None``，由调用方
决定回退策略。本模块只做判定，不修改任何 ``Document``，也不执行剥离动作。
"""

from __future__ import annotations

import logging
import re
from collections.abc import Callable, Iterator, Mapping, Sequence
from typing import Any, cast

from graphon.model_runtime.entities.llm_entities import LLMResult
from graphon.model_runtime.entities.message_entities import (
    PromptMessage,
    PromptMessageContentUnionTypes,
    TextPromptMessageContent,
    UserPromptMessage,
)

logger = logging.getLogger(__name__)

DECORATIVE = "decorative"
MEANINGFUL = "meaningful"

# --------------------------------------------------------------------------------------
# 送模型判定的规模上限（与 PptxExtractor 对齐）
# --------------------------------------------------------------------------------------
# 单张图片字节上限：过大的图又贵又容易超时，直接保留、不判定。
MAX_IMAGE_BYTES = 1_500_000
# 一次解析最多判定多少张图，避免超大文档打爆预算。
MAX_CANDIDATES = 60
# 单批图片的张数上限与字节上限。
BATCH_MAX_COUNT = 4
BATCH_MAX_BYTES = 4_000_000

# 可交给视觉模型判定的 MIME 白名单：跳过 emf / wmf 等矢量格式（模型多不支持）。
JUDGEABLE_MIME_TYPES: frozenset[str] = frozenset(
    {"image/png", "image/jpeg", "image/gif", "image/bmp", "image/tiff", "image/webp"}
)

# 判定输出温度固定为 0，保证同图同判。
JUDGE_MODEL_PARAMETERS: Mapping[str, Any] = {"temperature": 0}

# 模型若把结果包在 ``` 代码块里，需要剥掉最外层围栏。
_CODE_FENCE_PATTERN = re.compile(r"^\s*```[a-zA-Z0-9_+\-]*\s*\n(.*?)\n\s*```\s*$", re.DOTALL)

# 图片清单的开场白：明确"图片前的文字 = 该图的上下文"，避免模型把上下文串台。
JUDGE_INTRO = (
    "以下按顺序给出同一份文档里抽出的图片。每张图片前面紧跟的文字是它自己的上下文"
    "（出现位置与所在页/段正文节选），请逐张独立判断，不要把不同图片的上下文混淆。"
)

# 判定提示词：只要一个 JSON，键是图片序号，值是 decorative / meaningful。
# 与 PptxExtractor._IMAGE_VERDICT_PROMPT 判定口径一致，仅把"PPT"改成文档通用表述。
JUDGE_PROMPT = (
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
_VERDICT_PATTERN = re.compile(
    r"(?<!\d)(\d{1,3})(?!\d)\s*[\"']?\s*[:：=\-]?\s*[\"']?\s*(decorative|meaningful|装饰|信息)",
    re.IGNORECASE,
)
_VERDICT_ALIASES: dict[str, str] = {
    "decorative": DECORATIVE,
    "装饰": DECORATIVE,
    "meaningful": MEANINGFUL,
    "信息": MEANINGFUL,
}


# --------------------------------------------------------------------------------------
# 候选筛选与分批
# --------------------------------------------------------------------------------------
def normalize_mime_type(mime_type: str | None) -> str:
    """规范化 MIME：去掉参数、转小写，并把 image/jpg 归到 image/jpeg。"""
    normalized = (mime_type or "").split(";", 1)[0].strip().lower()
    if normalized == "image/jpg":
        normalized = "image/jpeg"
    return normalized


def is_judgeable_mime(mime_type: str | None) -> bool:
    """判断该 MIME 是否在白名单内（可交给视觉模型判定）。"""
    return normalize_mime_type(mime_type) in JUDGEABLE_MIME_TYPES


def select_candidates(*, occurrences: Mapping[str, int], sizes: Mapping[str, int]) -> list[str]:
    """挑出可送模型判定的图片，跨页重复的优先，最多 ``MAX_CANDIDATES`` 个。"""
    usable = [key for key in sizes if sizes[key] > 0 and sizes[key] <= MAX_IMAGE_BYTES]
    if not usable:
        return []

    usable.sort(key=lambda key: (-occurrences.get(key, 0), sizes[key]))

    if len(usable) > MAX_CANDIDATES:
        logger.info(
            "装饰图判定：候选图片 %s 张，超过单次上限 %s，只判定前 %s 张。",
            len(usable),
            MAX_CANDIDATES,
            MAX_CANDIDATES,
        )
        usable = usable[:MAX_CANDIDATES]
    return usable


def iter_batches(*, sizes: Mapping[str, int], candidates: Sequence[str]) -> Iterator[list[str]]:
    """按张数与总字节双上限把候选列表切成若干批。"""
    batch: list[str] = []
    batch_bytes = 0
    for key in candidates:
        size = sizes.get(key, 0)
        if batch and (len(batch) >= BATCH_MAX_COUNT or batch_bytes + size > BATCH_MAX_BYTES):
            yield batch
            batch = []
            batch_bytes = 0
        batch.append(key)
        batch_bytes += size
    if batch:
        yield batch


# --------------------------------------------------------------------------------------
# 模型输出解析
# --------------------------------------------------------------------------------------
def parse_image_verdicts(text: str) -> dict[int, str]:
    """从模型输出里抽出 ``{图片序号: 判定}``；解析不出来就返回空字典（fail-open）。"""
    if not text:
        return {}

    stripped = text.strip()
    fence = _CODE_FENCE_PATTERN.match(stripped)
    if fence:
        stripped = fence.group(1).strip()

    verdicts: dict[int, str] = {}
    for match in _VERDICT_PATTERN.finditer(stripped):
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


# --------------------------------------------------------------------------------------
# 判定编排
# --------------------------------------------------------------------------------------
def build_batch_messages(
    *,
    keys: Sequence[str],
    contexts: Mapping[str, str],
    contents: Mapping[str, PromptMessageContentUnionTypes],
) -> list[PromptMessageContentUnionTypes]:
    """拼出一批图片的判定消息：每张图先给上下文，再给图片本身。

    ``keys`` 必须与 ``contents`` 的键一致，否则模型看到的序号会与回填序号错位。
    """
    messages: list[PromptMessageContentUnionTypes] = [TextPromptMessageContent(data=JUDGE_INTRO)]
    for index, key in enumerate(keys, start=1):
        context = (contexts.get(key) or "").strip() or "出现位置：未知。"
        messages.append(TextPromptMessageContent(data=f"第 {index} 张图的上下文：\n{context}\n这张图是："))
        messages.append(contents[key])
    messages.append(TextPromptMessageContent(data=JUDGE_PROMPT))
    return messages


def judge_images(
    *,
    model_instance: Any,
    occurrences: Mapping[str, int],
    sizes: Mapping[str, int],
    contexts: Mapping[str, str],
    load_image_contents: Callable[[Sequence[str]], Mapping[str, PromptMessageContentUnionTypes]],
    log_prefix: str = "装饰图判定",
) -> set[str] | None:
    """用视觉模型判定装饰性图片，返回判定为装饰性的 key 集合。

    返回 ``None`` 表示判定不可用（没有候选或全部批次失败），调用方据此回退到其他策略。
    """
    if not occurrences:
        return set()

    candidates = select_candidates(occurrences=occurrences, sizes=sizes)
    if not candidates:
        return None

    decorative: set[str] = set()
    judged_batches = 0

    for batch in iter_batches(sizes=sizes, candidates=candidates):
        contents = load_image_contents(batch)
        ordered = [key for key in batch if key in contents]
        if not ordered:
            continue

        try:
            verdicts = _judge_one_batch(
                model_instance=model_instance,
                keys=ordered,
                contexts=contexts,
                contents=contents,
            )
        except Exception:
            logger.warning("%s：模型调用失败，保留该批 %s 张图片。", log_prefix, len(ordered), exc_info=True)
            continue

        judged_batches += 1
        for index, key in enumerate(ordered, start=1):
            if verdicts.get(index) == DECORATIVE:
                decorative.add(key)

    if judged_batches == 0:
        # 一张都没判成：交给调用方决定回退策略（不在这里替它做决定）。
        logger.warning("%s：全部批次判定失败，本次不剥离任何图片。", log_prefix)
        return None

    logger.info(
        "%s：候选 %s 张（可判定 %s 张，成功判定 %s 批），识别出 %s 张装饰性图片。",
        log_prefix,
        len(occurrences),
        len(candidates),
        judged_batches,
        len(decorative),
    )
    return decorative


def _judge_one_batch(
    *,
    model_instance: Any,
    keys: Sequence[str],
    contexts: Mapping[str, str],
    contents: Mapping[str, PromptMessageContentUnionTypes],
) -> dict[int, str]:
    """把一批图片连同各自上下文交给视觉模型二分类，返回 ``{批内序号: 判定}``。"""
    messages = build_batch_messages(keys=keys, contexts=contexts, contents=contents)

    result = model_instance.invoke_llm(
        prompt_messages=cast(list[PromptMessage], [UserPromptMessage(content=messages)]),
        model_parameters=dict(JUDGE_MODEL_PARAMETERS),
        stream=False,
    )

    # stream=False 时必然返回 LLMResult，而不是 Generator。
    if not isinstance(result, LLMResult):
        raise ValueError("装饰图判定：stream=False 时预期返回 LLMResult")
    return parse_image_verdicts(result.message.get_text_content() or "")

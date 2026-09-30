"""Functionality for splitting text."""

from __future__ import annotations

import codecs
import re
from collections.abc import Set as AbstractSet
from typing import Any, Literal, override

from core.model_manager import ModelInstance
from core.rag.splitter.text_splitter import RecursiveCharacterTextSplitter
from graphon.model_runtime.model_providers.base.tokenizers.gpt2_tokenizer import GPT2Tokenizer

# ---------------------------------------------------------------------------
# Markdown 结构感知切分相关常量
# ---------------------------------------------------------------------------
# Markdown ATX 标题（# ~ ######）
_HEADING_RE = re.compile(r"^(#{1,6})[ \t]+(.*?)[ \t]*#*[ \t]*$")
# 代码围栏，避免把代码块里的 "#" 误判成标题
_HEADING_FENCE_RE = re.compile(r"^\s*(`{3,}|~{3,})")
# 章节路径标记（写入 chunk 正文，供检索与引用定位）
_HEADING_PATH_PREFIX = "【章节路径】"
_HEADING_PATH_SEPARATOR = " > "
_MAX_HEADING_PATH_CHARS = 200


class EnhanceRecursiveCharacterTextSplitter(RecursiveCharacterTextSplitter):
    """
    This class is used to implement from_gpt2_encoder, to prevent using of tiktoken
    """

    @classmethod
    def from_encoder[T: EnhanceRecursiveCharacterTextSplitter](
        cls: type[T],
        embedding_model_instance: ModelInstance | None,
        allowed_special: Literal["all"] | AbstractSet[str] = frozenset(),
        disallowed_special: Literal["all"] | AbstractSet[str] = "all",
        **kwargs: Any,
    ) -> T:
        def _token_encoder(texts: list[str]) -> list[int]:
            if not texts:
                return []

            if embedding_model_instance:
                return embedding_model_instance.get_text_embedding_num_tokens(texts=texts)
            else:
                return [GPT2Tokenizer.get_num_tokens(text) for text in texts]

        def _character_encoder(texts: list[str]) -> list[int]:
            if not texts:
                return []

            return [len(text) for text in texts]

        _ = _token_encoder  # kept for future token-length wiring
        return cls(length_function=_character_encoder, **kwargs)


class FixedRecursiveCharacterTextSplitter(EnhanceRecursiveCharacterTextSplitter):
    def __init__(
        self,
        fixed_separator: str = "\n\n",
        separators: list[str] | None = None,
        markdown_aware: bool = True,
        max_heading_path_chars: int = _MAX_HEADING_PATH_CHARS,
        **kwargs: Any,
    ):
        """Create a new TextSplitter.

        Args:
            fixed_separator: 首层强制切分标识符（来自知识库「分段标识符」）。
            separators: 递归切分时使用的次级分隔符列表。
            markdown_aware: 是否启用 Markdown 结构感知（标题栈 + 章节路径注入）。
            max_heading_path_chars: 章节路径最大长度，超出时保留最深层标题。
        """
        super().__init__(**kwargs)
        self._fixed_separator = codecs.decode(fixed_separator, "unicode_escape")
        self._separators = separators or ["\n\n", "\n", "。", ". ", " ", ""]
        self._markdown_aware = markdown_aware
        self._max_heading_path_chars = max_heading_path_chars

    @override
    def split_text(self, text: str) -> list[str]:
        """Split incoming text and return chunks.

        与上游实现的关键差异：

        1. 合并式切分：小块先进入缓冲区，用 ``_merge_splits`` 合并到不超过
           chunk_size；超过 chunk_size 的大块才交给 ``recursive_split_text``。
        2. Markdown 结构感知：切分前先按 ATX 标题维护标题栈，在每个章节起始处注入
           「章节路径」标记，避免标题成为孤立片段，并让每个 chunk 自带章节上下文。
        """
        if not text:
            return []

        if self._markdown_aware:
            text = self._inject_heading_paths(text)

        if self._fixed_separator:
            chunks = text.split(self._fixed_separator)
        else:
            chunks = [text]

        final_chunks: list[str] = []
        pending_splits: list[str] = []
        pending_lengths: list[int] = []

        def _flush_pending() -> None:
            if not pending_splits:
                return
            final_chunks.extend(
                self._merge_splits(
                    list(pending_splits),
                    self._fixed_separator or "\n\n",
                    list(pending_lengths),
                )
            )
            pending_splits.clear()
            pending_lengths.clear()

        chunks_lengths = self._length_function(chunks)
        for chunk, chunk_length in zip(chunks, chunks_lengths):
            if chunk_length > self._chunk_size:
                # 超大块：先把已积累的小块合并输出，保证结果顺序与原文一致
                _flush_pending()
                final_chunks.extend(self.recursive_split_text(chunk))
            elif chunk_length == 0:
                # 连续分隔符产生的空片段直接丢弃，避免产生空 chunk
                continue
            else:
                pending_splits.append(chunk)
                pending_lengths.append(chunk_length)

        _flush_pending()

        return [chunk for chunk in final_chunks if chunk]

    def _inject_heading_paths(self, text: str) -> str:
        """在 Markdown 标题前注入「章节路径」标记，并让标题紧跟其后的首个内容块。

        处理后标题不会再被「分段标识符」（通常是空行）切成孤立小块，且每个
        chunk 都带有完整的章节路径，便于检索与引用溯源。
        """
        if "#" not in text:
            # 没有 Markdown 标题（例如 Word 直出文本），保持原样
            return text

        lines = text.split("\n")
        out: list[str] = []
        heading_stack: list[tuple[int, str]] = []
        in_fence = False
        fence_token = ""
        index = 0

        while index < len(lines):
            line = lines[index]

            fence_match = _HEADING_FENCE_RE.match(line)
            if fence_match:
                token = fence_match.group(1)
                if not in_fence:
                    in_fence = True
                    fence_token = token
                elif token == fence_token:
                    in_fence = False
                out.append(line)
                index += 1
                continue

            heading_match = None if in_fence else _HEADING_RE.match(line)
            if heading_match is None:
                out.append(line)
                index += 1
                continue

            level = len(heading_match.group(1))
            title = heading_match.group(2).strip()
            while heading_stack and heading_stack[-1][0] >= level:
                heading_stack.pop()
            heading_stack.append((level, title))

            path = _HEADING_PATH_SEPARATOR.join(item[1] for item in heading_stack).strip()
            # 幂等保护：父子分段会对父块再次切分，父块已带标记时不再重复注入
            already_marked = bool(out) and out[-1].startswith(_HEADING_PATH_PREFIX)
            if path and not already_marked:
                if len(path) > self._max_heading_path_chars:
                    path = "…" + path[-self._max_heading_path_chars :]
                out.append(f"{_HEADING_PATH_PREFIX}{path}")

            out.append(line)

            # 跳过标题后面的空行，让「章节路径 + 标题 + 首个内容块」落在同一片段内
            next_index = index + 1
            while next_index < len(lines) and not lines[next_index].strip():
                next_index += 1
            index = next_index

        return "\n".join(out)

    def recursive_split_text(self, text: str) -> list[str]:
        """Split incoming text and return chunks."""

        final_chunks = []
        separator = self._separators[-1]
        new_separators = []

        for i, _s in enumerate(self._separators):
            if _s == "":
                separator = _s
                break
            if _s in text:
                separator = _s
                new_separators = self._separators[i + 1 :]
                break

        # Now that we have the separator, split the text
        if separator:
            if separator == " ":
                splits = re.split(r" +", text)
            else:
                splits = text.split(separator)
            if self._keep_separator:
                splits = [s + separator for s in splits[:-1]] + splits[-1:]
        else:
            splits = list(text)
        if separator == "\n":
            splits = [s for s in splits if s != ""]
        else:
            splits = [s for s in splits if (s not in {"", "\n"})]
        _good_splits = []
        _good_splits_lengths = []  # cache the lengths of the splits
        _separator = "" if self._keep_separator else separator
        s_lens = self._length_function(splits)
        if separator != "":
            for s, s_len in zip(splits, s_lens):
                if s_len < self._chunk_size:
                    _good_splits.append(s)
                    _good_splits_lengths.append(s_len)
                else:
                    if _good_splits:
                        merged_text = self._merge_splits(_good_splits, _separator, _good_splits_lengths)
                        final_chunks.extend(merged_text)
                        _good_splits = []
                        _good_splits_lengths = []
                    if not new_separators:
                        final_chunks.append(s)
                    else:
                        other_info = self._split_text(s, new_separators)
                        final_chunks.extend(other_info)

            if _good_splits:
                merged_text = self._merge_splits(_good_splits, _separator, _good_splits_lengths)
                final_chunks.extend(merged_text)
        else:
            current_part = ""
            current_length = 0
            overlap_part = ""
            overlap_part_length = 0
            for s, s_len in zip(splits, s_lens):
                if current_length + s_len <= self._chunk_size - self._chunk_overlap:
                    current_part += s
                    current_length += s_len
                elif current_length + s_len <= self._chunk_size:
                    current_part += s
                    current_length += s_len
                    overlap_part += s
                    overlap_part_length += s_len
                else:
                    final_chunks.append(current_part)
                    current_part = overlap_part + s
                    current_length = s_len + overlap_part_length
                    overlap_part = ""
                    overlap_part_length = 0
            if current_part:
                final_chunks.append(current_part)

        return final_chunks

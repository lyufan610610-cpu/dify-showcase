# Dify 1.17.0 · PPTX 原生解析与 Track A 多模态转写

把上传的文档先交给视觉语言模型转写成统一的 Markdown，再进入 Dify 原有的清洗与向量化流程；原文中的图片链接逐字保留，前端渲染不受影响。

这个仓库是 **Dify 1.17.0 的 fork**，在官方 `1.17.0` 标签之上叠加了 **6 个提交**。它不是一份一次性的大 diff，而是刻意按演进顺序切分的：每个提交都能独立读懂、独立回退，`git show` 看到的就是当时真实的改动与当时的取舍，包括中途走错又修回来的地方。

- **基线**：`09a855dcef` — `chore: bump version to 1.17.0 (#41247)`，即 `1.17.0` 标签
- **叠加**：6 个提交，涉及 13 个文件（12 个 Track A + 1 个相互独立的 weaviate 修复）
- **范围**：只有 **Track A**（多模态转写）；语义分块（Track B）不在本仓库
- 上游 README 已移到 [`README.upstream.md`](./README.upstream.md)

## 提交演进

| # | commit | 主题 | 规模 |
| --- | --- | --- | --- |
| 1 | `857adea984` | feat(rag): 引入 PPTX 原生解析器并预留转写接入点 | 4 files, +221/-1 |
| 2 | `34c02ebf25` | feat(rag): 接入 Track A 转写入口，收敛办公扩展名白名单 | 3 files, +84/-61 |
| 3 | `72f822dbbf` | fix(rag): 补回 pptx 原生分支并显式断言 upload_file 必填 | 1 file, +5/-0 |
| 4 | `828a119eb9` | feat(rag): Track A 多模态转写服务落地，双轨制改造主体 | 9 files, +1538/-51 |
| 5 | `dc07e90e69` | feat(rag): 装饰性图片改走视觉模型判定，Track A 定稿 | 3 files, +588/-11 |
| 6 | `8b3ee8486b` | fix(vdb): 修复 weaviate 删除失效导致的旧向量残留 | 1 file, +48/-11 |

```bash
git log --oneline 09a855dcef..HEAD
git show 828a119eb9                 # 主体阶段
git diff 09a855dcef..HEAD --stat    # 全量改动
```

## 逐阶段说明

### 1 · 先把 pptx 读进来（`857adea984`）

新增 `api/core/rag/extractor/pptx_extractor.py`（201 行）：用 python-pptx 遍历每页形状，抽出文本框与备注，保留图片引用，统一产出 Dify 的 `Document`。`extract_processor.py` 注册 `.pptx` 分支，`constants/__init__.py` 的扩展名白名单补上 `pptx`。

这一阶段也如实带进了两处**与转写无关**的东西：`feature/__init__.py` 新增的 `WEBSOCKET_MAX_HTTP_BUFFER_SIZE`（Engine.IO 的 HTTP 缓冲区，用于协作请求的大 payload），以及一处拼写回归 `Unstructure  dWordExtractor`。两者都在后续阶段被处理掉，保留原样是为了让演进轨迹真实。

### 2 · 找到接入点（`34c02ebf25`）

修掉上一阶段的拼写回归，把白名单从 `pptx / ppt / doc` 收敛回只留 `pptx`（`ppt` 是老二进制格式，python-pptx 读不了）。

真正的重点是 `extract_processor.py`：撤掉 `.pptx` 直出分支，改为在原生 extractor 产出 `list[Document]` 之后接入 `transcribe_documents()`。选这个位置的理由写在代码注释里 —— 它是 paragraph / parent_child / qa 三个 index processor 的唯一收口，且此时图片已经落存储、`UploadFile` 已经入库，转写可以直接把 `![image](/files/<id>/file-preview)` 转成视觉模型可读的输入。接入点外围有三层防御：函数内延迟导入规避循环依赖、外层 `try` 兜底、`transcribe_documents()` 自身保证不抛异常。

`pptx_extractor.py` 同时补齐模块与类 docstring（含 Args）、`_closed` / `_session` 类型注解，`session` 改为 keyword-only，`extract()` 收敛成单行。

### 3 · 补回被顺手删掉的分支（`72f822dbbf`）

阶段 2 撤 `.pptx` 直出分支时，`upload_file` 的显式前置条件也跟着消失了。这里恢复 `.pptx` 原生分支，并加 `assert upload_file is not None, "upload_file is required"`，让约束写在入口而不是靠下游的 None 检查兜底。

### 4 · 主体：多模态转写服务（`828a119eb9`）

新增 `api/core/rag/extractor/transcribe_service.py`（690 行），是整个改造的核心：

- 以**整篇文档**为单位转写，而不是以 chunk 为单位，保证跨页、跨块的上下文连贯
- 对外只暴露 `is_transcription_enabled()` 与 `transcribe_documents()`
- `SKIP_EXTENSIONS` 跳过 `.md / .markdown / .mdx`（本来就是 Markdown，无需转写）
- `_split_chunks()` 按行累积、绝不切断单行，超长文档按 `TRANSCRIBE_MAX_CHARS_PER_CHUNK` 分片
- 用 `<!-- page: N -->` 标注页码（配合 `TRANSCRIBE_PAGE_OFFSET`），表格行用 `<!-- row: N -->`
- `_CODE_FENCE_PATTERN` 剥掉模型自己加上去的最外层代码围栏

同阶段配套：

| 文件 | 变化 |
| --- | --- |
| `pptx_extractor.py` | 206 → 706 行。`_IMAGE_PLACEHOLDER` + `_render_body` 把图片回填原位；`_save_image` / `_persist_images` 先按内容 sha1 去重；装饰图视觉判定两遍（每批 ≤4 张、≤4 MB）；`_shape_sort_key` 稳定排序；`_table_to_markdown` 还原表格 |
| `pdf_extractor.py` | 内容 sha1 跨整篇去重，同一个 Logo 不再逐页重复落盘 |
| `fixed_text_splitter.py` | 149 → 280 行。合并式切分 + Markdown 结构感知：ATX 标题栈、`【章节路径】` 前缀注入、` > ` 分隔 |
| `index_processor_base.py` / `indexing_runner.py` | 各 2 行，切分调用处传入 `markdown_aware=True` |
| `prompts.py` | 新增 `DEFAULT_TRANSCRIBE_PROMPT`，10 条硬约束 |
| `feature/__init__.py` | 新增 `TRANSCRIBE_*` 配置项，并撤销阶段 1 夹带的 `WEBSOCKET_MAX_HTTP_BUFFER_SIZE` |
| `test_text_splitter.py` | 6 个新用例 |

### 5 · 装饰图不再靠猜（`dc07e90e69`）

第一版判定装饰图只用一条启发式：同一张图跨页重复出现。它覆盖不全 —— 只在封面出现一次的 Logo、只出现在 2/5 页的水印、页脚装饰线条都会漏掉，而这些图同样会吃掉视觉模型的输入预算、把注意力引向无意义内容。

于是新增 `decorative_image_judge.py`（322 行），把「图片内容 + 所在位置 + 附近正文」一起交给视觉模型做 decorative / meaningful 二分类：候选先按 MIME 与体积过滤（单张 >1.5 MB 跳过，候选上限 60 张），每批 ≤4 张且 ≤4 MB，提示词与 `PptxExtractor` 同一口径。全链路 fail-open —— 单批失败只丢该批、解析不出判定等价全部保留、整体不可用返回 `None` 交调用方决定。

`transcribe_service.py`（690 → 938 行）的 `_drop_decorative_images()` 相应改为**视觉优先、启发式兜底**：取到判定就用判定，取不到才回落到「跨页重复」的频率启发式。开关是 `TRANSCRIBE_DECORATIVE_JUDGE`（默认 `true`）。

### 6 · 顺带修掉的 weaviate 删除失效（`8b3ee8486b`）

与 Track A 没有依赖，可以单独摘掉。原实现写入时用内容哈希（UUID5）当对象 id，删除时收到的却是 Dify 节点 id，两者永远对不上 —— 删除事实上无效，重复索引会不断累积无人引用的残留向量。改动后写入优先取 `metadata["doc_id"]` 作为对象 id，删除按 `doc_id` 属性批量清理（每批 100 条），对 UUID 形态的入参保留旧的 `by_id` 行为，删除失败统一打 `warning` 而不是静默吞掉。

## 设计上的三条硬约束

1. **永不向上抛异常。** 模型超时、provider 未配置、返回内容不合规，一律拦下并回落到原生解析结果。
2. **图片链接必须逐字符保留。** 转写输出若丢了原文的图片链接，该结果整体被丢弃。唯一例外是装饰性图片 —— 它们在切分前就被剥离，因此不进这条校验。
3. **不得导入 `ParagraphIndexProcessor`。** 会形成循环导入，所以图片正则在这套代码里独立实现。

## 配置

把下面这些追加到 `docker/.env`（该文件被 `docker/.gitignore` 的 `*.env` 规则排除，不在仓库内）：

```
TRANSCRIBE_ENABLED=true
TRANSCRIBE_MODEL_PROVIDER=langgenius/openai_api_compatible/openai_api_compatible
TRANSCRIBE_MODEL_NAME=qwen3.8-27b
TRANSCRIBE_LANGUAGE=中文
TRANSCRIBE_IMAGE_DETAIL=low
TRANSCRIBE_DECORATIVE_JUDGE=true
TRANSCRIBE_MAX_CHARS_PER_CHUNK=12000
TRANSCRIBE_MAX_CHUNKS=0
TRANSCRIBE_PAGE_OFFSET=1
```

| 变量 | 默认值 | 含义 |
| --- | --- | --- |
| `TRANSCRIBE_ENABLED` | `false` | 总开关，关闭时 `transcribe_documents()` 直接返回 `None`，走原生解析 |
| `TRANSCRIBE_MODEL_PROVIDER` | 空 | provider id，需与控制台中已配置的 provider 一致 |
| `TRANSCRIBE_MODEL_NAME` | 空 | 模型名，需是所选 provider 下的视觉模型 |
| `TRANSCRIBE_PROMPT` | 空 | 自定义提示词，留空使用 `DEFAULT_TRANSCRIBE_PROMPT`，支持 `{language}` 占位 |
| `TRANSCRIBE_LANGUAGE` | 空 | 输出语言，留空跟随原文 |
| `TRANSCRIBE_TEMPERATURE` | 空 | 采样温度，留空用模型默认 |
| `TRANSCRIBE_IMAGE_DETAIL` | `low` | 图片细节档位 |
| `TRANSCRIBE_DECORATIVE_JUDGE` | `true` | 是否让视觉模型逐张判定装饰性图片（Logo、水印、页眉页脚配图）。关闭后回退到「同一图片跨页重复出现」的廉价启发式 |
| `TRANSCRIBE_MAX_CHARS_PER_CHUNK` | `12000` | 超长文档的分片阈值，转写按片进行 |
| `TRANSCRIBE_MAX_CHUNKS` | `0` | 分片数上限，`0` 表示不限制 |
| `TRANSCRIBE_PAGE_OFFSET` | `1` | 页码起始值，用于生成 `<!-- page: n -->` 标记 |

`TRANSCRIBE_MODEL_PROVIDER` 与 `TRANSCRIBE_MODEL_NAME` 留空时转写会被跳过并回落到原生解析，不会报错。

## 怎么验证

```bash
# 整体改动范围
git diff 09a855dcef..HEAD --stat

# 开关是否生效
docker compose exec api python -c "from configs import dify_config; print(dify_config.TRANSCRIBE_ENABLED)"

# 切分器用例
docker compose exec api uv run pytest tests/unit_tests/core/rag/splitter/test_text_splitter.py -q
```

功能侧：在知识库中上传一份带图片的 pptx 或 pdf，索引完成后查看 api 日志里的转写记录，并确认文档分段中出现了 `【章节路径】` 前缀与 `<!-- page: n -->` 标记。

配置项由 api 与 worker 读取，改完要重建这两个服务：

```bash
docker compose build api worker
docker compose up -d api worker
```

## 边界

- **不含 Track B。** `fixed_text_splitter.py` 里没有任何语义分块逻辑。`index_processor_base.py` 与 `indexing_runner.py` 中留有 4 处 `Track A+B 改造` 注释，是历史命名遗留，不影响行为。
- **失败不阻断，是刻意的设计选择**，不是缺失的容错。
- **装饰图判定有额外开销。** 转写侧每篇文档最多多出 15 次视觉模型调用（每批最多 4 张图），单张超过 1.5 MB 或候选超过 60 张的部分会被跳过。想省掉这笔开销就把 `TRANSCRIBE_DECORATIVE_JUDGE` 设为 `false`。
- **只针对 1.17.0。** 上游后续版本若改动了 `extract_processor.py` 或 index processor，需要重新核对上下文。

## 与上游的关系

基于 [Dify](https://github.com/langgenius/dify) `1.17.0`（commit `09a855dcef`）。本仓库只叠加了上述 6 个提交，其余内容与上游保持一致；许可证沿用上游，见仓库根目录的 `LICENSE`。

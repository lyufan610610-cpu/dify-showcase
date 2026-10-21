# Dify Multimodal RAG & Intelligent Document Transcription Extension

[English](README.md) | [中文](README_zh.md)

An enterprise-grade multimodal extension based on **Dify 1.17.0**. Because **upstream Dify provides zero native support for PowerPoint (`.pptx`) documents**, this repository builds a custom PPTX parsing engine from scratch, introduces Vision-Language Model (VLM) visual triage, provides end-to-end multimodal document transcription, and implements structure-aware chunking.

---

## Overview

In enterprise knowledge bases, PowerPoint presentations (`.pptx`) represent a primary format for technical reports, product architectures, and business roadmaps. However, deploying slide decks into standard RAG systems faces steep obstacles:
1. **Zero Upstream Support in Dify**: Upstream Dify does not support `.pptx` documents out of the box—its file whitelist and extractors lack PowerPoint handling entirely.
2. **Loss of Spatial & Visual Context**: Conventional document parsers flatten slides into linear text streams, destroying multi-column layouts, tabular relationships, and visual notes.
3. **Severe Visual Noise Contamination**: Enterprise slide decks are inundated with repetitive decorative graphics (corporate logos, recurring header/footer elements, watermark icons, and cosmetic dividers). If ingested unchecked, these images waste vector database storage, pollute multi-modal indices, and degrade retrieval precision.
4. **Traceability Degradation**: Standard chunkers split content arbitrarily across boundaries, stripping chunk context and breaking page-level citation links back to the original source.

This fork extends Dify `1.17.0` with an end-to-end multimodal ingestion pipeline that:
- **Builds Custom PPTX Parsing from Scratch**: Implements a dedicated slide parser (`pptx_extractor.py`), registering `.pptx` into Dify's document whitelist and extraction dispatch pipeline.
- **Intelligently Triages Visuals**: Uses Vision-Language Models to distinguish decorative graphics (logos, watermarks) from informative figures.
- **Transcribes to Structured Markdown**: Translates unstructured slides and pages into uniform, semantic Markdown while strictly preserving image preview anchors (`![image](/files/<id>/file-preview)`) and citation markers (`<!-- page: N -->`).
- **Preserves Hierarchical Context**: Employs a structure-aware splitter that injects breadcrumb chapter paths (`【Chapter > Section】`) into each chunk.
- **Hardens Vector Store Integrity**: Resolves an upstream Weaviate vector database issue to guarantee that deleted documents do not leave orphaned vector residuals.

```
+-------------------------------------------------------------------------------------------------+
|                                    Document Ingestion Pipeline                                  |
|                                                                                                 |
|  +----------------+      +---------------------+      +--------------------------------------+  |
|  | Uploaded PPTX  | ---> | Custom PptxExtractor| ---> | First-Pass Layout & Shape Extraction |  |
|  | (Zero Upstream |      | (Two-Pass Pipeline) |      | (Collects Text, Tables & Image Meta) |  |
|  |  Support)      |      +---------------------+      +--------------------------------------+  |
|  +----------------+                                                      |                      |
|                                                                          v                      |
|                          +-------------------------------------------------------------------+  |
|                          |   VLM Visual Triage (decorative_image_judge.py)                   |  |
|                          |   Classifies images in context (decorative vs. informative)       |  |
|                          +-------------------------------------------------------------------+  |
|                                       /                                 \                       |
|                        [Decorative: Logo/Watermark]               [Informative Figure]          |
|                                     /                                     \                     |
|                                    v                                       v                    |
|                             Discard & Purge                        Persist to Storage Layer     |
|                                                                    Generate Preview Anchor      |
|                                                                            |                    |
|                                                                            v                    |
|  +-------------------------------------------------------------------------------------------+  |
|  |   Multimodal Document Transcription (transcribe_service.py)                               |  |
|  |   VLM translates document to structured Markdown; preserves ![image](...) & <!-- page: N -->|  |
|  +-------------------------------------------------------------------------------------------+  |
|                                              |                                                  |
|                                              v                                                  |
|  +-------------------------------------------------------------------------------------------+  |
|  |   Structure-Aware Splitter (fixed_text_splitter.py)                                       |  |
|  |   Injects ATX Heading Breadcrumbs (【Chapter > Section】) & Merges Split Fragments        |  |
|  +-------------------------------------------------------------------------------------------+  |
|                                              |                                                  |
|                                              v                                                  |
|                       Dify Index Processors & Vector Storage (Weaviate)                         |
+-------------------------------------------------------------------------------------------------+
```

---

## Architectural Evolution

The architecture evolved through four deliberate engineering milestones driven by production requirements:

### Phase 1: Zero-to-One — Building Custom PPTX Parsing for Dify
**Upstream Dify 1.17.0 completely lacked PowerPoint (`.pptx`) support.** Neither the file extension whitelist (`api/constants/__init__.py`) nor the extraction dispatcher (`extract_processor.py`) recognized `.pptx` files.
- We authored a custom PowerPoint extraction engine from scratch (`api/core/rag/extractor/pptx_extractor.py`, initially ~200 LOC) based on `python-pptx`.
- Unlike basic text extractors, it systematically traversed slide shapes, preserved table cell geometries, extracted presenter speaker notes, cataloged embedded image relationships, and yielded slide-level `Document` objects with 0-indexed page metadata aligned with `PdfExtractor`.
- Hooked the custom extractor cleanly into Dify's document processing pipeline, making PowerPoint presentations fully ingestible for the first time.

### Phase 2: Visual Noise Discovery & VLM-Driven Triage
Once `.pptx` documents could be ingested into Dify, real-world enterprise slide decks immediately exposed a critical bottleneck: recurring corporate logos, header graphics, and watermark ornaments appeared across dozens of slides.
- Traditional heuristics (e.g. counting image cross-slide frequency) failed on single-occurrence title logos or dynamic header variations.
- To solve this, we upgraded the custom parser into a **two-pass visual triage pipeline** (`decorative_image_judge.py`):
  1. **Pass 1**: Extracts layout geometry and records image binary hashes, occurrence frequencies, and surrounding text excerpts without immediately committing images to disk.
  2. **Pass 2**: Submits candidate images along with their surrounding textual context and slide locations to a Vision-Language Model in controlled batches. The model classifies images into `decorative` vs. `meaningful`.
  3. **Outcome**: Decorative images are discarded before disk or vector store commitment, while meaningful figures are persisted to Dify's storage layer and assigned standard preview URLs.

### Phase 3: The Leap to Full Multimodal Document Transcription
The success of using a VLM for slide visual triage sparked the key architectural innovation: **Why stop at filtering images when vision models can comprehend entire slide layouts?**
- Enterprise presentations and complex office documents rely heavily on non-linear reading orders, multi-column blocks, callout boxes, and nested tables that raw rule-based extractors scramble.
- Taking inspiration from the VLM's spatial awareness, we developed the multimodal transcription service (`api/core/rag/extractor/transcribe_service.py`).
- The entire document stream is converted into high-fidelity, structured Markdown by a vision LLM, retaining rich layout relationships, tables, and formula notations.
- **Strict Invariants**:
  - Image preview anchors (`![image](/files/<id>/file-preview)`) must be preserved verbatim.
  - Page tracking comments (`<!-- page: N -->`) and table row tracking (`<!-- row: N -->`) are inserted to preserve citation integrity during vector retrieval.
  - Documents are chunked across line boundaries respecting `TRANSCRIBE_MAX_CHARS_PER_CHUNK` budgets to preserve cross-page continuity.

### Phase 4: Production Resilience & Retrieval Optimization
- **Fail-Open Resilience**: If the vision model times out, encounters rate limits, or returns malformed outputs, the system gracefully falls back to the native extractor output. Document ingestion is never blocked.
- **Structure-Aware Chunking (`fixed_text_splitter.py`)**: Enhanced Dify's fixed-size text splitter with an ATX heading stack tracker. It injects chapter breadcrumbs (e.g., `【Chapter 2 > Section 1.3】`) into each chunk and merges sub-chunks to avoid fragmented, context-free snippets.
- **Weaviate Vector Persistence Fix (`weaviate_vector.py`)**: Identified and fixed an upstream defect where document deletions in Dify failed to remove corresponding Weaviate vectors due to mismatched ID hashing algorithms, eliminating vector leaks.

---

## Key Features

- **Custom PPTX Support for Dify**: Bridges the gap in upstream Dify by adding full `.pptx` parsing, table extraction, and slide note ingestion.
- **Multimodal Document Transcription**: Automatically turns slides, PDF pages, and complex documents into standard Markdown with restored table formatting and LaTeX math expressions.
- **Context-Aware Visual Filtering**: Removes logos, decorative bars, and background watermarks before vectorization.
- **Preserved Citation Traceability**: Injects non-intrusive page tags (`<!-- page: N -->`) and heading paths (`【Heading Path】`) so that retrieval citations pinpoint the exact slide, section, and page.
- **Non-Breaking Architecture**: Placed cleanly at the output of `extract_processor.py`. Seamlessly interfaces with all Dify index processors (`paragraph`, `hierarchical`, `qa`).
- **Production-Ready Fallback**: Strict try/catch sandboxing ensures zero ingestion interruptions even if downstream VLM APIs become unavailable.

---

## Configuration Reference

Add the following environment variables to your deployment environment (e.g., `docker/.env`):

```bash
# General Toggle
TRANSCRIBE_ENABLED=true

# Vision Model Configuration
TRANSCRIBE_MODEL_PROVIDER=langgenius/openai_api_compatible/openai_api_compatible
TRANSCRIBE_MODEL_NAME=qwen3.8-27b
TRANSCRIBE_IMAGE_DETAIL=low

# Visual Filtering Toggle (true = VLM classification, false = frequency heuristic)
TRANSCRIBE_DECORATIVE_JUDGE=true

# Formatting & Prompt Customization
TRANSCRIBE_LANGUAGE=中文
TRANSCRIBE_PROMPT=
TRANSCRIBE_TEMPERATURE=

# Chunking & Pagination Budget
TRANSCRIBE_MAX_CHARS_PER_CHUNK=12000
TRANSCRIBE_MAX_CHUNKS=0
TRANSCRIBE_PAGE_OFFSET=1
```

### Parameter Details

| Variable | Default | Description |
|---|---|---|
| `TRANSCRIBE_ENABLED` | `false` | Master toggle. When `false`, documents bypass VLM and use native extraction. |
| `TRANSCRIBE_MODEL_PROVIDER` | `""` | The Dify model provider ID configured in the admin console. |
| `TRANSCRIBE_MODEL_NAME` | `""` | The vision model identifier supported under the selected provider. |
| `TRANSCRIBE_DECORATIVE_JUDGE` | `true` | Enables VLM-based visual filtering for decorative images. If disabled, falls back to cross-slide frequency heuristics. |
| `TRANSCRIBE_IMAGE_DETAIL` | `low` | Detail tier for images passed to the vision model (`low` or `high`). |
| `TRANSCRIBE_LANGUAGE` | `""` | Desired output language for transcription (e.g. `中文`, `English`). Defaults to document language if unset. |
| `TRANSCRIBE_MAX_CHARS_PER_CHUNK` | `12000` | Character threshold for splitting long documents during transcription. |
| `TRANSCRIBE_PAGE_OFFSET` | `1` | Starting page index for generated `<!-- page: n -->` citation tags. |

---

## Quickstart & Verification

### 1. Apply Configuration
After setting the environment variables in `docker/.env`, rebuild and restart the `api` and `worker` services:

```bash
docker compose build api worker
docker compose up -d api worker
```

### 2. Verify Service Settings
Ensure the configuration has taken effect inside the container:

```bash
docker compose exec api python -c "from configs import dify_config; print('Transcribe Enabled:', dify_config.TRANSCRIBE_ENABLED)"
```

### 3. Run Unit Tests
Execute the structure-aware text splitter test suite:

```bash
docker compose exec api uv run pytest tests/unit_tests/core/rag/splitter/test_text_splitter.py -q
```

### 4. End-to-End Validation
1. Upload a multi-slide `.pptx` containing charts, tables, and corporate logos into a Dify Knowledge Base dataset.
2. Inspect the API service logs:
   ```bash
   docker compose logs -f api | grep -E "transcribe|decorative"
   ```
3. Once indexing finishes, check the document preview in Dify:
   - Formatted Markdown with structured tables.
   - Chapter breadcrumbs (e.g., `【Overview > Architecture】`) at the head of chunks.
   - Preserved image preview URLs for informative figures; corporate logos and decorative graphics cleanly eliminated.

---

## Repository Structure & Commit Log

The commits layered on top of Dify `1.17.0` (`09a855dcef`) represent clean, modular feature increments:

```
09a855dcef chore: bump version to 1.17.0 (Baseline upstream release)
b100a01822 feat(rag): introduce custom PPTX extractor and register entrypoint in Dify
1fbc2d586f feat(rag): wire transcription entrypoint and converge extension whitelist
72afef16d4 fix(rag): restore pptx dispatch branch with explicit file assertions
f4d3da069b feat(rag): implement multimodal transcription service & heading-aware splitter
7d3b3057fc feat(rag): upgrade decorative image filtering to VLM context classification
7204e8fa3b fix(vdb): resolve Weaviate vector deletion defect to prevent orphan embeddings
```

---

## Upstream Relationship & License

This repository is maintained as a feature-enhanced fork of [Dify](https://github.com/langgenius/dify) `v1.17.0`. All enhancements maintain strict backward compatibility with upstream APIs.

Original Dify documentation is preserved in [`README.upstream.md`](./README.upstream.md). Distributed under the original project license (see [`LICENSE`](./LICENSE)).

# Dify Multimodal RAG & Intelligent Document Transcription Extension

[English](README.md) | [中文](README_zh.md)

An enterprise-grade multimodal extension based on **Dify 1.17.0**, introducing native presentation parsing, vision-language model (VLM) visual triage, end-to-end multimodal document transcription, and structure-aware chunking.

---

## Overview

In enterprise Retrieval-Augmented Generation (RAG) systems, slide presentations (`.pptx`) and complex multi-page documents present severe structural challenges:
1. **Loss of Spatial & Visual Context**: Conventional document parsers flatten slides into linear text streams, destroying multi-column layouts, tabular relationships, and visual notes.
2. **Visual Noise Contamination**: Real-world slide decks are inundated with repetitive decorative graphics (corporate logos, recurring header/footer elements, watermark icons, and cosmetic dividers). If ingested unchecked, these images waste vector database storage, pollute multi-modal indices, and degrade retrieval precision.
3. **Traceability Degradation**: Standard chunkers split content arbitrarily across boundaries, stripping chunk context and breaking page-level citation links back to the original source.

This fork extends Dify `1.17.0` with an end-to-end multimodal ingestion pipeline that:
- **Intelligently Triages Visuals**: Uses Vision-Language Models to distinguish decorative graphics from informative figures.
- **Transcribes to Structured Markdown**: Translates unstructured slides and pages into uniform, semantic Markdown while strictly preserving image preview anchors (`![image](/files/<id>/file-preview)`) and citation markers (`<!-- page: N -->`).
- **Preserves Hierarchical Context**: Employs a structure-aware splitter that injects breadcrumb chapter paths (`【Chapter > Section】`) into each chunk.
- **Hardens Vector Store Integrity**: Resolves an upstream Weaviate vector database issue to guarantee that deleted documents do not leave orphaned vector residuals.

```
+-------------------------------------------------------------------------------------------------+
|                                    Document Ingestion Pipeline                                  |
|                                                                                                 |
|  +----------------+      +---------------------+      +--------------------------------------+  |
|  | Uploaded PPTX  | ---> |   PptxExtractor     | ---> | First-Pass Layout & Shape Extraction |  |
|  | / PDF Document |      | (Two-Pass Pipeline) |      | (Collects Text, Tables & Image Meta) |  |
|  +----------------+      +---------------------+      +--------------------------------------+  |
|                                                                  |                              |
|                                                                  v                              |
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

The architecture evolved through deliberate engineering iterations driven by production requirements:

### Phase 1: Native Structural Parsing for PPTX
The project initially introduced a focused native parser (`api/core/rag/extractor/pptx_extractor.py`, ~200 LOC) based on `python-pptx`. Unlike naive text dumpers, it traversed slide shapes, preserved table layouts, extracted presenter notes, and yielded slide-level `Document` objects with 0-indexed page metadata aligned with `PdfExtractor`.

### Phase 2: Visual Noise Discovery & VLM-Driven Triage
Real-world enterprise slide decks revealed a critical challenge: recurring corporate logos, header graphics, and watermark ornaments appeared on nearly every slide. 
- Traditional heuristic rules (e.g., counting image cross-slide frequency) failed on single-occurrence title logos or dynamic header variations.
- To address this, we developed an intelligent visual triage pipeline (`decorative_image_judge.py`). The extractor runs in two passes:
  1. **Pass 1**: Extracts layout geometry and records image binary hashes, occurrence frequencies, and surrounding text excerpts without immediately committing images to disk.
  2. **Pass 2**: Submits candidate images along with their surrounding textual context and slide locations to a Vision-Language Model in controlled batches. The model classifies images into `decorative` vs. `meaningful`.
  3. **Outcome**: Decorative images are discarded before disk or vector store commitment, while meaningful figures are persisted to Dify's storage layer and assigned standard preview URLs.

### Phase 3: The Leap to Full Multimodal Document Transcription
Leveraging the success of VLM-based visual evaluation, the system expanded from image filtering to full-document multimodal transcription (`api/core/rag/extractor/transcribe_service.py`).
- Complex office documents and slides frequently have non-linear reading orders, multi-column blocks, callout boxes, and nested tables that raw rule-based extractors scramble.
- By piping the extracted text and preserved image preview URLs into a high-capacity vision model, the entire document is transcribed into clean, unified Markdown.
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
b100a01822 feat(rag): introduce native PPTX extractor with transcription hook
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

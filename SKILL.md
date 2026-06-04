---
name: document-intelligence-skill
description: PDF-first local document parsing and extraction skill. Use when Codex needs to compare PDF text extraction quality, diagnose `(cid:...)`, blank, garbled, scanned, or low-quality PDF output, decide whether OCR or a cleaner source is needed, convert local documents to Markdown/TXT/JSON, extract PDF tables or layout coordinates, render PDF pages, run optional Tesseract OCR, chunk extracted text for indexing/RAG, classify documents, build traceable knowledge packs, or extract/verify/export invoice fields. Also supports DOCX, PPTX, XLSX, HTML, text-like files, images, audio, EPub, and ZIP through the bundled parser script when the request is about parsing, conversion, quality gates, or structured extraction.
---

# Document Intelligence Skill

## Core Rule

Prefer `scripts/parse_document_compare.py` for deterministic work. It writes outputs beside the source file or to an explicit output directory and should not modify source documents.

Use this skill as `$document-intelligence-skill`. Claude-style `/pdf-compare` files are legacy compatibility only.

## Default Workflow

1. Confirm the input path exists and identify the extension.
2. If dependency state is unclear, run `doctor` before parsing:
   `python scripts\parse_document_compare.py doctor --format pdf --ocr --opendataloader --json`
3. For PDF quality diagnosis, start with a small sample:
   `python scripts\parse_document_compare.py compare "<file.pdf>" --max-pages 3 --output-format md`
4. Read `compare_report.md` first. Use its recommendation, quality score, diagnostics, and previews before inspecting individual parser outputs.
5. If the sample is good and the user needs full output, rerun with the needed command and scope. Use `--max-pages 0` only when full-document parsing is actually needed.
6. If multiple parsers produce `(cid:...)`, control characters, mojibake, near-empty text, or a bad/empty quality label, say ordinary text-layer extraction is insufficient and recommend OCR or a cleaner source file.

## Command Selection

- User asks "just parse this document", "自动解析", or wants a best end-to-end result: use `auto`.
- User asks which parser is best, whether OCR is needed, or why extraction is bad: use `compare`, then read `compare_report.md`.
- User asks for Markdown/TXT/JSON: use `convert`.
- User asks to process a folder: use `scan-dir` first for quality, then `batch` or `batch-knowledge`.
- User asks for tables: use `tables` for PDFs.
- User asks for metadata, page count, TOC, or text-layer hints: use `metadata`.
- User asks for RAG/indexing chunks: use `chunk`; use `--chunk-by page` when page traceability matters.
- User asks to inspect pages visually or prepare OCR inputs: use `render-pages`.
- User explicitly asks for OCR or extraction is clearly image/scanned: check dependencies, then use `ocr` or `auto --auto-ocr`.
- User asks for coordinates, page maps, or source traceability: use `layout-json` or `knowledge-pack`.
- User asks questions over an existing knowledge pack, chunks file, or parsed document: use `qa`.
- User asks to compare two document versions: use `diff-docs`.
- User asks for invoice fields or invoice summaries: use `extract-fields --profile invoice`, `verify-fields --profile invoice`, or `export-xlsx --profile invoice`.
- User asks what type of document it is or which strategy to use: use `classify`.

## Quality And OCR Rules

- Treat `bad` and `empty` quality labels as failed ordinary extraction.
- Use `--min-quality` and `--fail-on-bad` for batch or CI-style gates.
- `--ocr-fallback` only records fallback intent in reports/metadata. It does not run OCR.
- Real OCR requires Python packages plus a system `tesseract` executable on PATH and the requested language data.
- Do not install missing dependencies automatically. Report skipped parsers and suggest targeted install commands only when useful.
- For `opendataloader-pdf`, check both the Python package and Java availability.

## References

- Read `references/command-guide.md` when exact command syntax, supported formats, outputs, or examples are needed.
- Read `references/dependencies.md` only when diagnosing missing parser/OCR dependencies or preparing install instructions.

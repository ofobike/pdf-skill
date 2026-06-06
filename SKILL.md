---
name: pdf-parse-skill
description: PDF-first local document parsing and extraction skill. Use when Codex needs to compare PDF text extraction quality, diagnose `(cid:...)`, blank, garbled, scanned, or low-quality PDF output, decide whether OCR or a cleaner source is needed, convert local documents to Markdown/TXT/JSON, run parser voting with OCR, extract/vote PDF tables, extract layout coordinates, render PDF pages, run optional Tesseract OCR, chunk extracted text for indexing/RAG, classify documents, build customer packs or traceable knowledge packs, or extract/verify/export invoice fields with field-level confidence. Also supports DOCX, PPTX, XLSX, HTML, text-like files, images, audio, EPub, and ZIP through the bundled parser script when the request is about parsing, conversion, quality gates, or structured extraction.
---

# PDF Parse Skill

## Core Rule

Prefer `scripts/parse_document_compare.py` for deterministic work. It writes outputs beside the source file or to an explicit output directory and should not modify source documents.

Use this skill as `$pdf-parse-skill`. Claude-style `/pdf-compare` references are legacy compatibility only.

## Default Workflow

1. Confirm the input path exists and identify the extension.
2. If dependency state is unclear, run `doctor` before parsing:
   `python scripts\parse_document_compare.py doctor --format pdf --ocr --opendataloader --json`
3. If installed parsers may need runtime models, external commands, or network-backed first-run setup, run `probe` on the actual file:
   `python scripts\parse_document_compare.py probe "<file.pdf>" --max-pages 1`
4. For PDF quality diagnosis, start with a small sample:
   `python scripts\parse_document_compare.py compare "<file.pdf>" --max-pages 3 --output-format md`
5. Read `compare_report.md` first. Use its recommendation, quality score, diagnostics, and previews before inspecting individual parser outputs.
6. If the sample is good and the user needs full output, rerun with the needed command and scope. Use `--max-pages 0` only when full-document parsing is actually needed.
7. If multiple parsers produce `(cid:...)`, control characters, mojibake, near-empty text, or a bad/empty quality label, say ordinary text-layer extraction is insufficient and recommend OCR or a cleaner source file.

## Command Selection

- User asks "just parse this document", "自动解析", or wants a best end-to-end result: use `auto`.
- User asks whether installed parsers can really run on a file, or dependencies were just installed: use `probe`, then read `probe_report.md`.
- User asks which parser is best, whether OCR is needed, or why extraction is bad: use `compare`, then read `compare_report.md`.
- User asks to deliver the best parsed PDF text to a customer/downstream system after multiple parsers agree: use `vote --probe-before-vote`, then provide `best.md/txt/json`, `vote_report.md/json`, and the preflight `probe_report.md/json`.
- User asks to include OCR in parser voting: include `ocr-tesseract` in `--parsers` for `probe`, `vote`, or `customer-pack`; do not treat OCR only as a last fallback.
- User asks for a customer-ready package for a complex PDF with tables, company data, layout, or source traceability: use `customer-pack`; provide `manifest.json`, `README.md`, `best.md/txt/json`, `tables/table_vote_report.md/json`, `tables/best_tables.md/csv/json`, `layout.json`, `metadata.json`, `vote_report.md/json`, and preflight `probe_report.md/json`.
- User asks to process a directory of customer PDFs: use `batch-customer-pack`; provide `index.json/index.md` and the per-file package paths.
- User asks for the most reliable invoice PDF result for a customer: use `vote --probe-before-vote --profile invoice --customer --format all`; provide `customer_best.md/txt/json`, `best.md/txt/json`, `vote_report.md/json`, and `preflight_probe/probe_report.md/json`.
- User asks for PDF-to-Markdown optimized for LLM/RAG: prefer `convert --parser pymupdf4llm`; try `docling` for heavier structure-aware conversion; try `pspdfkit` only when the external `pdf-to-markdown` CLI is installed.
- User asks for Markdown/TXT/JSON: use `convert`.
- User asks to process a folder: use `scan-dir` first for quality, then `batch`, `batch-knowledge`, or `batch-customer-pack` depending on the requested deliverable.
- User asks for simple tables: use `tables` for PDFs.
- User asks for complex, cross-page, merged-cell, or borderless tables, or asks which table extraction is best: use `table-vote`.
- User asks for metadata, page count, TOC, or text-layer hints: use `metadata`.
- User asks for RAG/indexing chunks: use `chunk`; use `--chunk-by page` when page traceability matters.
- User asks to inspect pages visually or prepare OCR inputs: use `render-pages`.
- User explicitly asks for OCR or extraction is clearly image/scanned: check dependencies, then use `ocr` or `auto --auto-ocr`.
- User asks for coordinates, page maps, or source traceability: use `layout-json` or `knowledge-pack`.
- User asks questions over an existing knowledge pack, chunks file, or parsed document: use `qa`.
- User asks to compare two document versions: use `diff-docs`.
- User asks for invoice fields or invoice summaries: use `extract-fields --profile invoice`, `verify-fields --profile invoice`, or `export-xlsx --profile invoice`.
- User asks what type of document it is or which strategy to use: use `classify`; profiles include invoice, contract, bank statement, quotation, purchase order, report, annual report, textbook, notice, scanned PDF, and generic.

## Quality And OCR Rules

- Treat `bad` and `empty` quality labels as failed ordinary extraction.
- Use `probe` after installing heavy parsers such as `docling`; import availability does not guarantee runtime model downloads or external commands will work.
- In `vote`, repeated lines/spans reduce the score so duplicated parser output does not win only by text length.
- In `vote --probe-before-vote`, run a real-file sample probe first and only send `ready` parsers into the final vote.
- In `vote --profile invoice`, invoice field completeness, validation status, amount/tax checks, and duplicate line-item checks are included in the score.
- In `vote --profile invoice --customer`, customer outputs include field-level confidence with source parser, support count, estimated page, bbox placeholder, and confidence.
- `ocr-tesseract` is a PDF parser candidate for voting and probing when OCR dependencies are available.
- `table-vote` scores table density, row/column consistency, header completeness, coverage, and cross-method consensus.
- Use `--timeout` for heavy parsers such as `docling`, OCR, or external CLIs when the user cares about finishing reliably.
- Use `--parser-health-cache` when repeated runs should skip parsers that recently failed or timed out for the same file/page scope.
- Use `--min-quality` and `--fail-on-bad` for batch or CI-style gates.
- `--ocr-fallback` only records fallback intent in reports/metadata. It does not run OCR.
- Real OCR requires Python packages plus a system `tesseract` executable on PATH and the requested language data.
- Do not install missing dependencies automatically. Report skipped parsers and suggest targeted install commands only when useful.
- For `opendataloader-pdf`, check both the Python package and Java availability.
- For `pspdfkit`, check the external `pdf-to-markdown` CLI with `doctor`; do not treat it as a Python package.

## References

- Read `references/command-guide.md` when exact command syntax, supported formats, outputs, or examples are needed.
- Read `references/dependencies.md` only when diagnosing missing parser/OCR dependencies or preparing install instructions.

---
name: pdf-parse-skill
description: PDF-first local document parsing and extraction skill. Use when Codex needs to compare PDF text extraction quality, diagnose `(cid:...)`, blank, garbled, scanned, or low-quality PDF output, decide whether OCR or a cleaner source is needed, convert local documents to Markdown/TXT/JSON, run parser voting with OCR, extract/vote PDF tables, extract layout coordinates, render PDF pages, run Tesseract OCR, chunk text for indexing/RAG, classify documents, build customer packs or traceable knowledge packs, run golden PDF regression evaluation, cache parser results, generate customer JSON Schema contracts, crop field evidence images from bbox coordinates, generate HTML review pages, or extract/verify/export invoice and business-profile fields with confidence, parser fusion, and optional bbox traceability. Also supports DOCX, PPTX, XLSX, HTML, text-like files, images, audio, EPub, and ZIP for parsing, conversion, quality gates, or structured extraction.
---

# PDF Parse Skill

## Core Rule

Prefer `scripts/parse_document_compare.py` for deterministic work. It writes outputs beside the source file or to an explicit output directory and should not modify source documents.

Use this skill as `$pdf-parse-skill`. Claude-style `/pdf-compare` references are legacy compatibility only.

## Natural Language Triggers

Treat these Chinese requests as direct triggers for this skill:

- “这批 PDF 很适合做 Golden 样本库 / 回归样本库”：use `init-golden`; add `--include-ocr` when the batch may contain scanned PDFs.
- “跑一下 Golden 回归 / 检查解析策略有没有退步”：use `eval-golden`, usually with `--recursive --timeout 120 --parser-health-cache --result-cache`.
- “给客户一份最可靠的 PDF 转 Markdown/TXT/JSON 结果”：use `vote --probe-before-vote --format all`; for invoices add `--profile invoice --customer`.
- “生成客户交付包 / 字段证据 / HTML 审阅页 / schema”：use `customer-pack`.
- “这个 PDF 表格复杂 / 跨页表格 / 无边框表格”：use `table-vote`, or `customer-pack` when a full deliverable is requested.
- “解析很慢 / docling 或 OCR 不要重复跑”：use `--timeout --parser-health-cache --result-cache`.
- “这不是发票，是考试资料/题库/报告，只要检查文本质量”：for Golden cases prefer generic expectations instead of invoice field checks.

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
- User asks for a customer-ready package for a complex PDF with tables, company data, layout, source traceability, field evidence, schema, or visual review: use `customer-pack`; provide `manifest.json`, `README.md`, `best.md/txt/json`, `customer_best.md/txt/json`, `customer_best.schema.json`, `field_evidence/*`, `review.html`, `tables/table_vote_report.md/json`, `tables/best_tables.md/csv/json`, `layout.json`, `metadata.json`, `vote_report.md/json`, and preflight `probe_report.md/json`.
- User asks to process a directory of customer PDFs: use `batch-customer-pack`; provide `index.json/index.md` and the per-file package paths.
- User asks for the most reliable invoice PDF result for a customer: use `vote --probe-before-vote --profile invoice --customer --format all`; field-level fusion is enabled by default, and add `--field-layout` when field page/bbox traceability is requested. Add `--field-evidence` when the user asks for field screenshots/evidence images, and `--review-html` when the user asks for a visual review page. Provide `customer_best.md/txt/json`, `customer_best.schema.json`, `best.md/txt/json`, `vote_report.md/json`, optional `field_layout.json`, optional `field_evidence/*`, optional `review.html`, and `preflight_probe/probe_report.md/json`.
- User asks for PDF-to-Markdown optimized for LLM/RAG: prefer `convert --parser pymupdf4llm`; try `docling` for heavier structure-aware conversion; try `pspdfkit` only when the external `pdf-to-markdown` CLI is installed.
- User asks for Markdown/TXT/JSON: use `convert`.
- User asks to process a folder: use `scan-dir` first for quality, then `batch`, `batch-knowledge`, or `batch-customer-pack` depending on the requested deliverable.
- User asks for simple tables: use `tables` for PDFs.
- User asks for complex, cross-page, merged-cell, or borderless tables, or asks which table extraction is best: use `table-vote`.
- User asks for metadata, page count, TOC, or text-layer hints: use `metadata`.
- User asks for RAG/indexing chunks: use `chunk`; use `--chunk-by page` when page traceability matters.
- User asks to inspect pages visually or prepare OCR inputs: use `render-pages`.
- User explicitly asks for OCR or extraction is clearly image/scanned: check dependencies, then use `ocr` or `auto --auto-ocr`.
- User asks for coordinates, page maps, source traceability, field evidence images, or a visual review page: use `layout-json`, `knowledge-pack`, or `customer-pack`; for field-level bbox use `vote --customer --field-layout`, add `--field-evidence` or `--review-html` as requested, or use `customer-pack`.
- User asks to turn a batch of PDFs into a Golden sample library, regression set, or stable parser acceptance set: use `init-golden` to create case JSON/index/README, then use `eval-golden` for regression.
- User asks whether parser changes regressed, wants to run an existing golden set, regression testing, or stable customer acceptance checks: use `eval-golden` with case JSON files.
- User asks questions over an existing knowledge pack, chunks file, or parsed document: use `qa`.
- User asks to compare two document versions: use `diff-docs`.
- User asks for invoice fields or invoice summaries: use `extract-fields --profile invoice`, `verify-fields --profile invoice`, or `export-xlsx --profile invoice`.
- User asks what type of document it is or which strategy to use: use `classify`; profiles include invoice, contract, bank statement, quotation, purchase order, report, annual report, textbook, notice, scanned PDF, and generic. If the user asks to extract business fields for contract/bank statement/quotation/purchase order/report/annual report, use `vote --profile <profile> --customer`.

## Quality And OCR Rules

- Treat `bad` and `empty` quality labels as failed ordinary extraction.
- Use `probe` after installing heavy parsers such as `docling`; import availability does not guarantee runtime model downloads or external commands will work.
- In `vote`, repeated lines/spans reduce the score so duplicated parser output does not win only by text length.
- In `vote --probe-before-vote`, run a real-file sample probe first and only send `ready` parsers into the final vote.
- In `vote --profile invoice`, invoice field completeness, validation status, amount/tax checks, and duplicate line-item checks are included in the score.
- In `vote --profile invoice --customer`, customer outputs use field-level parser fusion by default: each invoice field can come from the strongest parser candidate, then invoice validation is rerun. If fusion would make validation worse than the whole-document winner, it is not used. Use `--no-field-fusion` to disable it.
- Customer invoice outputs include field-level confidence with source parser, support count, page, bbox, source location, and confidence. Bbox is filled when `--field-layout` or `customer-pack` layout matching can locate the value; otherwise it remains `null`.
- Customer outputs include `schema_version` and `customer_best.schema.json`; treat this as the downstream contract for `customer_best.json`.
- `customer-pack` generates field evidence images and `review.html` by default. Disable only when the user asks for a lighter package or when rendering/cropping dependencies are unavailable.
- `vote --customer` generates field evidence only when `--field-evidence` is passed, and review HTML only when `--review-html` is passed. These need layout extraction; include `--field-layout` for reliable bbox evidence.
- For contract, bank statement, quotation, purchase order, report, and annual report profiles, the script performs lightweight structured extraction into `customer_best.fields` and field confidence; do not claim the same depth as invoice validation.
- `ocr-tesseract` is a PDF parser candidate for voting and probing when OCR dependencies are available.
- `table-vote` scores table density, row/column consistency, header completeness, coverage, and cross-method consensus.
- Use `--timeout` for heavy parsers such as `docling`, OCR, or external CLIs when the user cares about finishing reliably.
- Use `--parser-health-cache` when repeated runs should skip parsers that recently failed or timed out for the same file/page scope.
- Use `--result-cache` when repeated runs should avoid rerunning slow successful parsers for the same file/page scope. Use `--result-cache-dir` to share cache across output directories.
- For non-invoice Golden sets such as exams, textbooks, reports, or question banks, prefer generic expectations: `min_vote_score`, `quality_gate_passed`, `min_non_space_chars`, `quality_label_not`, `text_contains_any/all`, and repetition thresholds. Avoid pinning `winner_parser` unless a specific parser must remain the winner.
- Use `--min-quality` and `--fail-on-bad` for batch or CI-style gates.
- `--ocr-fallback` only records fallback intent in reports/metadata. It does not run OCR.
- Real OCR requires Python packages plus a system `tesseract` executable on PATH and the requested language data.
- Do not install missing dependencies automatically. Report skipped parsers and suggest targeted install commands only when useful.
- For `opendataloader-pdf`, check both the Python package and Java availability.
- For `pspdfkit`, check the external `pdf-to-markdown` CLI with `doctor`; do not treat it as a Python package.

## References

- Read `references/command-guide.md` when exact command syntax, supported formats, outputs, or examples are needed.
- Read `references/dependencies.md` only when diagnosing missing parser/OCR dependencies or preparing install instructions.

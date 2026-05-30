---
name: document-parse-compare
description: Parse, convert, batch-convert, scan, auto-process, classify, package, and compare text extraction quality for local documents. Use when Codex needs to extract or compare content from PDF, DOCX, PPTX, XLSX, or HTML files; diagnose `(cid:...)`, blank, garbled, or low-quality extraction; choose among markitdown and format-specific parsers; enforce quality gates; render PDF page screenshots; output PDF layout JSON; chunk text by character or page; extract PDF tables; extract and verify structured invoice fields; export invoice summaries to XLSX; run optional Tesseract OCR or auto OCR fallback; classify document type; generate traceable RAG/knowledge packs; or decide whether OCR or a better source file is needed.
---

# Document Parse Compare

## Triggering

Explicit Codex invocation:

```text
$document-parse-compare
```

Use this skill explicitly when the user says things like:

- `Use $document-parse-compare to compare this PDF parser output.`
- `Use $document-parse-compare to convert this DOCX to JSON.`
- `Use $document-parse-compare to extract PDF tables as CSV and JSON.`
- `Use $document-parse-compare to scan a directory and fail if PDF quality is bad.`
- `Use $document-parse-compare to render PDF pages and OCR page 1.`
- `Use $document-parse-compare to auto-process this PDF and extract invoice fields.`
- `Use $document-parse-compare to export these invoices to XLSX.`
- `Use $document-parse-compare to classify this document and recommend a strategy.`
- `Use $document-parse-compare to build a knowledge pack for this PDF.`

The skill should also be used implicitly for requests such as:

- Compare parser quality for a PDF, DOCX, PPTX, XLSX, or HTML file.
- Convert a document to Markdown, TXT, or JSON.
- Diagnose `(cid:...)`, blank, garbled, or suspicious extraction output.
- Decide whether OCR or a cleaner source document is needed.
- Extract PDF tables.
- Batch-convert a directory of supported documents.
- Check parser dependencies before a parse run.
- Extract document metadata, especially PDF page count, embedded metadata, and TOC preview.
- Split extracted text into JSONL/JSON/Markdown/TXT chunks for downstream LLM or indexing workflows.
- Split PDFs by page with `chunk --chunk-by page`.
- Batch-scan directories for extraction quality.
- Enforce extraction quality gates with `--min-quality` and `--fail-on-bad`.
- Render selected PDF pages as PNG screenshots.
- Run real OCR through the optional `ocr` command when ordinary extraction is poor.
- Auto-process one document with metadata, best text, chunks, classification, and fields.
- Extract structured invoice fields.
- Export invoice fields to XLSX for one file or a directory.
- Run auto OCR fallback when ordinary extraction is poor.
- Output PDF layout JSON with coordinates.
- Verify extracted invoice fields.
- Classify documents and choose a processing strategy.
- Generate traceable knowledge packs for one file or a directory.

Claude-style `/pdf-compare` is legacy compatibility only. Codex triggering uses `$document-parse-compare` and the frontmatter description above.

## Use The Bundled Script

Prefer `scripts/parse_document_compare.py` for deterministic document parsing work. It supports:

- `compare`: run multiple parsers against one file and create Markdown/JSON reports plus parser outputs.
- `convert`: convert one file to `md`, `txt`, `json`, or all three using a selected parser.
- `batch`: convert matching files in a directory.
- `tables`: extract PDF tables to `md`, `csv`, `json`, or all three.
- `doctor`: check parser dependencies and print install hints, optionally as JSON.
- `metadata`: extract file metadata; PDF gets page count, embedded metadata, TOC preview, and sample text-layer hints.
- `scan-dir`: scan a directory for extraction quality and write Markdown/JSON/CSV reports.
- `chunk`: extract text and split it into `jsonl`, `json`, `md`, `txt`, or all formats, by character or by PDF page.
- `render-pages`: render selected PDF pages to PNG screenshots.
- `ocr`: run Tesseract OCR on selected PDF pages and write TXT/Markdown/JSON.
- `auto`: run a smart one-file pipeline that writes metadata, best text, chunks, classification, and invoice fields when detected.
- `extract-fields`: extract structured fields; v3.6 supports `--profile invoice`.
- `export-xlsx`: batch extract invoice fields and write an XLSX summary.
- `layout-json`: output PDF page, block, line, span, font, and coordinate data.
- `verify-fields`: validate structured fields; invoice profile supports totals and strict checks.
- `classify`: detect document profile and recommended processing strategy.
- `knowledge-pack`: generate chunks, metadata, manifest, quality report, layout, source text, and page map.
- `batch-knowledge`: generate knowledge packs for a directory.

Supported inputs:

- PDF: `markitdown`, `pymupdf`, `pypdf`, `pdfplumber`, `pdfminer`, `liteparse`
- DOCX: `markitdown`, `python-docx`
- PPTX: `markitdown`, `python-pptx`
- XLSX: `markitdown`, `openpyxl`
- HTML: `markitdown`, `beautifulsoup4`

Legacy binary Office formats such as `.doc`, `.ppt`, and `.xls` may be detected, but prefer asking for modern `.docx`, `.pptx`, or `.xlsx` files when parser support is unreliable.

## Workflow

1. Confirm the file path exists and identify the extension.
2. Run `doctor --format pdf` or the relevant format when dependency state is unclear.
3. For quality diagnosis, start with `compare`.
4. For large PDFs, sample first:
   - Use `--max-pages 3` for a smoke test.
   - Use `--start-page` to skip covers, front matter, or image-heavy pages.
   - Use `--max-pages 0` only when a full parse is needed and the sample is useful.
5. Read `compare_report.md` first. Use the recommendation and quality diagnostics before inspecting individual parser outputs.
6. If multiple parsers produce `(cid:...)`, control characters, mojibake, or near-empty output, state that ordinary embedded-text extraction is insufficient and recommend OCR or a cleaner source document.
7. Use `metadata` to inspect page count, PDF metadata, TOC, and whether sampled pages have a text layer.
8. Use `scan-dir` before large batch work to find files that are likely to need OCR.
9. Use `--min-quality` and `--fail-on-bad` when a script or CI-style workflow should stop on poor extraction.
10. Use `chunk --chunk-by page` when page boundaries matter more than fixed-size text chunks.
11. Use `render-pages` to inspect visual pages or create OCR inputs.
12. Use `ocr` only when the user wants real OCR and the local OCR dependencies are installed.
13. Use `auto` when the user asks for "just parse this document" or wants the best end-to-end output without choosing every step.
14. Use `auto --auto-ocr` when poor extraction should attempt OCR automatically.
15. Use `layout-json` when coordinates, page maps, or source traceability matter.
16. Use `extract-fields --profile invoice` for electronic invoices; use `verify-fields --profile invoice` for validation and `export-xlsx --profile invoice` for invoice directories.
17. Use `classify` and `knowledge-pack` for document-intelligence and RAG/indexing workflows.
18. Do not install missing dependencies automatically. Report skipped parsers and suggest the relevant dependency only when useful.

## Commands

Run from the skill directory or pass the absolute script path.

```powershell
python scripts\parse_document_compare.py compare "D:\path\file.pdf" --max-pages 30 --output-format md
python scripts\parse_document_compare.py compare "D:\path\file.pdf" --start-page 100 --max-pages 10 --output-format all
python scripts\parse_document_compare.py convert "D:\path\file.docx" --parser markitdown --format json -o "D:\path\file.json"
python scripts\parse_document_compare.py batch "D:\docs" --ext .pdf,.docx --parser markitdown --format txt
python scripts\parse_document_compare.py scan-dir "D:\docs" --ext .pdf --max-pages 3 --min-quality 0.6 --fail-on-bad
python scripts\parse_document_compare.py tables "D:\path\file.pdf" --pages 1-5 --format all
python scripts\parse_document_compare.py doctor --format pdf --ocr --json
python scripts\parse_document_compare.py metadata "D:\path\file.pdf" --format json
python scripts\parse_document_compare.py chunk "D:\path\file.pdf" --parser pymupdf --format jsonl --chunk-size 2000 --overlap 200
python scripts\parse_document_compare.py chunk "D:\path\file.pdf" --parser pymupdf --chunk-by page --format jsonl
python scripts\parse_document_compare.py render-pages "D:\path\file.pdf" --pages 1-3 --dpi 150
python scripts\parse_document_compare.py ocr "D:\path\file.pdf" --pages 1 --lang chi_sim+eng --format md
python scripts\parse_document_compare.py auto "D:\path\file.pdf" --profile auto
python scripts\parse_document_compare.py auto "D:\path\file.pdf" --auto-ocr --ocr-pages 1-3 --layout
python scripts\parse_document_compare.py extract-fields "D:\path\invoice.pdf" --profile invoice --format json
python scripts\parse_document_compare.py verify-fields "D:\path\invoice_fields.json" --profile invoice --strict
python scripts\parse_document_compare.py export-xlsx "D:\invoices" --profile invoice --recursive -o "D:\invoices\summary.xlsx"
python scripts\parse_document_compare.py layout-json "D:\path\file.pdf" --max-pages 3 -o "D:\path\layout.json"
python scripts\parse_document_compare.py classify "D:\path\file.pdf" --max-pages 3
python scripts\parse_document_compare.py knowledge-pack "D:\path\file.pdf" --chunk-by page --out-dir "D:\path\pack"
python scripts\parse_document_compare.py batch-knowledge "D:\docs" --recursive --out-dir "D:\docs\packs"
python scripts\parse_document_compare.py compare "D:\path\file.pdf" --max-pages 3 --ocr-fallback
```

The script also accepts a file path without a subcommand as a shortcut for `compare`.

## Outputs

`compare` creates an output directory beside the input file unless `--out-dir` is provided. Generated files:

- `compare_report.md`: human-readable parser comparison, stats, similarity, and previews.
- `compare_report.json`: machine-readable report.
- `<parser>.md`, `<parser>.txt`, or `<parser>.json`: normalized extraction output for each successful parser, depending on `--output-format`.

`convert`, `batch`, and `tables` write the requested outputs without modifying source files. Markdown remains the default because it is usually best for LLM reading, but use TXT for simple search/indexing and JSON for downstream automation.

`compare_report.md` includes a recommended parser, quality score, quality label, and diagnostics for common extraction issues such as `(cid:...)`, control characters, and near-empty output.

`chunk` defaults to JSONL because it is convenient for ingestion pipelines: one chunk per line with source metadata, parser name, offsets, and text.

`scan-dir` writes `scan_report.md`, `scan_report.json`, and `scan_report.csv`. Use `needs_ocr`, `quality_score`, `quality_label`, and `gate_passed` to decide next steps.

`render-pages` writes PNG files plus `render_pages.json`.

`ocr` writes recognized text from rendered PDF pages. It is real OCR, not text-layer extraction. It requires PyMuPDF, Pillow, pytesseract, and a system Tesseract binary with the requested language data.

`auto` writes `auto_report.md`, `auto_report.json`, `metadata.json`, `best.md`, `best.txt`, `best.json`, chunk outputs, and `fields.json`/`fields.md` when the document is classified as an invoice.

`extract-fields --profile invoice` writes structured invoice fields and validation checks. `export-xlsx --profile invoice` writes an invoice summary workbook with `invoices` and `items` sheets.

`layout-json` writes page/block/line/span coordinates for PDFs. Use it to trace chunks back to source pages and positions.

`classify` writes or prints the detected profile and strategy. `knowledge-pack` writes a RAG-ready package with `manifest.json`, `chunks.jsonl`, `metadata.json`, `quality_report.json`, `layout.json`, `page_map.json`, and source text.

`metadata` writes JSON by default and can write Markdown with `--format md`.

## Dependency Hints

Use these only when a parser is skipped or the user asks to install dependencies:

```powershell
pip install markitdown pymupdf pypdf pdfplumber pdfminer.six liteparse
pip install python-docx python-pptx openpyxl beautifulsoup4
pip install pytesseract pillow
```

Install system Tesseract OCR separately when using `ocr`, and ensure `tesseract` is on PATH.

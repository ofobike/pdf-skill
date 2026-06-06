# Command Guide

Load this file when exact command syntax, supported formats, or output files are needed.

## Script

Run from the skill directory or pass the absolute script path:

```powershell
python scripts\parse_document_compare.py <command> <path> [options]
```

The script accepts a file path without a subcommand as a shortcut for `compare`.

Current exposed commands:

| Command | Use |
|---|---|
| `compare` | Run multiple parsers against one file and create quality reports. |
| `vote` | Run multiple PDF parsers, vote by quality/consensus/coverage/structure/repetition, and write best output. Supports invoice-weighted customer delivery, lightweight business profile fields, default field-level invoice fusion, optional field layout bbox matching, field evidence images, review HTML, OCR parser voting, optional preflight parser probing, timeout, parser health cache, and result cache. |
| `customer-pack` | Build a customer-ready PDF package with best text, customer fields, customer JSON schema, field evidence images, review HTML, table voting, layout, metadata, manifest, vote report, and preflight report. |
| `batch-customer-pack` | Build customer-ready PDF packages for a directory and write `index.json/index.md`. |
| `probe` | Run real extraction smoke tests for registered parsers and report runtime readiness. |
| `convert` | Convert one file to `md`, `txt`, `json`, or `all`. |
| `batch` | Convert matching files in a directory. |
| `scan-dir` | Scan a directory for extraction quality and optional quality gates. |
| `tables` | Extract PDF tables to `md`, `csv`, `json`, or `all`. |
| `table-vote` | Run multiple table extractors and vote on the best table output. |
| `init-golden` | Initialize a Golden PDF case library from files or directories. |
| `eval-golden` | Run golden PDF case JSON files and detect parser voting/text-quality regressions. |
| `doctor` | Check parser, OCR, and opendataloader dependency availability. |
| `metadata` | Extract metadata; PDFs include page count, embedded metadata, TOC preview, and text-layer hints. |
| `chunk` | Split extracted text into `jsonl`, `json`, `md`, `txt`, or `all` by character or PDF page. |
| `render-pages` | Render selected PDF pages to PNG screenshots. |
| `ocr` | Run real Tesseract OCR on selected PDF pages. |
| `auto` | Run metadata, best text extraction, chunks, classification, report, and invoice fields when detected. |
| `extract-fields` | Extract structured fields; invoice profile is supported. |
| `export-xlsx` | Batch extract invoice fields and write an XLSX summary. |
| `layout-json` | Output PDF page, block, line, span, font, and coordinate data. |
| `verify-fields` | Validate structured fields; invoice profile supports strict checks. |
| `classify` | Detect document profile and recommend a processing strategy. |
| `knowledge-pack` | Generate chunks, metadata, manifest, quality report, layout, source text, and page map. |
| `batch-knowledge` | Generate knowledge packs for a directory. |
| `qa` | Run local extractive QA over a knowledge pack, chunks file, or document. |
| `diff-docs` | Compare two document versions by extracted text, line diff, classification, and optional invoice fields. |

`qa` is extractive local retrieval. It does not call an LLM; it returns source-backed snippets and citations.

## Supported Inputs

| Format | Extensions | Parsers |
|---|---|---|
| PDF | `.pdf` | `markitdown`, `pymupdf4llm`, `docling`, `pspdfkit`/`pdf-to-markdown` CLI, `pymupdf`, `pypdf`, `pdfplumber`, `pdfminer`, `liteparse`, `opendataloader`, `ocr-tesseract` |
| Word | `.docx` | `markitdown`, `python-docx` |
| PPT | `.pptx` | `markitdown`, `python-pptx` |
| Excel | `.xlsx` | `markitdown`, `openpyxl` |
| HTML | `.html`, `.htm` | `markitdown`, `beautifulsoup4` |
| Text-like | `.txt`, `.csv`, `.json`, `.xml`, `.md` | `markitdown` |
| Image | `.jpg`, `.png`, `.gif`, `.bmp`, `.tiff`, `.webp` | `markitdown` metadata/OCR paths, dependency-sensitive |
| Audio | `.mp3`, `.wav` | `markitdown`, dependency-sensitive |
| EPub | `.epub` | `markitdown` |
| Archive | `.zip` | `markitdown` |

Legacy binary Office formats such as `.doc`, `.ppt`, and `.xls` may be detected, but prefer asking for modern `.docx`, `.pptx`, or `.xlsx` when parser support is unreliable.

## Common Commands

```powershell
python scripts\parse_document_compare.py doctor --format pdf --ocr --opendataloader --json
python scripts\parse_document_compare.py compare "D:\path\file.pdf" --max-pages 3 --output-format md
python scripts\parse_document_compare.py vote "D:\path\file.pdf" --probe-before-vote --max-pages 3 --format all
python scripts\parse_document_compare.py vote "D:\path\file.pdf" --probe-before-vote --max-pages 0 --format all --parsers markitdown,pymupdf4llm,docling,pymupdf,pdfplumber,ocr-tesseract --timeout 120 --parser-health-cache
python scripts\parse_document_compare.py customer-pack "D:\path\file.pdf" --max-pages 0 --table-pages all --layout-max-pages 30
python scripts\parse_document_compare.py customer-pack "D:\path\file.pdf" --max-pages 0 --table-pages all --layout-max-pages 30 --timeout 120 --parser-health-cache --result-cache
python scripts\parse_document_compare.py batch-customer-pack "D:\docs" --recursive --max-pages 0 --table-pages all --layout-max-pages 30 --timeout 120 --parser-health-cache --result-cache
python scripts\parse_document_compare.py vote "D:\path\invoice.pdf" --probe-before-vote --max-pages 0 --format all --profile invoice --customer
python scripts\parse_document_compare.py vote "D:\path\invoice.pdf" --probe-before-vote --max-pages 0 --format all --profile invoice --customer --field-layout --field-layout-max-pages 0
python scripts\parse_document_compare.py vote "D:\path\invoice.pdf" --probe-before-vote --max-pages 0 --format all --profile invoice --customer --field-layout --field-layout-max-pages 0 --field-evidence --review-html
python scripts\parse_document_compare.py vote "D:\path\invoice.pdf" --probe-before-vote --max-pages 0 --format all --profile invoice --customer --no-field-fusion
python scripts\parse_document_compare.py vote "D:\path\contract.pdf" --probe-before-vote --max-pages 0 --format all --profile contract --customer --field-layout --field-evidence --review-html
python scripts\parse_document_compare.py probe "D:\path\file.pdf" --max-pages 1 --timeout 120 --parser-health-cache --result-cache --keep-outputs
python scripts\parse_document_compare.py compare "D:\path\file.pdf" --start-page 100 --max-pages 10 --output-format all
python scripts\parse_document_compare.py convert "D:\path\file.docx" --parser markitdown --format json -o "D:\path\file.json"
python scripts\parse_document_compare.py batch "D:\docs" --ext .pdf,.docx --parser markitdown --format txt
python scripts\parse_document_compare.py scan-dir "D:\docs" --ext .pdf --max-pages 3 --min-quality 0.6 --fail-on-bad
python scripts\parse_document_compare.py tables "D:\path\file.pdf" --pages 1-5 --format all
python scripts\parse_document_compare.py table-vote "D:\path\file.pdf" --pages all --format all
python scripts\parse_document_compare.py init-golden "D:\path\2022年上半年" "D:\path\模拟题库" --recursive --out-dir "D:\path\golden_cases" --max-pages 3 --include-ocr
python scripts\parse_document_compare.py eval-golden "D:\path\golden_cases" --recursive --timeout 120 --parser-health-cache --result-cache
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
python scripts\parse_document_compare.py qa "D:\path\pack" "合同金额是多少？" --format md
python scripts\parse_document_compare.py qa "D:\path\pack\chunks.jsonl" "付款期限是什么？" --format all
python scripts\parse_document_compare.py diff-docs "D:\path\old.pdf" "D:\path\new.pdf" --format all
python scripts\parse_document_compare.py compare "D:\path\file.pdf" --max-pages 3 --ocr-fallback
```

## Outputs

`compare` creates an output directory beside the input file unless `--out-dir` is provided. It writes:

- `compare_report.md`
- `compare_report.json`
- `<parser>.md`, `<parser>.txt`, or `<parser>.json`, depending on `--output-format`

Other command outputs:

- `vote`: `vote_report.md`, `vote_report.json`, and `best.md`/`best.txt`/`best.json` depending on `--format`. With `--customer`, also writes `customer_best.md`/`customer_best.txt`/`customer_best.json` and `customer_best.schema.json`; with `--profile invoice`, those customer files contain structured invoice fields, validation results, `field_confidence`, and `audit.field_fusion`. Field-level fusion is enabled by default and can be disabled with `--no-field-fusion`. With business profiles such as `contract`, `bank_statement`, `quotation`, `purchase_order`, `report`, and `annual_report`, customer files contain lightweight `fields` and `field_confidence`. With `--field-layout`, it also writes `field_layout.json` and fills field `bbox/location` when layout matching succeeds. With `--field-evidence`, it writes `field_evidence/*`; with `--review-html`, it writes `review.html`. With `--probe-before-vote`, it first writes `preflight_probe/probe_report.md` and `preflight_probe/probe_report.json`, then only sends `ready` parsers into the final vote.
- `customer-pack`: `manifest.json`, `README.md`, `metadata.json`, `vote_report.md/json`, `preflight_probe/probe_report.md/json`, `best.md/txt/json`, `customer_best.md/txt/json`, `customer_best.schema.json`, `field_evidence/field_evidence.md/json`, field evidence PNGs, `review.html`, `tables/table_vote_report.md/json`, `tables/best_tables.md/csv/json`, `layout.json`, and `page_map.json`.
- `batch-customer-pack`: per-file customer-pack directories plus top-level `index.json` and `index.md`.
- `probe`: `probe_report.md`, `probe_report.json`, and per-parser sample text. It distinguishes `ready`, `dependency_missing`, `runtime_failed`, `quality_failed`, and `timeout`.
- `table-vote`: `table_vote_report.md/json` and `best_tables.md/csv/json`.
- `init-golden`: `cases/*.json`, `index.json`, and `README.md`; case files reference the original PDFs rather than copying them.
- `eval-golden`: top-level `golden_report.md/json`, plus per-case vote outputs under the evaluation output directory.
- `auto`: `auto_report.md`, `auto_report.json`, `metadata.json`, `best.md`, `best.txt`, `best.json`, chunk outputs, and `fields.json`/`fields.md` for invoices.
- `scan-dir`: `scan_report.md`, `scan_report.json`, and `scan_report.csv`.
- `chunk`: JSONL by default, with source metadata, parser name, offsets, and text.
- `render-pages`: PNG files plus `render_pages.json`.
- `ocr`: recognized text in `txt`, `md`, or `json`.
- `extract-fields --profile invoice`: structured invoice fields and validation checks.
- `export-xlsx --profile invoice`: workbook with `invoices` and `items` sheets.
- `layout-json`: PDF coordinates for page/block/line/span data.
- `classify`: detected profile and recommended strategy.
- `knowledge-pack`: `manifest.json`, `chunks.jsonl`, `metadata.json`, `quality_report.json`, `layout.json`, `page_map.json`, and source text.
- `qa`: `qa_report.md`/`qa_report.json` with answer snippets, retrieval scores, and citations.
- `diff-docs`: `diff_report.md`/`diff_report.json` with similarity, changed blocks, unified diff, and invoice field changes when applicable.

## Validation And Packaging

```powershell
python -m py_compile scripts\parse_document_compare.py parse_pdf_compare.py scripts\package_skill.py
python -m unittest discover -s tests
python scripts\package_skill.py --force
```

`scripts\package_skill.py` writes a clean skill package to `dist\pdf-parse-skill` by default.

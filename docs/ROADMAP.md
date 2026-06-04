# PDF Parse Skill Roadmap

This project is evolving from a PDF parser comparison helper into a PDF-first document intelligence skill.

## Version Priorities

### v4.1 - Skill Usability And Progressive Disclosure

Goal: make the skill easier for Codex to trigger correctly and cheaper to load.

- Done in `4.1.0`: renamed the Codex-facing skill trigger to `pdf-parse-skill`.
- Done in `4.1.0`: reduced `SKILL.md` to workflow, command selection, and OCR quality rules.
- Done in `4.1.0`: moved detailed command examples and output descriptions to `references/command-guide.md`.
- Done in `4.1.0`: expanded dependency guidance for Tesseract OCR, Chinese OCR language data, `opendataloader-pdf`, and Java.

### v3.6 - Smart Extraction Foundation

Goal: make one-file and invoice workflows feel automatic.

- `auto`: run metadata, sampled parser comparison, quality routing, best-text export, and a compact decision report.
- `extract-fields --profile invoice`: extract structured invoice fields from ordinary text extraction output.
- `export-xlsx --profile invoice`: batch extract invoice fields and write an Excel summary.
- Update `SKILL.md` and `README.md` with trigger examples for automatic extraction and invoice workflows.

### v3.7 - OCR Routing And Layout Awareness

Goal: make poor PDFs recoverable and outputs more traceable.

- Done in `4.0.0`: `auto --auto-ocr` triggers OCR when ordinary extraction quality is below threshold.
- Done in `4.0.0`: `layout-json` outputs pages, blocks, lines, spans, fonts, and coordinates for PDFs.
- Done in `4.0.0`: `verify-fields` validates invoice totals, line-item sums, tax amount, required fields, and ID formats in strict mode.
- Done in `4.0.0`: `auto` records OCR fallback attempts and uses OCR text when OCR quality improves the result.

### v4.0 - Document Intelligence Pack

Goal: produce reusable knowledge and business-data packages.

- Done in `4.0.0`: `classify` detects invoice, contract, textbook, notice, scanned/image PDFs, and generic documents with strategy hints.
- Done in `4.0.0`: `knowledge-pack` generates chunks, metadata, manifest, quality report, layout, source text, and page map.
- Done in `4.0.0`: `batch-knowledge` produces knowledge packs for a whole directory.
- Done in `4.1.0`: `qa` exposes local extractive QA over knowledge packs, chunks files, or documents.
- Done in `4.1.0`: `diff-docs` compares two document versions by extracted text, line diff, classification, and invoice fields when applicable.

## Design Rules

- Keep core deterministic workflows in `scripts/parse_document_compare.py`.
- Keep `SKILL.md` concise; move long planning notes to this `docs/` folder.
- Prefer clear JSON outputs for automation, Markdown for human review, and XLSX for business summaries.
- Do not install dependencies automatically. Use `doctor` and actionable dependency hints.
- Preserve source files. All commands should write outputs beside the source or to explicit output directories.

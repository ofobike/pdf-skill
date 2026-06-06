# Dependency Notes

Load this file only when diagnosing missing parser dependencies or preparing install instructions.

## Parser Packages

| Parser name | Python package | Typical formats |
|---|---|---|
| `markitdown` | `markitdown` | PDF, DOCX, PPTX, XLSX, HTML |
| `pymupdf4llm` | `pymupdf4llm` | PDF to LLM/RAG-oriented Markdown |
| `docling` | `docling` | PDF to structured document/Markdown |
| `pspdfkit` | `@pspdfkit/pdf-to-markdown` CLI | PDF to Markdown via external Node CLI |
| `pymupdf` | `pymupdf` | PDF |
| `pypdf` | `pypdf` | PDF |
| `pdfplumber` | `pdfplumber` | PDF text and tables |
| `pdfminer` | `pdfminer.six` | PDF |
| `liteparse` | `liteparse` | PDF |
| `opendataloader` | `opendataloader-pdf` + Java | PDF |
| `ocr-tesseract` | `pymupdf`, `pytesseract`, `pillow` + system Tesseract | PDF OCR parser for vote/probe/customer-pack |
| `python-docx` | `python-docx` | DOCX |
| `python-pptx` | `python-pptx` | PPTX |
| `openpyxl` | `openpyxl` | XLSX |
| `beautifulsoup4` | `beautifulsoup4` | HTML |
| `pytesseract` | `pytesseract` | OCR Python bridge |
| `Pillow` | `pillow` | OCR image handoff |

`ocr` and the `ocr-tesseract` parser also require the system Tesseract OCR executable. Installing Python packages alone is not enough.
Chinese OCR also requires matching Tesseract language data such as `chi_sim`.

`opendataloader-pdf` requires both the Python wrapper/package and a working `java` command.

## Install Groups

```powershell
pip install markitdown pymupdf4llm docling pymupdf pypdf pdfplumber pdfminer.six liteparse
pip install opendataloader-pdf
npm install -g @pspdfkit/pdf-to-markdown
pip install python-docx python-pptx openpyxl beautifulsoup4
pip install pytesseract pillow
```

`pspdfkit` is implemented as an external CLI parser. The command must be available as `pdf-to-markdown` or `pspdfkit-pdf-to-markdown` on PATH. Official binaries currently target macOS/Linux; Windows availability may lag.

Do not install dependencies automatically unless the user explicitly asks.

## Checks

Use the bundled doctor command before OCR or dependency-heavy parsing:

```powershell
python scripts\parse_document_compare.py doctor --format pdf --ocr --opendataloader --json
```

On Windows, if Python packages are importable but OCR still fails, check that:

- `tesseract` is on PATH.
- The requested language data is installed, for example `chi_sim` for Simplified Chinese.
- `java` is on PATH before using `opendataloader-pdf`.

# Dependency Notes

Load this file only when diagnosing missing parser dependencies or preparing install instructions.

## Parser Packages

| Parser name | Python package | Typical formats |
|---|---|---|
| `markitdown` | `markitdown` | PDF, DOCX, PPTX, XLSX, HTML |
| `pymupdf` | `pymupdf` | PDF |
| `pypdf` | `pypdf` | PDF |
| `pdfplumber` | `pdfplumber` | PDF text and tables |
| `pdfminer` | `pdfminer.six` | PDF |
| `liteparse` | `liteparse` | PDF |
| `python-docx` | `python-docx` | DOCX |
| `python-pptx` | `python-pptx` | PPTX |
| `openpyxl` | `openpyxl` | XLSX |
| `beautifulsoup4` | `beautifulsoup4` | HTML |
| `pytesseract` | `pytesseract` | OCR Python bridge |
| `Pillow` | `pillow` | OCR image handoff |

`ocr` also requires the system Tesseract OCR executable. Installing Python packages alone is not enough.

## Install Groups

```powershell
pip install markitdown pymupdf pypdf pdfplumber pdfminer.six liteparse
pip install python-docx python-pptx openpyxl beautifulsoup4
pip install pytesseract pillow
```

Do not install dependencies automatically unless the user explicitly asks.

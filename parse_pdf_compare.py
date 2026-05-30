"""Compatibility entrypoint for the document parse comparison tool.

The maintained implementation lives in scripts/parse_document_compare.py.
This wrapper keeps older commands such as `python parse_pdf_compare.py ...`
working while the skill uses the standard scripts/ layout.
"""

from __future__ import annotations

import runpy
from pathlib import Path


SCRIPT = Path(__file__).resolve().parent / "scripts" / "parse_document_compare.py"


if __name__ == "__main__":
    runpy.run_path(str(SCRIPT), run_name="__main__")

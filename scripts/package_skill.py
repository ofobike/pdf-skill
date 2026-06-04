"""Create a clean distributable skill directory.

The package contains only files needed for Codex skill use:
SKILL.md, agents/, scripts/, references/, and top-level compatibility entrypoints.
"""
from __future__ import annotations

import argparse
import shutil
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_OUT = ROOT / "dist" / "pdf-parse-skill"
INCLUDE_PATHS = [
    "SKILL.md",
    "agents",
    "scripts",
    "references",
    "parse_pdf_compare.py",
    "requirements.txt",
]


def copy_item(src: Path, dst: Path) -> None:
    if src.is_dir():
        shutil.copytree(
            src,
            dst,
            ignore=shutil.ignore_patterns("__pycache__", "*.pyc", ".pytest_cache", "package_skill.py"),
        )
    else:
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)


def package_skill(out_dir: Path, force: bool = False) -> list[Path]:
    out_dir = out_dir.resolve()
    if out_dir.exists():
        if not force:
            raise FileExistsError(f"Output already exists: {out_dir}. Use --force to replace it.")
        shutil.rmtree(out_dir)
    out_dir.mkdir(parents=True)

    copied: list[Path] = []
    for rel in INCLUDE_PATHS:
        src = ROOT / rel
        if not src.exists():
            raise FileNotFoundError(src)
        dst = out_dir / rel
        copy_item(src, dst)
        copied.append(dst)
    return copied


def main() -> int:
    parser = argparse.ArgumentParser(description="Package pdf-parse-skill into a clean dist directory.")
    parser.add_argument("--out-dir", type=Path, default=DEFAULT_OUT, help=f"Output directory. Default: {DEFAULT_OUT}")
    parser.add_argument("--force", action="store_true", help="Replace output directory if it exists.")
    args = parser.parse_args()

    try:
        copied = package_skill(args.out_dir, args.force)
    except Exception as exc:
        print(f"ERROR: {exc}")
        return 1

    print(f"Packaged skill: {args.out_dir.resolve()}")
    for path in copied:
        print(path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

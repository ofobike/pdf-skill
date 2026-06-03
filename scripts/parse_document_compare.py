"""多格式文档解析工具。

功能：
    compare  — 多解析器对比（默认）
    convert  — 文档转 md/txt/json
    batch    — 批量转换
    scan-dir — 批量质量扫描
    tables   — 表格提取（仅 PDF）
    chunk    — 文本分块
    render-pages — PDF 页面截图
    ocr      — 可选真实 OCR
    qa       — 基于知识包或文档的带引用抽取式问答
    diff-docs — 文档版本差异对比

支持格式：
    PDF  (.pdf)  — markitdown / pymupdf / pypdf / pdfplumber / pdfminer / liteparse
    Word (.docx) — markitdown / python-docx
    PPT  (.pptx) — markitdown / python-pptx
    Excel(.xlsx) — markitdown / openpyxl
    HTML (.html) — markitdown / beautifulsoup4

用法：
    python parse_pdf_compare.py compare <文件> [选项]
    python parse_pdf_compare.py convert <文件> --parser pymupdf --format json -o output.json
    python parse_pdf_compare.py batch <目录> --parser pymupdf --output-dir ./out/
    python parse_pdf_compare.py tables <文件> --pages 1-5 --format csv
    python parse_pdf_compare.py <文件> [选项]  # 等同于 compare
"""
from __future__ import annotations

__version__ = "4.1.0"

import argparse
import csv
import difflib
import importlib.util
import io
import json
import logging
import re
import shutil
import sys
import tempfile
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import asdict, dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Callable

logger = logging.getLogger(__name__)

# 预编译正则
_RE_WHITESPACE = re.compile(r"[ \t]+")
_RE_MULTI_NEWLINE = re.compile(r"\n{3,}")
_RE_INLINE_WHITESPACE = re.compile(r"\s+")
_RE_PARAGRAPH_SPLIT = re.compile(r"\n\s*\n")
_RE_CHINESE = re.compile(r"[一-鿿㐀-䶿\U00020000-\U0002a6df]")
_RE_TOKEN = re.compile(r"[A-Za-z0-9_]+|[一-鿿]")

QUALITY_BAD_LABELS = {"empty", "bad"}
PAGE_CHUNK_PARSERS = {"pymupdf", "pypdf", "pdfplumber", "pdfminer"}
RAG_SNIPPET_CHARS = 500

# ---------------------------------------------------------------------------
# 数据结构
# ---------------------------------------------------------------------------

@dataclass
class ParseResult:
    parser: str
    status: str
    seconds: float
    output_file: str | None = None
    error: str | None = None
    chars: int = 0
    non_space_chars: int = 0
    lines: int = 0
    paragraphs: int = 0
    headings: int = 0
    chinese_chars: int = 0
    cid_markers: int = 0
    control_chars: int = 0
    quality_score: float = 0.0
    quality_label: str = "unknown"
    preview: str = ""


# ---------------------------------------------------------------------------
# 工具函数
# ---------------------------------------------------------------------------

def module_exists(module_name: str) -> bool:
    return importlib.util.find_spec(module_name) is not None


def normalize_text(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = _RE_WHITESPACE.sub(" ", text)
    text = _RE_MULTI_NEWLINE.sub("\n\n", text)
    return text.strip()


def compact_text(text: str) -> str:
    """Return text with all whitespace removed for regexes across layout breaks."""
    return re.sub(r"\s+", "", text or "")


def decimal_from_text(value: str | None) -> float | None:
    """Parse a currency/number field into float without raising."""
    if not value:
        return None
    cleaned = value.replace(",", "").replace("¥", "").replace("￥", "").strip()
    try:
        return float(Decimal(cleaned))
    except (InvalidOperation, ValueError):
        return None


def quality_from_text(
    text: str,
    non_space_chars: int | None = None,
    cid_markers: int | None = None,
    control_chars: int | None = None,
) -> tuple[float, str]:
    """Estimate extraction quality from text-only signals."""
    if non_space_chars is None:
        non_space_chars = len(_RE_INLINE_WHITESPACE.sub("", text))
    if cid_markers is None:
        cid_markers = text.count("(cid:")
    if control_chars is None:
        control_chars = sum(1 for ch in text if ord(ch) < 32 and ch not in "\n\t")

    if non_space_chars == 0:
        return 0.0, "empty"

    chinese_chars = len(_RE_CHINESE.findall(text))
    cid_penalty = min(0.65, cid_markers / max(1, non_space_chars) * 10)
    control_penalty = min(0.65, control_chars / max(1, non_space_chars) * 8)
    length_penalty = 0.25 if non_space_chars < 80 else 0.1 if non_space_chars < 250 else 0.0
    language_bonus = min(0.2, chinese_chars / max(1, non_space_chars) * 0.5)

    score = max(0.0, min(1.0, 0.85 + language_bonus - cid_penalty - control_penalty - length_penalty))
    if score >= 0.75:
        label = "high"
    elif score >= 0.5:
        label = "medium"
    elif score >= 0.25:
        label = "low"
    else:
        label = "bad"
    return round(score, 3), label


def parse_extensions(ext_str: str | None) -> set[str]:
    """Parse comma-separated extensions into normalized dotted suffixes."""
    raw = ext_str or ".pdf"
    extensions = {
        e.strip().lower() if e.strip().startswith(".") else f".{e.strip().lower()}"
        for e in raw.split(",")
        if e.strip()
    }
    return extensions or {".pdf"}


def quality_gate_status(
    result: ParseResult | None,
    min_quality: float,
    fail_on_bad: bool,
) -> dict[str, object]:
    """Return a reusable quality gate payload for compare and scan commands."""
    threshold = max(0.0, min(1.0, min_quality))
    if result is None:
        passed = False
        reason = "no_successful_output"
    elif result.quality_score < threshold:
        passed = False
        reason = "below_min_quality"
    elif fail_on_bad and result.quality_label in QUALITY_BAD_LABELS:
        passed = False
        reason = "bad_quality_label"
    else:
        passed = True
        reason = "ok"
    return {
        "enabled": bool(threshold > 0 or fail_on_bad),
        "passed": passed,
        "min_quality": threshold,
        "fail_on_bad": fail_on_bad,
        "reason": reason,
        "best_parser": result.parser if result else None,
        "best_quality_score": result.quality_score if result else 0.0,
        "best_quality_label": result.quality_label if result else "none",
    }


def compute_page_range(page_count: int, start_page: int, max_pages: int | None) -> tuple[int, int]:
    """根据总页数、起始页、最大页数计算实际的 (start_index, end_index)。"""
    start_index = min(start_page - 1, page_count - 1)
    if max_pages is None:
        end_index = page_count - 1
    else:
        end_index = min(start_index + max_pages - 1, page_count - 1)
    return start_index, end_index


def get_pdf_page_count(pdf_path: Path) -> int:
    """获取 PDF 总页数，优先用 fitz，后备用 pypdf。"""
    if module_exists("fitz"):
        import fitz  # type: ignore
        doc = fitz.open(pdf_path)
        count = doc.page_count
        doc.close()
        return count
    if module_exists("pypdf"):
        from pypdf import PdfReader  # type: ignore
        return len(PdfReader(str(pdf_path)).pages)
    raise ImportError("需要 fitz 或 pypdf 来获取 PDF 页数")


def selected_pdf_page_indices(pdf_path: Path, start_page: int, max_pages: int | None) -> list[int]:
    """Return selected PDF page indices using 1-based CLI page arguments."""
    page_count = get_pdf_page_count(pdf_path)
    start_index, end_index = compute_page_range(page_count, start_page, max_pages)
    return list(range(start_index, end_index + 1))


def text_stats(parser: str, status: str, seconds: float, text: str, output_file: Path | None) -> ParseResult:
    """从已标准化的文本生成统计数据。"""
    lines = text.splitlines()
    paragraphs = [p for p in _RE_PARAGRAPH_SPLIT.split(text) if p.strip()]
    non_space_chars = len(_RE_INLINE_WHITESPACE.sub("", text))
    cid_markers = text.count("(cid:")
    control_chars = sum(1 for ch in text if ord(ch) < 32 and ch not in "\n\t")
    quality_score, quality_label = quality_from_text(text, non_space_chars, cid_markers, control_chars)
    return ParseResult(
        parser=parser,
        status=status,
        seconds=round(seconds, 3),
        output_file=str(output_file) if output_file else None,
        chars=len(text),
        non_space_chars=non_space_chars,
        lines=len(lines),
        paragraphs=len(paragraphs),
        headings=sum(1 for line in lines if line.lstrip().startswith("#")),
        chinese_chars=len(_RE_CHINESE.findall(text)),
        cid_markers=cid_markers,
        control_chars=control_chars,
        quality_score=quality_score,
        quality_label=quality_label,
        preview=text[:600],
    )


def result_from_text(parser: str, text: str, seconds: float = 0.0) -> ParseResult:
    """Build a ParseResult from already-normalized text."""
    return text_stats(parser, "ok", seconds, normalize_text(text), None)


TEXT_OUTPUT_FORMATS = {"md", "txt", "json", "all"}


def output_suffixes(output_format: str) -> list[str]:
    """Return concrete output suffixes for a requested text export format."""
    if output_format == "all":
        return ["md", "txt", "json"]
    return [output_format]


def default_output_path(source_file: Path, output_format: str) -> Path:
    """Build the default convert output path for md/txt/json/all."""
    if output_format == "all":
        return source_file.parent / f"{source_file.stem}_converted"
    return source_file.parent / f"{source_file.stem}.{output_format}"


def write_extracted_outputs(
    out_dir: Path,
    parser_name: str,
    text: str,
    output_format: str = "md",
) -> tuple[Path, str]:
    """Normalize text, write requested formats, and return the primary output path."""
    normalized = normalize_text(text)
    out_dir.mkdir(parents=True, exist_ok=True)
    paths: dict[str, Path] = {}

    if output_format in ("md", "all"):
        md_path = out_dir / f"{parser_name}.md"
        md_path.write_text(normalized + "\n", encoding="utf-8")
        paths["md"] = md_path

    if output_format in ("txt", "all"):
        txt_path = out_dir / f"{parser_name}.txt"
        txt_path.write_text(normalized + "\n", encoding="utf-8")
        paths["txt"] = txt_path

    if output_format in ("json", "all"):
        json_path = out_dir / f"{parser_name}.json"
        payload = {
            "version": __version__,
            "parser": parser_name,
            "format": "text",
            "chars": len(normalized),
            "text": normalized,
        }
        json_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        paths["json"] = json_path

    primary_suffix = "md" if output_format == "all" else output_format
    return paths[primary_suffix], normalized


def write_single_conversion_output(
    out_path: Path,
    parser_name: str,
    source_file: Path,
    text: str,
    output_format: str,
) -> list[Path]:
    """Write one conversion result as md, txt, json, or all formats."""
    normalized = normalize_text(text)
    written: list[Path] = []
    suffixes = output_suffixes(output_format)

    if output_format == "all":
        out_path.mkdir(parents=True, exist_ok=True)
        targets = {suffix: out_path / f"{source_file.stem}.{suffix}" for suffix in suffixes}
    else:
        out_path.parent.mkdir(parents=True, exist_ok=True)
        targets = {output_format: out_path}

    for suffix, target in targets.items():
        if suffix in ("md", "txt"):
            target.write_text(normalized + "\n", encoding="utf-8")
        elif suffix == "json":
            payload = {
                "version": __version__,
                "source_file": str(source_file),
                "parser": parser_name,
                "format": "text",
                "chars": len(normalized),
                "text": normalized,
            }
            target.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        written.append(target)

    return written


def write_text(out_dir: Path, parser_name: str, text: str) -> tuple[Path, str]:
    """Backward-compatible Markdown writer."""
    return write_extracted_outputs(out_dir, parser_name, text, "md")


def parse_page_spec(spec: str, total_pages: int) -> list[int]:
    """解析页码规格，如 '1-5' 或 '1,3,5' 或 '1-3,7,10-12'。返回 0-based 索引列表。"""
    pages: set[int] = set()
    for part in spec.split(","):
        part = part.strip()
        if "-" in part:
            start_s, end_s = part.split("-", 1)
            start = max(1, int(start_s))
            end = min(total_pages, int(end_s))
            pages.update(range(start - 1, end))
        else:
            p = int(part)
            if 1 <= p <= total_pages:
                pages.add(p - 1)
    return sorted(pages)


def resolve_page_indices(spec: str | None, total_pages: int, default_all: bool = False) -> list[int]:
    """Resolve CLI page specs; supports None, 'all', '*', and parse_page_spec syntax."""
    if total_pages <= 0:
        return []
    if spec is None or not spec.strip():
        return list(range(total_pages)) if default_all else [0]
    normalized = spec.strip().lower()
    if normalized in {"all", "*"}:
        return list(range(total_pages))
    pages = parse_page_spec(spec, total_pages)
    if not pages:
        raise ValueError(f"页码范围无效：{spec}")
    return pages


# ---------------------------------------------------------------------------
# 格式检测
# ---------------------------------------------------------------------------

EXTENSION_FORMAT_MAP: dict[str, str] = {
    ".pdf": "pdf",
    ".docx": "word",
    ".doc": "word",
    ".pptx": "ppt",
    ".ppt": "ppt",
    ".xlsx": "excel",
    ".xls": "excel",
    ".html": "html",
    ".htm": "html",
    # 文本格式
    ".txt": "text",
    ".csv": "text",
    ".json": "text",
    ".xml": "text",
    ".md": "text",
    ".markdown": "text",
    # 图片格式
    ".jpg": "image",
    ".jpeg": "image",
    ".png": "image",
    ".gif": "image",
    ".bmp": "image",
    ".tiff": "image",
    ".tif": "image",
    ".webp": "image",
    # 音频格式
    ".mp3": "audio",
    ".wav": "audio",
    # 电子书
    ".epub": "epub",
    # 压缩包
    ".zip": "archive",
}

FORMAT_PARSERS: dict[str, list[str]] = {
    "pdf":     ["markitdown", "pymupdf", "pypdf", "pdfplumber", "pdfminer", "liteparse", "opendataloader"],
    "word":    ["markitdown", "python-docx"],
    "ppt":     ["markitdown", "python-pptx"],
    "excel":   ["markitdown", "openpyxl"],
    "html":    ["markitdown", "beautifulsoup4"],
    "text":    ["markitdown"],
    "image":   ["markitdown"],
    "audio":   ["markitdown"],
    "epub":    ["markitdown"],
    "archive": ["markitdown"],
}


PARSER_MODULES: dict[str, list[str]] = {
    "markitdown": ["markitdown"],
    "pymupdf": ["fitz"],
    "pypdf": ["pypdf", "PyPDF2"],
    "pdfplumber": ["pdfplumber"],
    "pdfminer": ["pdfminer.high_level"],
    "liteparse": ["liteparse"],
    "opendataloader": ["opendataloader_pdf"],
    "python-docx": ["docx"],
    "python-pptx": ["pptx"],
    "openpyxl": ["openpyxl"],
    "beautifulsoup4": ["bs4"],
}


PARSER_PACKAGES: dict[str, str] = {
    "markitdown": "markitdown",
    "pymupdf": "pymupdf",
    "pypdf": "pypdf",
    "pdfplumber": "pdfplumber",
    "pdfminer": "pdfminer.six",
    "liteparse": "liteparse",
    "opendataloader": "opendataloader-pdf",
    "python-docx": "python-docx",
    "python-pptx": "python-pptx",
    "openpyxl": "openpyxl",
    "beautifulsoup4": "beautifulsoup4",
}


OCR_DEPENDENCIES: list[dict[str, object]] = [
    {
        "name": "pytesseract",
        "available": module_exists("pytesseract"),
        "kind": "python",
        "package": "pytesseract",
    },
    {
        "name": "Pillow",
        "available": module_exists("PIL"),
        "kind": "python",
        "package": "pillow",
    },
    {
        "name": "tesseract",
        "available": shutil.which("tesseract") is not None,
        "kind": "system",
        "package": "tesseract-ocr",
    },
]

OPENLOADER_DEPENDENCIES: list[dict[str, object]] = [
    {
        "name": "opendataloader_pdf",
        "available": module_exists("opendataloader_pdf"),
        "kind": "python",
        "package": "opendataloader-pdf",
    },
    {
        "name": "java",
        "available": shutil.which("java") is not None,
        "kind": "system",
        "package": "JDK/JRE (java command)",
    },
]


def detect_format(file_path: Path) -> str | None:
    """根据文件扩展名检测格式。"""
    return EXTENSION_FORMAT_MAP.get(file_path.suffix.lower())


def get_parsers_for_format(fmt: str) -> list[str]:
    """获取指定格式的所有可用解析器。"""
    return FORMAT_PARSERS.get(fmt, [])


def parser_available(parser_name: str) -> bool:
    """Return whether at least one module for a parser is importable."""
    modules = PARSER_MODULES.get(parser_name, [parser_name])
    return any(module_exists(module_name) for module_name in modules)


def parser_dependency_rows(fmt_filter: str | None = None) -> list[dict[str, object]]:
    """Build parser dependency rows for doctor output."""
    formats = [fmt_filter] if fmt_filter else sorted(FORMAT_PARSERS)
    rows: list[dict[str, object]] = []
    for fmt in formats:
        for parser_name in FORMAT_PARSERS.get(fmt, []):
            modules = PARSER_MODULES.get(parser_name, [parser_name])
            package = PARSER_PACKAGES.get(parser_name, parser_name)
            available = parser_available(parser_name)
            rows.append(
                {
                    "format": fmt,
                    "parser": parser_name,
                    "available": available,
                    "status": "ok" if available else "missing",
                    "modules": modules,
                    "package": package,
                }
            )
    return rows


def extract_text_with_parser(
    file_path: Path,
    parser_name: str,
    start_page: int = 1,
    max_pages: int | None = None,
) -> str:
    """Extract text with one parser and handle MarkItDown PDF page slicing."""
    fmt = detect_format(file_path)
    if fmt is None:
        raise ValueError(f"不支持的文件格式：{file_path.suffix}")
    parser_name = resolve_parser_name(parser_name)
    if parser_name not in get_parsers_for_format(fmt):
        raise ValueError(f"{fmt} 格式不支持解析器 {parser_name}")
    extract_func = get_extractor(fmt, parser_name)
    if extract_func is None:
        raise ValueError(f"未找到解析器 {parser_name}")

    temp_dir_obj: tempfile.TemporaryDirectory[str] | None = None
    try:
        target_path = file_path
        if fmt == "pdf" and parser_name == "markitdown":
            temp_dir_obj = tempfile.TemporaryDirectory(prefix="parse_extract_")
            target_path, skip_reason = prepare_markitdown_pdf_target(
                file_path, start_page, max_pages, Path(temp_dir_obj.name)
            )
            if skip_reason:
                raise RuntimeError(skip_reason)
        return extract_func(target_path, start_page, max_pages)
    finally:
        if temp_dir_obj:
            temp_dir_obj.cleanup()


def extract_best_text(
    file_path: Path,
    parser_name: str = "auto",
    start_page: int = 1,
    max_pages: int | None = None,
) -> tuple[str, ParseResult]:
    """Extract text with an explicit parser or choose the best available parser."""
    fmt = detect_format(file_path)
    if fmt is None:
        raise ValueError(f"不支持的文件格式：{file_path.suffix}")

    if parser_name != "auto":
        start = time.perf_counter()
        raw_text = extract_text_with_parser(file_path, parser_name, start_page, max_pages)
        normalized = normalize_text(raw_text)
        return normalized, text_stats(resolve_parser_name(parser_name), "ok", time.perf_counter() - start, normalized, None)

    results: list[tuple[str, ParseResult]] = []
    errors: list[str] = []
    for candidate in get_parsers_for_format(fmt):
        start = time.perf_counter()
        try:
            raw_text = extract_text_with_parser(file_path, candidate, start_page, max_pages)
            normalized = normalize_text(raw_text)
            results.append((normalized, text_stats(candidate, "ok", time.perf_counter() - start, normalized, None)))
        except Exception as exc:
            errors.append(f"{candidate}: {exc}")

    recommended = choose_recommended_result([result for _, result in results])
    if recommended is None:
        raise RuntimeError("没有成功解析结果：" + "; ".join(errors))
    for text, result in results:
        if result.parser == recommended.parser:
            return text, result
    raise RuntimeError("无法匹配推荐解析器结果")


def file_metadata(file_path: Path) -> dict[str, object]:
    """Collect file-level metadata; add PDF details when possible."""
    fmt = detect_format(file_path)
    stat = file_path.stat()
    payload: dict[str, object] = {
        "version": __version__,
        "source_file": str(file_path),
        "file_name": file_path.name,
        "extension": file_path.suffix.lower(),
        "format": fmt,
        "size_bytes": stat.st_size,
        "modified_at": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(stat.st_mtime)),
    }

    if fmt == "pdf":
        payload["pdf"] = pdf_metadata(file_path)

    return payload


def pdf_metadata(pdf_path: Path) -> dict[str, object]:
    """Collect PDF metadata using PyMuPDF when available, falling back to pypdf."""
    if module_exists("fitz"):
        import fitz  # type: ignore

        doc = fitz.open(pdf_path)
        try:
            sample_pages = min(5, doc.page_count)
            text_chars_by_page: list[int] = []
            for index in range(sample_pages):
                text_chars_by_page.append(len((doc.load_page(index).get_text("text") or "").strip()))
            toc = doc.get_toc(simple=True)
            return {
                "page_count": doc.page_count,
                "metadata": dict(doc.metadata or {}),
                "toc_count": len(toc),
                "toc_preview": [{"level": item[0], "title": item[1], "page": item[2]} for item in toc[:30]],
                "sample_text_chars_by_page": text_chars_by_page,
                "has_sample_text_layer": any(count > 0 for count in text_chars_by_page),
                "parser": "pymupdf",
            }
        finally:
            doc.close()

    if module_exists("pypdf"):
        from pypdf import PdfReader  # type: ignore

        reader = PdfReader(str(pdf_path))
        metadata = {}
        if reader.metadata:
            metadata = {str(k): str(v) for k, v in dict(reader.metadata).items()}
        return {
            "page_count": len(reader.pages),
            "metadata": metadata,
            "toc_count": 0,
            "toc_preview": [],
            "has_sample_text_layer": None,
            "parser": "pypdf",
        }

    raise ImportError("需要 PyMuPDF/fitz 或 pypdf 才能读取 PDF 元数据")


def metadata_to_markdown(metadata: dict[str, object]) -> str:
    """Render metadata as Markdown."""
    lines = [
        "# Document Metadata",
        "",
        f"- File: `{metadata.get('source_file')}`",
        f"- Format: `{metadata.get('format')}`",
        f"- Size bytes: `{metadata.get('size_bytes')}`",
        f"- Modified: `{metadata.get('modified_at')}`",
    ]

    pdf = metadata.get("pdf")
    if isinstance(pdf, dict):
        lines.extend(
            [
                "",
                "## PDF",
                "",
                f"- Page count: `{pdf.get('page_count')}`",
                f"- Metadata parser: `{pdf.get('parser')}`",
                f"- TOC entries: `{pdf.get('toc_count')}`",
                f"- Has sample text layer: `{pdf.get('has_sample_text_layer')}`",
                f"- Sample text chars by page: `{pdf.get('sample_text_chars_by_page')}`",
            ]
        )
        doc_meta = pdf.get("metadata")
        if isinstance(doc_meta, dict) and doc_meta:
            lines.extend(["", "## Embedded Metadata", ""])
            for key, value in doc_meta.items():
                if value:
                    lines.append(f"- `{key}`: {value}")
        toc_preview = pdf.get("toc_preview")
        if isinstance(toc_preview, list) and toc_preview:
            lines.extend(["", "## TOC Preview", ""])
            for item in toc_preview:
                if isinstance(item, dict):
                    lines.append(f"- L{item.get('level')} p{item.get('page')}: {item.get('title')}")

    return "\n".join(lines) + "\n"


def split_text_into_chunks(text: str, chunk_size: int, overlap: int) -> list[dict[str, object]]:
    """Split text into overlapping character chunks."""
    normalized = normalize_text(text)
    if chunk_size <= 0:
        raise ValueError("chunk_size must be positive")
    if overlap < 0:
        raise ValueError("overlap must be >= 0")
    if overlap >= chunk_size:
        raise ValueError("overlap must be smaller than chunk_size")

    chunks: list[dict[str, object]] = []
    start = 0
    chunk_id = 1
    while start < len(normalized):
        end = min(start + chunk_size, len(normalized))
        chunk_text = normalized[start:end]
        chunks.append(
            {
                "chunk_id": chunk_id,
                "start": start,
                "end": end,
                "chars": len(chunk_text),
                "text": chunk_text,
            }
        )
        if end == len(normalized):
            break
        start = end - overlap
        chunk_id += 1
    return chunks


def split_pages_into_chunks(pages: list[dict[str, object]]) -> list[dict[str, object]]:
    """Build one chunk per source page."""
    chunks: list[dict[str, object]] = []
    for chunk_id, page in enumerate(pages, start=1):
        text = normalize_text(str(page.get("text", "")))
        quality_score, quality_label = quality_from_text(text)
        chunks.append(
            {
                "chunk_id": chunk_id,
                "chunk_by": "page",
                "page": page.get("page"),
                "start": 0,
                "end": len(text),
                "chars": len(text),
                "quality_score": quality_score,
                "quality_label": quality_label,
                "text": text,
            }
        )
    return chunks


def write_chunks(chunks: list[dict[str, object]], out_path: Path, output_format: str, metadata: dict[str, object]) -> list[Path]:
    """Write chunks as jsonl, json, md, txt, or all."""
    if output_format == "all":
        out_path.mkdir(parents=True, exist_ok=True)
        targets = {
            "jsonl": out_path / "chunks.jsonl",
            "json": out_path / "chunks.json",
            "md": out_path / "chunks.md",
            "txt": out_path / "chunks.txt",
        }
    else:
        out_path.parent.mkdir(parents=True, exist_ok=True)
        targets = {output_format: out_path}

    written: list[Path] = []
    for suffix, target in targets.items():
        if suffix == "jsonl":
            lines = [json.dumps({**metadata, **chunk}, ensure_ascii=False) for chunk in chunks]
            target.write_text("\n".join(lines) + ("\n" if lines else ""), encoding="utf-8")
        elif suffix == "json":
            payload = {"metadata": metadata, "chunk_count": len(chunks), "chunks": chunks}
            target.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        elif suffix == "md":
            parts = ["# Document Chunks", ""]
            for chunk in chunks:
                parts.extend([f"## Chunk {chunk['chunk_id']}", "", str(chunk["text"]), ""])
            target.write_text("\n".join(parts), encoding="utf-8")
        elif suffix == "txt":
            parts = []
            for chunk in chunks:
                page_suffix = f" page {chunk['page']}" if "page" in chunk else ""
                parts.append(f"--- chunk {chunk['chunk_id']}{page_suffix} [{chunk['start']}:{chunk['end']}] ---")
                parts.append(str(chunk["text"]))
            target.write_text("\n\n".join(parts) + ("\n" if parts else ""), encoding="utf-8")
        written.append(target)
    return written


# ---------------------------------------------------------------------------
# PDF 解析器
# ---------------------------------------------------------------------------

def _extract_pdf_markitdown(pdf_path: Path, start_page: int, max_pages: int | None) -> str:
    if not module_exists("markitdown"):
        raise ImportError("未安装 markitdown")
    from markitdown import MarkItDown  # type: ignore
    result = MarkItDown().convert(str(pdf_path))
    return getattr(result, "text_content", "") or ""


def prepare_markitdown_pdf_target(
    pdf_path: Path,
    start_page: int,
    max_pages: int | None,
    tmp_dir: Path,
) -> tuple[Path, str | None]:
    """Return a page-limited PDF for MarkItDown when possible."""
    if max_pages is None:
        return pdf_path, None

    subset = create_subset_pdf(pdf_path, start_page, max_pages, tmp_dir)
    if subset:
        return subset, None

    return pdf_path, "设置了 --max-pages，但当前环境无法生成临时子 PDF；已跳过 markitdown"


def _extract_pdf_pymupdf(pdf_path: Path, start_page: int, max_pages: int | None) -> str:
    if not module_exists("fitz"):
        raise ImportError("未安装 PyMuPDF/fitz")
    import fitz  # type: ignore

    parts: list[str] = []
    doc = fitz.open(pdf_path)
    try:
        start_index, end_index = compute_page_range(doc.page_count, start_page, max_pages)
        for index in range(start_index, end_index + 1):
            page = doc.load_page(index)
            parts.append(f"\n\n<!-- page {index + 1} -->\n\n")
            parts.append(page.get_text("text"))
    finally:
        doc.close()
    return "".join(parts)


def _extract_pdf_pages_pymupdf(pdf_path: Path, start_page: int, max_pages: int | None) -> list[dict[str, object]]:
    if not module_exists("fitz"):
        raise ImportError("未安装 PyMuPDF/fitz")
    import fitz  # type: ignore

    pages: list[dict[str, object]] = []
    doc = fitz.open(pdf_path)
    try:
        start_index, end_index = compute_page_range(doc.page_count, start_page, max_pages)
        for index in range(start_index, end_index + 1):
            page = doc.load_page(index)
            pages.append({"page": index + 1, "text": page.get_text("text") or ""})
    finally:
        doc.close()
    return pages


def _extract_pdf_pypdf(pdf_path: Path, start_page: int, max_pages: int | None) -> str:
    reader_cls = None
    if module_exists("pypdf"):
        from pypdf import PdfReader  # type: ignore
        reader_cls = PdfReader
    elif module_exists("PyPDF2"):
        from PyPDF2 import PdfReader  # type: ignore
        reader_cls = PdfReader
    else:
        raise ImportError("未安装 pypdf 或 PyPDF2")

    reader = reader_cls(str(pdf_path))
    start_index, end_index = compute_page_range(len(reader.pages), start_page, max_pages)
    parts: list[str] = []
    for index in range(start_index, end_index + 1):
        parts.append(f"\n\n<!-- page {index + 1} -->\n\n")
        parts.append(reader.pages[index].extract_text() or "")
    return "".join(parts)


def _extract_pdf_pages_pypdf(pdf_path: Path, start_page: int, max_pages: int | None) -> list[dict[str, object]]:
    reader_cls = None
    if module_exists("pypdf"):
        from pypdf import PdfReader  # type: ignore
        reader_cls = PdfReader
    elif module_exists("PyPDF2"):
        from PyPDF2 import PdfReader  # type: ignore
        reader_cls = PdfReader
    else:
        raise ImportError("未安装 pypdf 或 PyPDF2")

    reader = reader_cls(str(pdf_path))
    start_index, end_index = compute_page_range(len(reader.pages), start_page, max_pages)
    return [
        {"page": index + 1, "text": reader.pages[index].extract_text() or ""}
        for index in range(start_index, end_index + 1)
    ]


def _extract_pdf_pdfplumber(pdf_path: Path, start_page: int, max_pages: int | None) -> str:
    if not module_exists("pdfplumber"):
        raise ImportError("未安装 pdfplumber")
    import pdfplumber  # type: ignore

    parts: list[str] = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        start_index, end_index = compute_page_range(len(pdf.pages), start_page, max_pages)
        for index in range(start_index, end_index + 1):
            parts.append(f"\n\n<!-- page {index + 1} -->\n\n")
            parts.append(pdf.pages[index].extract_text() or "")
    return "".join(parts)


def _extract_pdf_pages_pdfplumber(pdf_path: Path, start_page: int, max_pages: int | None) -> list[dict[str, object]]:
    if not module_exists("pdfplumber"):
        raise ImportError("未安装 pdfplumber")
    import pdfplumber  # type: ignore

    pages: list[dict[str, object]] = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        start_index, end_index = compute_page_range(len(pdf.pages), start_page, max_pages)
        for index in range(start_index, end_index + 1):
            pages.append({"page": index + 1, "text": pdf.pages[index].extract_text() or ""})
    return pages


def _extract_pdf_pdfminer(pdf_path: Path, start_page: int, max_pages: int | None) -> str:
    if not module_exists("pdfminer.high_level"):
        raise ImportError("未安装 pdfminer.six")
    from pdfminer.high_level import extract_text  # type: ignore

    page_count = get_pdf_page_count(pdf_path)
    start_index, end_index = compute_page_range(page_count, start_page, max_pages)
    text = extract_text(str(pdf_path), page_numbers=list(range(start_index, end_index + 1)))
    markers = "".join(
        f"\n\n<!-- page {i + 1} -->\n\n"
        for i in range(start_index, end_index + 1)
    )
    return markers + text


def _extract_pdf_pages_pdfminer(pdf_path: Path, start_page: int, max_pages: int | None) -> list[dict[str, object]]:
    if not module_exists("pdfminer.high_level"):
        raise ImportError("未安装 pdfminer.six")
    from pdfminer.high_level import extract_text  # type: ignore

    page_count = get_pdf_page_count(pdf_path)
    start_index, end_index = compute_page_range(page_count, start_page, max_pages)
    pages: list[dict[str, object]] = []
    for index in range(start_index, end_index + 1):
        text = extract_text(str(pdf_path), page_numbers=[index])
        pages.append({"page": index + 1, "text": text or ""})
    return pages


def _extract_pdf_liteparse(pdf_path: Path, start_page: int, max_pages: int | None) -> str:
    """使用 liteparse（Rust 原生绑定）解析 PDF。"""
    if not module_exists("liteparse"):
        raise ImportError("未安装 liteparse，运行: pip install liteparse")
    from liteparse import LiteParse  # type: ignore

    target_pages = None
    if max_pages is not None:
        end_page = start_page + max_pages - 1
        target_pages = f"{start_page}-{end_page}"

    parser = LiteParse(target_pages=target_pages, quiet=True, ocr_enabled=False)
    result = parser.parse(str(pdf_path))

    parts: list[str] = []
    for page in result.pages:
        parts.append(f"\n\n<!-- page {page.page_num} -->\n\n")
        parts.append(page.text or "")
    return "".join(parts)


def _extract_pdf_opendataloader(pdf_path: Path, start_page: int, max_pages: int | None) -> str:
    """使用 opendataloader-pdf（Java JAR）解析 PDF。"""
    if not module_exists("opendataloader_pdf"):
        raise ImportError("未安装 opendataloader-pdf，运行: pip install opendataloader-pdf（需要系统 Java 环境）")
    from opendataloader_pdf import convert  # type: ignore

    logger.info("opendataloader-pdf: 开始解析 %s (start_page=%s, max_pages=%s)", pdf_path.name, start_page, max_pages)

    with tempfile.TemporaryDirectory(prefix="opendataloader_") as tmpdir:
        # 构造页码范围参数
        pages_spec = None
        if max_pages is not None:
            end_page = start_page + max_pages - 1
            pages_spec = f"{start_page}-{end_page}"
            logger.debug("opendataloader-pdf: 页码范围 %s", pages_spec)
        elif start_page > 1:
            pages_spec = f"{start_page}-"
            logger.debug("opendataloader-pdf: 从第 %s 页到末尾", start_page)

        try:
            convert(
                str(pdf_path),
                output_dir=tmpdir,
                format="text",
                quiet=True,
                **({"pages": pages_spec} if pages_spec else {}),
            )
        except Exception as e:
            logger.error("opendataloader-pdf: 解析失败 - %s", e)
            raise

        # 读取输出文件（文件名包含原 PDF 名）
        output_files = list(Path(tmpdir).glob("*.txt"))
        if not output_files:
            logger.warning("opendataloader-pdf: 未生成输出文件")
            return ""

        output_file = output_files[0]
        logger.info("opendataloader-pdf: 输出文件 %s", output_file.name)
        text = output_file.read_text(encoding="utf-8", errors="replace")
        logger.info("opendataloader-pdf: 提取 %d 字符", len(text))
        return text


# ---------------------------------------------------------------------------
# Word 解析器
# ---------------------------------------------------------------------------

def _extract_word_markitdown(file_path: Path, start_page: int, max_pages: int | None) -> str:
    if not module_exists("markitdown"):
        raise ImportError("未安装 markitdown")
    from markitdown import MarkItDown  # type: ignore
    result = MarkItDown().convert(str(file_path))
    return getattr(result, "text_content", "") or ""


def _extract_word_docx(file_path: Path, start_page: int, max_pages: int | None) -> str:
    if not module_exists("docx"):
        raise ImportError("未安装 python-docx")
    from docx import Document  # type: ignore

    doc = Document(str(file_path))
    parts: list[str] = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            style_name = para.style.name if para.style else ""
            if style_name.startswith("Heading"):
                level = style_name.replace("Heading", "").strip()
                parts.append(f"\n{'#' * int(level or 1)} {text}\n")
            else:
                parts.append(text + "\n")

    for table in doc.tables:
        parts.append("\n")
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            parts.append("| " + " | ".join(cells) + " |")
        parts.append("\n")

    return "\n".join(parts)


# ---------------------------------------------------------------------------
# PPT 解析器
# ---------------------------------------------------------------------------

def _extract_ppt_markitdown(file_path: Path, start_page: int, max_pages: int | None) -> str:
    if not module_exists("markitdown"):
        raise ImportError("未安装 markitdown")
    from markitdown import MarkItDown  # type: ignore
    result = MarkItDown().convert(str(file_path))
    return getattr(result, "text_content", "") or ""


def _extract_ppt_pptx(file_path: Path, start_page: int, max_pages: int | None) -> str:
    if not module_exists("pptx"):
        raise ImportError("未安装 python-pptx")
    from pptx import Presentation  # type: ignore

    prs = Presentation(str(file_path))
    parts: list[str] = []

    for slide_num, slide in enumerate(prs.slides, 1):
        parts.append(f"\n\n<!-- slide {slide_num} -->\n\n")
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    text = para.text.strip()
                    if text:
                        parts.append(text)
            if shape.has_table:
                table = shape.table
                for row in table.rows:
                    cells = [cell.text.strip() for cell in row.cells]
                    parts.append("| " + " | ".join(cells) + " |")

    return "\n".join(parts)


# ---------------------------------------------------------------------------
# Excel 解析器
# ---------------------------------------------------------------------------

def _extract_excel_markitdown(file_path: Path, start_page: int, max_pages: int | None) -> str:
    if not module_exists("markitdown"):
        raise ImportError("未安装 markitdown")
    from markitdown import MarkItDown  # type: ignore
    result = MarkItDown().convert(str(file_path))
    return getattr(result, "text_content", "") or ""


def _extract_excel_openpyxl(file_path: Path, start_page: int, max_pages: int | None) -> str:
    if not module_exists("openpyxl"):
        raise ImportError("未安装 openpyxl")
    from openpyxl import load_workbook  # type: ignore

    wb = load_workbook(str(file_path), read_only=True, data_only=True)
    parts: list[str] = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        parts.append(f"\n\n## {sheet_name}\n\n")
        for row in ws.iter_rows(values_only=True):
            cells = [str(cell) if cell is not None else "" for cell in row]
            parts.append("| " + " | ".join(cells) + " |")

    wb.close()
    return "\n".join(parts)


# ---------------------------------------------------------------------------
# HTML 解析器
# ---------------------------------------------------------------------------

def _extract_html_markitdown(file_path: Path, start_page: int, max_pages: int | None) -> str:
    if not module_exists("markitdown"):
        raise ImportError("未安装 markitdown")
    from markitdown import MarkItDown  # type: ignore
    result = MarkItDown().convert(str(file_path))
    return getattr(result, "text_content", "") or ""


def _extract_html_bs4(file_path: Path, start_page: int, max_pages: int | None) -> str:
    if not module_exists("bs4"):
        raise ImportError("未安装 beautifulsoup4")
    from bs4 import BeautifulSoup  # type: ignore

    html = file_path.read_text(encoding="utf-8", errors="ignore")
    soup = BeautifulSoup(html, "html.parser")

    for tag in soup(["script", "style"]):
        tag.decompose()

    parts: list[str] = []
    for element in soup.find_all(["h1", "h2", "h3", "h4", "h5", "h6", "p", "li", "td", "th", "pre", "blockquote"]):
        text = element.get_text(strip=True)
        if not text:
            continue
        tag = element.name
        if tag.startswith("h"):
            level = int(tag[1])
            parts.append(f"\n{'#' * level} {text}\n")
        elif tag == "li":
            parts.append(f"- {text}")
        elif tag in ("td", "th"):
            parts.append(f"| {text}", )
        elif tag == "pre":
            parts.append(f"\n```\n{text}\n```\n")
        elif tag == "blockquote":
            parts.append(f"> {text}")
        else:
            parts.append(text + "\n")

    return "\n".join(parts)


# ---------------------------------------------------------------------------
# 文本格式解析器（TXT, CSV, JSON, XML, MD）
# ---------------------------------------------------------------------------

def _extract_text_markitdown(file_path: Path, start_page: int, max_pages: int | None) -> str:
    """使用 markitdown 解析文本格式（TXT, CSV, JSON, XML, MD 等）。"""
    if not module_exists("markitdown"):
        raise ImportError("未安装 markitdown")
    from markitdown import MarkItDown  # type: ignore

    logger.info("markitdown: 解析文本文件 %s", file_path.name)
    result = MarkItDown().convert(str(file_path))
    text = getattr(result, "text_content", "") or ""
    logger.info("markitdown: 提取 %d 字符", len(text))
    return text


# ---------------------------------------------------------------------------
# 图片解析器（JPG, PNG, GIF, BMP, TIFF, WebP）
# ---------------------------------------------------------------------------

def _extract_image_markitdown(file_path: Path, start_page: int, max_pages: int | None) -> str:
    """使用 markitdown 解析图片（EXIF 元数据 + OCR 文字识别）。"""
    if not module_exists("markitdown"):
        raise ImportError("未安装 markitdown")
    from markitdown import MarkItDown  # type: ignore

    logger.info("markitdown: 解析图片 %s (EXIF + OCR)", file_path.name)
    result = MarkItDown().convert(str(file_path))
    text = getattr(result, "text_content", "") or ""
    logger.info("markitdown: 提取 %d 字符", len(text))
    return text


# ---------------------------------------------------------------------------
# 音频解析器（MP3, WAV）
# ---------------------------------------------------------------------------

def _extract_audio_markitdown(file_path: Path, start_page: int, max_pages: int | None) -> str:
    """使用 markitdown 解析音频（EXIF 元数据 + 语音转录）。"""
    if not module_exists("markitdown"):
        raise ImportError("未安装 markitdown")
    from markitdown import MarkItDown  # type: ignore

    logger.info("markitdown: 解析音频 %s (EXIF + 语音转录)", file_path.name)
    result = MarkItDown().convert(str(file_path))
    text = getattr(result, "text_content", "") or ""
    logger.info("markitdown: 提取 %d 字符", len(text))
    return text


# ---------------------------------------------------------------------------
# 电子书解析器（EPub）
# ---------------------------------------------------------------------------

def _extract_epub_markitdown(file_path: Path, start_page: int, max_pages: int | None) -> str:
    """使用 markitdown 解析 EPub 电子书。"""
    if not module_exists("markitdown"):
        raise ImportError("未安装 markitdown")
    from markitdown import MarkItDown  # type: ignore

    logger.info("markitdown: 解析 EPub %s", file_path.name)
    result = MarkItDown().convert(str(file_path))
    text = getattr(result, "text_content", "") or ""
    logger.info("markitdown: 提取 %d 字符", len(text))
    return text


# ---------------------------------------------------------------------------
# 压缩包解析器（ZIP）
# ---------------------------------------------------------------------------

def _extract_archive_markitdown(file_path: Path, start_page: int, max_pages: int | None) -> str:
    """使用 markitdown 解析 ZIP 压缩包（遍历内部内容）。"""
    if not module_exists("markitdown"):
        raise ImportError("未安装 markitdown")
    from markitdown import MarkItDown  # type: ignore

    logger.info("markitdown: 解析压缩包 %s (遍历内部文件)", file_path.name)
    result = MarkItDown().convert(str(file_path))
    text = getattr(result, "text_content", "") or ""
    logger.info("markitdown: 提取 %d 字符", len(text))
    return text


# ---------------------------------------------------------------------------
# 表格提取（PDF 专用）
# ---------------------------------------------------------------------------

def extract_tables_from_pdf(pdf_path: Path, page_indices: list[int]) -> list[dict]:
    """从 PDF 提取表格，返回 [{page, table_index, headers, rows}, ...]。"""
    if not module_exists("pdfplumber"):
        raise ImportError("表格提取需要 pdfplumber，运行: pip install pdfplumber")
    import pdfplumber  # type: ignore

    tables: list[dict] = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page_idx in page_indices:
            if page_idx >= len(pdf.pages):
                continue
            page = pdf.pages[page_idx]
            page_tables = page.extract_tables()
            for tbl_idx, tbl in enumerate(page_tables):
                if not tbl:
                    continue
                # 第一行作为表头
                headers = [str(h).strip() if h else "" for h in tbl[0]]
                rows = []
                for row in tbl[1:]:
                    rows.append([str(cell).strip() if cell else "" for cell in row])
                tables.append({
                    "page": page_idx + 1,
                    "table_index": tbl_idx + 1,
                    "headers": headers,
                    "rows": rows,
                })
    return tables


def tables_to_markdown(tables: list[dict]) -> str:
    """将表格列表转为 Markdown 格式。"""
    parts: list[str] = []
    for tbl in tables:
        parts.append(f"\n### Page {tbl['page']} - Table {tbl['table_index']}\n")
        parts.append("| " + " | ".join(tbl["headers"]) + " |")
        parts.append("|" + "|".join(["---"] * len(tbl["headers"])) + "|")
        for row in tbl["rows"]:
            parts.append("| " + " | ".join(row) + " |")
        parts.append("")
    return "\n".join(parts)


def tables_to_csv(tables: list[dict]) -> str:
    """将表格列表转为 CSV 格式（多个表格用空行分隔）。"""
    output = io.StringIO()
    writer = csv.writer(output)
    for i, tbl in enumerate(tables):
        if i > 0:
            writer.writerow([])  # 空行分隔
        writer.writerow([f"# Page {tbl['page']} Table {tbl['table_index']}"])
        writer.writerow(tbl["headers"])
        for row in tbl["rows"]:
            writer.writerow(row)
    return output.getvalue()


def tables_to_json(tables: list[dict]) -> str:
    """Render extracted tables as JSON."""
    return json.dumps(
        {
            "version": __version__,
            "format": "tables",
            "table_count": len(tables),
            "tables": tables,
        },
        ensure_ascii=False,
        indent=2,
    )


# ---------------------------------------------------------------------------
# 解析器注册表
# ---------------------------------------------------------------------------

_EXTRACTORS: dict[tuple[str, str], Callable[[Path, int, int | None], str]] = {
    # PDF
    ("pdf", "markitdown"):   _extract_pdf_markitdown,
    ("pdf", "pymupdf"):      _extract_pdf_pymupdf,
    ("pdf", "pypdf"):        _extract_pdf_pypdf,
    ("pdf", "pdfplumber"):   _extract_pdf_pdfplumber,
    ("pdf", "pdfminer"):     _extract_pdf_pdfminer,
    ("pdf", "liteparse"):    _extract_pdf_liteparse,
    ("pdf", "opendataloader"): _extract_pdf_opendataloader,
    # Word
    ("word", "markitdown"):  _extract_word_markitdown,
    ("word", "python-docx"): _extract_word_docx,
    # PPT
    ("ppt", "markitdown"):   _extract_ppt_markitdown,
    ("ppt", "python-pptx"):  _extract_ppt_pptx,
    # Excel
    ("excel", "markitdown"): _extract_excel_markitdown,
    ("excel", "openpyxl"):   _extract_excel_openpyxl,
    # HTML
    ("html", "markitdown"):      _extract_html_markitdown,
    ("html", "beautifulsoup4"):  _extract_html_bs4,
    # 文本格式 (TXT, CSV, JSON, XML, MD)
    ("text", "markitdown"):  _extract_text_markitdown,
    # 图片格式 (JPG, PNG, GIF, BMP, TIFF, WebP)
    ("image", "markitdown"): _extract_image_markitdown,
    # 音频格式 (MP3, WAV)
    ("audio", "markitdown"): _extract_audio_markitdown,
    # 电子书 (EPub)
    ("epub", "markitdown"):  _extract_epub_markitdown,
    # 压缩包 (ZIP)
    ("archive", "markitdown"): _extract_archive_markitdown,
}

_PDF_PAGE_EXTRACTORS: dict[str, Callable[[Path, int, int | None], list[dict[str, object]]]] = {
    "pymupdf": _extract_pdf_pages_pymupdf,
    "pypdf": _extract_pdf_pages_pypdf,
    "pdfplumber": _extract_pdf_pages_pdfplumber,
    "pdfminer": _extract_pdf_pages_pdfminer,
}

_PARSER_ALIASES: dict[str, str] = {
    "fitz": "pymupdf",
    "pypdf2": "pypdf",
    "docx": "python-docx",
    "pptx": "python-pptx",
    "bs4": "beautifulsoup4",
    "llama-parse": "liteparse",
    "odl": "opendataloader",
    "opendataloader-pdf": "opendataloader",
}


def resolve_parser_name(name: str) -> str:
    """解析器名称标准化，处理别名。"""
    return _PARSER_ALIASES.get(name, name)


def get_extractor(fmt: str, parser_name: str) -> Callable[[Path, int, int | None], str] | None:
    """获取指定格式和解析器的提取函数。"""
    return _EXTRACTORS.get((fmt, parser_name))


def get_pdf_page_extractor(parser_name: str) -> Callable[[Path, int, int | None], list[dict[str, object]]] | None:
    """Return page-by-page PDF extractor when available."""
    return _PDF_PAGE_EXTRACTORS.get(parser_name)


def list_all_parsers() -> list[str]:
    """列出所有已注册的解析器名称（去重）。"""
    names: set[str] = set()
    for _, parser_name in _EXTRACTORS:
        names.add(parser_name)
    return sorted(names)


# ---------------------------------------------------------------------------
# PDF 子集生成（仅 PDF 格式需要）
# ---------------------------------------------------------------------------

def create_subset_pdf_with_fitz(pdf_path: Path, start_page: int, max_pages: int, tmp_dir: Path) -> Path | None:
    if not module_exists("fitz"):
        return None
    import fitz  # type: ignore

    source = fitz.open(pdf_path)
    try:
        start_index, end_index = compute_page_range(source.page_count, start_page, max_pages)
        subset = fitz.open()
        subset.insert_pdf(source, from_page=start_index, to_page=end_index)
        subset_path = tmp_dir / f"{pdf_path.stem}.pages_{start_index + 1}_to_{end_index + 1}.pdf"
        subset.save(subset_path)
        subset.close()
        return subset_path
    finally:
        source.close()


def create_subset_pdf_with_pypdf(pdf_path: Path, start_page: int, max_pages: int, tmp_dir: Path) -> Path | None:
    reader_cls = None
    writer_cls = None

    if module_exists("pypdf"):
        from pypdf import PdfReader, PdfWriter  # type: ignore
        reader_cls = PdfReader
        writer_cls = PdfWriter
    elif module_exists("PyPDF2"):
        from PyPDF2 import PdfReader, PdfWriter  # type: ignore
        reader_cls = PdfReader
        writer_cls = PdfWriter
    else:
        return None

    reader = reader_cls(str(pdf_path))
    writer = writer_cls()
    start_index, end_index = compute_page_range(len(reader.pages), start_page, max_pages)
    for index in range(start_index, end_index + 1):
        writer.add_page(reader.pages[index])

    subset_path = tmp_dir / f"{pdf_path.stem}.pages_{start_index + 1}_to_{end_index + 1}.pdf"
    with subset_path.open("wb") as fh:
        writer.write(fh)
    return subset_path


def create_subset_pdf(pdf_path: Path, start_page: int, max_pages: int, tmp_dir: Path) -> Path | None:
    for creator_name, creator in [
        ("fitz", create_subset_pdf_with_fitz),
        ("pypdf", create_subset_pdf_with_pypdf),
    ]:
        try:
            subset = creator(pdf_path, start_page, max_pages, tmp_dir)
            if subset:
                return subset
        except Exception as exc:
            logger.debug("使用 %s 创建子 PDF 失败: %s", creator_name, exc)
    return None


# ---------------------------------------------------------------------------
# 相似度计算
# ---------------------------------------------------------------------------

def similarity_score(a: str, b: str, sample_chars: int) -> float:
    a_sample = a[:sample_chars]
    b_sample = b[:sample_chars]
    if not a_sample and not b_sample:
        return 1.0
    if not a_sample or not b_sample:
        return 0.0
    return round(difflib.SequenceMatcher(None, a_sample, b_sample, autojunk=True).ratio(), 4)


def choose_recommended_result(results: list[ParseResult]) -> ParseResult | None:
    """Choose a likely best parser result using quality first, then text volume and speed."""
    ok_results = [r for r in results if r.status == "ok" and r.non_space_chars > 0]
    if not ok_results:
        return None
    return max(ok_results, key=lambda r: (r.quality_score, r.non_space_chars, -r.seconds))


def build_diagnostics(results: list[ParseResult]) -> list[str]:
    """Build concise diagnostics for extraction quality issues."""
    ok_results = [r for r in results if r.status == "ok"]
    diagnostics: list[str] = []
    if not ok_results:
        return ["No parser produced usable output."]

    bad_or_empty = [r for r in ok_results if r.quality_label in ("empty", "bad")]
    if len(bad_or_empty) == len(ok_results):
        diagnostics.append("All successful parser outputs look empty or poor; OCR or a cleaner source file is likely needed.")

    if any(r.cid_markers > 0 for r in ok_results):
        diagnostics.append("One or more outputs contain `(cid:...)` markers, which usually indicates embedded font mapping problems.")

    if any(r.control_chars > max(20, r.non_space_chars * 0.05) for r in ok_results):
        diagnostics.append("One or more outputs contain many control characters, so visual searchability may not equal extractable text quality.")

    if not diagnostics:
        diagnostics.append("No obvious extraction-quality warning was detected from text-only heuristics.")

    return diagnostics


def ocr_fallback_notice(enabled: bool) -> str | None:
    """Return a clear OCR fallback notice."""
    if not enabled:
        return None
    return "OCR fallback was requested. Use the `ocr` command when ordinary extraction is bad or empty; OCR requires PyMuPDF, Pillow, pytesseract, and the system Tesseract binary."


def render_pdf_pages(pdf_path: Path, page_indices: list[int], out_dir: Path, dpi: int = 200) -> list[dict[str, object]]:
    """Render selected PDF pages to PNG files using PyMuPDF."""
    if not module_exists("fitz"):
        raise ImportError("页面截图需要 PyMuPDF/fitz，运行: pip install pymupdf")
    import fitz  # type: ignore

    out_dir.mkdir(parents=True, exist_ok=True)
    rendered: list[dict[str, object]] = []
    zoom = dpi / 72.0
    matrix = fitz.Matrix(zoom, zoom)
    doc = fitz.open(pdf_path)
    try:
        for page_idx in page_indices:
            if page_idx < 0 or page_idx >= doc.page_count:
                continue
            page = doc.load_page(page_idx)
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            out_path = out_dir / f"{pdf_path.stem}_page_{page_idx + 1:03d}.png"
            pix.save(out_path)
            rendered.append(
                {
                    "page": page_idx + 1,
                    "file": str(out_path),
                    "width": pix.width,
                    "height": pix.height,
                    "dpi": dpi,
                }
            )
    finally:
        doc.close()
    return rendered


def ensure_tesseract_available() -> None:
    """Validate OCR dependencies and raise a readable setup error."""
    missing_python = []
    if not module_exists("pytesseract"):
        missing_python.append("pytesseract")
    if not module_exists("PIL"):
        missing_python.append("pillow")
    if missing_python:
        raise ImportError(
            "OCR 需要 Python 依赖："
            + ", ".join(missing_python)
            + "。安装示例：pip install pytesseract pillow"
        )
    if shutil.which("tesseract") is None:
        raise RuntimeError("OCR 需要系统 Tesseract 可执行文件，请安装 tesseract-ocr 并加入 PATH")


def ocr_install_hint(missing_python: list[str], missing_system: list[str]) -> str | None:
    """Build a concise OCR install hint without empty commands."""
    hints: list[str] = []
    if missing_python:
        hints.append(f"pip install {' '.join(missing_python)}")
    if missing_system:
        hints.append("install system Tesseract OCR")
    return "; ".join(hints) if hints else None


def opendataloader_install_hint(missing_python: list[str], missing_system: list[str]) -> str | None:
    """Build a concise opendataloader-pdf install hint without empty commands."""
    hints: list[str] = []
    if missing_python:
        hints.append(f"pip install {' '.join(missing_python)}")
    if missing_system:
        hints.append("install JDK/JRE and ensure java is in PATH")
    return "; ".join(hints) if hints else None


def ocr_pdf_pages(pdf_path: Path, page_indices: list[int], lang: str, dpi: int) -> list[dict[str, object]]:
    """Run Tesseract OCR on selected PDF pages."""
    ensure_tesseract_available()
    if not module_exists("fitz"):
        raise ImportError("OCR 渲染 PDF 页面需要 PyMuPDF/fitz，运行: pip install pymupdf")

    import fitz  # type: ignore
    import pytesseract  # type: ignore
    from PIL import Image  # type: ignore

    zoom = dpi / 72.0
    matrix = fitz.Matrix(zoom, zoom)
    results: list[dict[str, object]] = []
    doc = fitz.open(pdf_path)
    try:
        for page_idx in page_indices:
            page = doc.load_page(page_idx)
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            image = Image.open(io.BytesIO(pix.tobytes("png")))
            start = time.perf_counter()
            text = pytesseract.image_to_string(image, lang=lang)
            normalized = normalize_text(text)
            quality_score, quality_label = quality_from_text(normalized)
            results.append(
                {
                    "page": page_idx + 1,
                    "seconds": round(time.perf_counter() - start, 3),
                    "chars": len(normalized),
                    "quality_score": quality_score,
                    "quality_label": quality_label,
                    "text": normalized,
                }
            )
    finally:
        doc.close()
    return results


def ocr_pages_to_text(pages: list[dict[str, object]]) -> str:
    """Join OCR page results into a page-marked text stream."""
    parts: list[str] = []
    for page in pages:
        parts.append(f"\n\n<!-- page {page.get('page')} ocr -->\n\n")
        parts.append(str(page.get("text") or ""))
    return normalize_text("\n".join(parts))


def maybe_ocr_fallback(
    file_path: Path,
    current_text: str,
    current_result: ParseResult,
    enabled: bool,
    min_quality: float,
    pages: str,
    lang: str,
    dpi: int,
) -> tuple[str, ParseResult, dict[str, object] | None]:
    """Run OCR when enabled and current quality is below the threshold."""
    if not enabled:
        return current_text, current_result, None
    if file_path.suffix.lower() != ".pdf":
        return current_text, current_result, {"attempted": False, "reason": "not_pdf"}
    if current_result.quality_score >= min_quality and current_result.quality_label not in QUALITY_BAD_LABELS:
        return current_text, current_result, {"attempted": False, "reason": "text_quality_ok"}

    try:
        page_count = get_pdf_page_count(file_path)
        page_indices = resolve_page_indices(pages, page_count, default_all=False)
        start = time.perf_counter()
        ocr_pages = ocr_pdf_pages(file_path, page_indices, lang, dpi)
        ocr_text = ocr_pages_to_text(ocr_pages)
        ocr_result = result_from_text("ocr-tesseract", ocr_text, time.perf_counter() - start)
        used = ocr_result.quality_score > current_result.quality_score or current_result.quality_label in QUALITY_BAD_LABELS
        payload = {
            "attempted": True,
            "used": used,
            "backend": "tesseract",
            "pages": [page["page"] for page in ocr_pages],
            "lang": lang,
            "dpi": dpi,
            "quality_score": ocr_result.quality_score,
            "quality_label": ocr_result.quality_label,
            "error": None,
        }
        if used:
            return ocr_text, ocr_result, payload
        return current_text, current_result, payload
    except Exception as exc:
        return current_text, current_result, {
            "attempted": True,
            "used": False,
            "backend": "tesseract",
            "pages": pages,
            "lang": lang,
            "dpi": dpi,
            "quality_score": 0.0,
            "quality_label": "failed",
            "error": repr(exc),
        }


def write_ocr_output(
    out_path: Path,
    source_file: Path,
    pages: list[dict[str, object]],
    output_format: str,
    lang: str,
    dpi: int,
) -> Path:
    """Write OCR output as txt, md, or json."""
    out_path.parent.mkdir(parents=True, exist_ok=True)
    if output_format == "json":
        payload = {
            "version": __version__,
            "source_file": str(source_file),
            "backend": "tesseract",
            "lang": lang,
            "dpi": dpi,
            "page_count": len(pages),
            "pages": pages,
        }
        out_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    elif output_format == "md":
        parts = [
            "# OCR Output",
            "",
            f"- Source: `{source_file}`",
            "- Backend: `tesseract`",
            f"- Language: `{lang}`",
            f"- DPI: `{dpi}`",
            "",
        ]
        for page in pages:
            parts.extend([f"## Page {page['page']}", "", str(page["text"]), ""])
        out_path.write_text("\n".join(parts), encoding="utf-8")
    else:
        parts = []
        for page in pages:
            parts.append(f"--- page {page['page']} ---")
            parts.append(str(page["text"]))
        out_path.write_text("\n\n".join(parts) + ("\n" if parts else ""), encoding="utf-8")
    return out_path


def pdf_layout_json(pdf_path: Path, start_page: int = 1, max_pages: int | None = None) -> dict[str, object]:
    """Extract PDF page/block/line/span layout with coordinates using PyMuPDF."""
    if not module_exists("fitz"):
        raise ImportError("layout-json 需要 PyMuPDF/fitz，运行: pip install pymupdf")
    import fitz  # type: ignore

    doc = fitz.open(pdf_path)
    pages: list[dict[str, object]] = []
    try:
        start_index, end_index = compute_page_range(doc.page_count, start_page, max_pages)
        for page_index in range(start_index, end_index + 1):
            page = doc.load_page(page_index)
            data = page.get_text("dict")
            page_payload: dict[str, object] = {
                "page": page_index + 1,
                "width": page.rect.width,
                "height": page.rect.height,
                "rotation": page.rotation,
                "blocks": [],
            }
            for block_index, block in enumerate(data.get("blocks", []), start=1):
                block_payload: dict[str, object] = {
                    "block_id": block_index,
                    "type": block.get("type"),
                    "bbox": block.get("bbox"),
                    "lines": [],
                }
                for line_index, line in enumerate(block.get("lines", []), start=1):
                    spans_payload = []
                    line_text_parts = []
                    for span_index, span in enumerate(line.get("spans", []), start=1):
                        span_text = span.get("text", "")
                        line_text_parts.append(span_text)
                        spans_payload.append(
                            {
                                "span_id": span_index,
                                "text": span_text,
                                "bbox": span.get("bbox"),
                                "font": span.get("font"),
                                "size": span.get("size"),
                                "flags": span.get("flags"),
                                "color": span.get("color"),
                            }
                        )
                    block_payload["lines"].append(
                        {
                            "line_id": line_index,
                            "bbox": line.get("bbox"),
                            "text": "".join(line_text_parts),
                            "spans": spans_payload,
                        }
                    )
                page_payload["blocks"].append(block_payload)
            pages.append(page_payload)
    finally:
        doc.close()

    return {
        "version": __version__,
        "source_file": str(pdf_path),
        "format": "pdf-layout",
        "start_page": start_page,
        "max_pages": max_pages,
        "page_count": len(pages),
        "pages": pages,
    }


def layout_to_page_map(layout: dict[str, object]) -> list[dict[str, object]]:
    """Build a lightweight page map from layout JSON."""
    page_map: list[dict[str, object]] = []
    for page in layout.get("pages", []):  # type: ignore[union-attr]
        if not isinstance(page, dict):
            continue
        texts: list[str] = []
        for block in page.get("blocks", []):
            if not isinstance(block, dict):
                continue
            for line in block.get("lines", []):
                if isinstance(line, dict) and line.get("text"):
                    texts.append(str(line["text"]))
        joined = normalize_text("\n".join(texts))
        page_map.append(
            {
                "page": page.get("page"),
                "width": page.get("width"),
                "height": page.get("height"),
                "chars": len(joined),
                "preview": joined[:500],
            }
        )
    return page_map


INVOICE_SUMMARY_COLUMNS = [
    "source_file",
    "parser",
    "quality_label",
    "quality_score",
    "invoice_type",
    "invoice_number",
    "invoice_date",
    "buyer_name",
    "buyer_tax_id",
    "seller_name",
    "seller_tax_id",
    "total_amount",
    "total_tax",
    "total_with_tax",
    "total_with_tax_cn",
    "drawer",
    "validation_status",
    "missing_fields",
]


def clean_invoice_scan_text(text: str) -> str:
    """Remove layout noise that commonly appears in invoice parser output."""
    return re.sub(r"[\s|`]+", "", text or "")


def regex_first(pattern: str, text: str, flags: int = 0) -> str | None:
    match = re.search(pattern, text, flags)
    if not match:
        return None
    value = match.group(1).strip()
    return value or None


def extract_invoice_items(text: str) -> list[dict[str, object]]:
    """Extract invoice line items from text-oriented parser output."""
    items: list[dict[str, object]] = []
    item_pattern = re.compile(
        r"(?P<name>\*[^*\n]+\*[^0-9\n|]+?)\s+"
        r"(?P<unit>[\u4e00-\u9fffA-Za-z]{1,8})\s+"
        r"(?P<quantity>[0-9]+(?:\.[0-9]+)?)\s+"
        r"(?P<unit_price>[0-9]+(?:\.[0-9]+)?)\s+"
        r"(?P<amount>[0-9]+(?:\.[0-9]{2})?)\s+"
        r"(?P<tax_rate>[0-9]+%)\s+"
        r"(?P<tax_amount>[0-9]+(?:\.[0-9]{2})?)"
    )
    for raw_line in text.splitlines():
        line = raw_line.replace("|", " ")
        line = re.sub(r"\s+", " ", line).strip()
        if not line or "项目名称" in line or "---" in line:
            continue
        match = item_pattern.search(line)
        if not match:
            continue
        data = match.groupdict()
        items.append(
            {
                "name": data["name"].strip(),
                "unit": data["unit"],
                "quantity": decimal_from_text(data["quantity"]),
                "unit_price": decimal_from_text(data["unit_price"]),
                "amount": decimal_from_text(data["amount"]),
                "tax_rate": data["tax_rate"],
                "tax_amount": decimal_from_text(data["tax_amount"]),
            }
        )
    return items


def validate_invoice_fields(fields: dict[str, object], strict: bool = False) -> dict[str, object]:
    """Run invoice consistency checks."""
    required = [
        "invoice_number",
        "invoice_date",
        "buyer_name",
        "buyer_tax_id",
        "seller_name",
        "seller_tax_id",
        "total_with_tax",
    ]
    missing = [key for key in required if not fields.get(key)]
    checks: list[dict[str, object]] = []

    total_amount = fields.get("total_amount")
    total_tax = fields.get("total_tax")
    total_with_tax = fields.get("total_with_tax")
    if isinstance(total_amount, float) and isinstance(total_tax, float) and isinstance(total_with_tax, float):
        expected = round(total_amount + total_tax, 2)
        passed = abs(expected - total_with_tax) <= 0.01
        checks.append(
            {
                "name": "amount_plus_tax_equals_total",
                "passed": passed,
                "expected": expected,
                "actual": total_with_tax,
            }
        )

    items = fields.get("items")
    if isinstance(items, list) and items:
        item_amount = round(sum(float(item.get("amount") or 0.0) for item in items if isinstance(item, dict)), 2)
        item_tax = round(sum(float(item.get("tax_amount") or 0.0) for item in items if isinstance(item, dict)), 2)
        if isinstance(total_amount, float):
            checks.append(
                {
                    "name": "item_amount_sum_equals_total_amount",
                    "passed": abs(item_amount - total_amount) <= 0.01,
                    "expected": item_amount,
                    "actual": total_amount,
                }
            )
        if isinstance(total_tax, float):
            checks.append(
                {
                    "name": "item_tax_sum_equals_total_tax",
                    "passed": abs(item_tax - total_tax) <= 0.01,
                    "expected": item_tax,
                    "actual": total_tax,
                }
            )

    if strict:
        if not isinstance(items, list) or not items:
            checks.append({"name": "has_line_items", "passed": False, "expected": ">=1", "actual": 0})
        if fields.get("invoice_number") and not re.fullmatch(r"[0-9]{8,30}", str(fields["invoice_number"])):
            checks.append({"name": "invoice_number_format", "passed": False, "expected": "8-30 digits", "actual": fields["invoice_number"]})
        for key in ("buyer_tax_id", "seller_tax_id"):
            value = fields.get(key)
            if value and not re.fullmatch(r"[0-9A-Z]{15,25}", str(value)):
                checks.append({"name": f"{key}_format", "passed": False, "expected": "15-25 alnum", "actual": value})

    status = "ok"
    if missing:
        status = "missing_fields"
    if any(not check["passed"] for check in checks):
        status = "failed_checks"
    return {"status": status, "missing_fields": missing, "checks": checks}


def extract_invoice_fields_from_text(
    text: str,
    source_file: Path | None = None,
    parser_name: str | None = None,
    result: ParseResult | None = None,
    strict: bool = False,
) -> dict[str, object]:
    """Extract structured fields from a Chinese electronic invoice."""
    normalized = normalize_text(text)
    scan = clean_invoice_scan_text(normalized)

    tax_ids = re.findall(r"统一社会信用代码/纳税人识别号[:：]?([0-9A-Z]{15,25})", scan)
    amount_pairs = re.findall(r"合计[¥￥]?([0-9]+(?:\.[0-9]{2}))[¥￥]?([0-9]+(?:\.[0-9]{2}))", scan)
    total_amount = decimal_from_text(amount_pairs[-1][0]) if amount_pairs else None
    total_tax = decimal_from_text(amount_pairs[-1][1]) if amount_pairs else None
    drawer = regex_first(r"开票人[:：]?\s*([^\s\n|]+)", normalized)

    fields: dict[str, object] = {
        "version": __version__,
        "profile": "invoice",
        "source_file": str(source_file) if source_file else None,
        "parser": parser_name,
        "quality_score": result.quality_score if result else None,
        "quality_label": result.quality_label if result else None,
        "invoice_type": regex_first(r"(电子发票[（(][^）)]+[）)]|增值税电子普通发票|增值税专用发票|增值税普通发票)", scan),
        "invoice_number": regex_first(r"发票号码[:：]?([0-9]{8,30})", scan),
        "invoice_date": regex_first(r"开票日期[:：]?([0-9]{4}年[0-9]{1,2}月[0-9]{1,2}日|[0-9]{4}[-/][0-9]{1,2}[-/][0-9]{1,2})", scan),
        "buyer_name": regex_first(r"购.*?名称[:：]?(.+?)(?:销|买售|统一社会信用代码)", scan),
        "buyer_tax_id": tax_ids[0] if len(tax_ids) >= 1 else None,
        "seller_name": regex_first(r"销.*?名称[:：]?(.+?)(?:买售|售方|方信|统一社会信用代码|项目名称|息)", scan),
        "seller_tax_id": tax_ids[1] if len(tax_ids) >= 2 else None,
        "total_amount": total_amount,
        "total_tax": total_tax,
        "total_with_tax": decimal_from_text(regex_first(r"[（(]小写[）)][¥￥]?([0-9]+(?:\.[0-9]{2}))", scan)),
        "total_with_tax_cn": regex_first(r"价税合计[（(]大写[）)](.+?)[（(]小写[）)]", scan),
        "drawer": drawer or regex_first(r"开票人[:：]?([\u4e00-\u9fffA-Za-z0-9·]{1,30})", scan),
        "items": extract_invoice_items(normalized),
        "raw_text_chars": len(normalized),
    }
    fields["validation"] = validate_invoice_fields(fields, strict=strict)
    return fields


def score_invoice_fields(fields: dict[str, object]) -> int:
    """Score invoice extraction completeness for choosing among parsers."""
    score = 0
    for key in [
        "invoice_number",
        "invoice_date",
        "buyer_name",
        "buyer_tax_id",
        "seller_name",
        "seller_tax_id",
        "total_with_tax",
        "drawer",
    ]:
        if fields.get(key):
            score += 2
    if fields.get("items"):
        score += 4
    validation = fields.get("validation")
    if isinstance(validation, dict) and validation.get("status") == "ok":
        score += 3
    return score


def extract_invoice_fields_from_file(
    file_path: Path,
    parser_name: str = "auto",
    start_page: int = 1,
    max_pages: int | None = None,
    strict: bool = False,
) -> dict[str, object]:
    """Extract invoice fields from a file, trying multiple parsers in auto mode."""
    fmt = detect_format(file_path)
    if fmt is None:
        raise ValueError(f"不支持的文件格式：{file_path.suffix}")

    candidates = [resolve_parser_name(parser_name)] if parser_name != "auto" else get_parsers_for_format(fmt)
    extracted: list[dict[str, object]] = []
    errors: list[str] = []
    for candidate in candidates:
        start = time.perf_counter()
        try:
            raw_text = extract_text_with_parser(file_path, candidate, start_page, max_pages)
            normalized = normalize_text(raw_text)
            result = text_stats(candidate, "ok", time.perf_counter() - start, normalized, None)
            fields = extract_invoice_fields_from_text(normalized, file_path, candidate, result, strict=strict)
            fields["extraction_seconds"] = result.seconds
            extracted.append(fields)
        except Exception as exc:
            errors.append(f"{candidate}: {exc}")

    if not extracted:
        raise RuntimeError("发票字段抽取失败：" + "; ".join(errors))
    return max(
        extracted,
        key=lambda item: (
            score_invoice_fields(item),
            float(item.get("quality_score") or 0.0),
            int(item.get("raw_text_chars") or 0),
        ),
    )


def classify_document(text: str, metadata: dict[str, object] | None = None) -> dict[str, object]:
    """Classify documents for document-intelligence workflows."""
    scan = clean_invoice_scan_text(text)
    invoice_hits = sum(1 for token in ["发票号码", "开票日期", "购买方", "销售方", "价税合计", "税率"] if token in scan)
    contract_hits = sum(1 for token in ["合同", "甲方", "乙方", "签订", "违约", "协议"] if token in scan)
    textbook_hits = sum(1 for token in ["目录", "第", "章", "练习", "考试", "知识点"] if token in scan)
    notice_hits = sum(1 for token in ["通知", "公告", "请", "附件", "日期"] if token in scan)

    fmt = metadata.get("format") if isinstance(metadata, dict) else None
    pdf_meta = metadata.get("pdf") if isinstance(metadata, dict) else None
    has_text_layer = None
    page_count = None
    if isinstance(pdf_meta, dict):
        has_text_layer = pdf_meta.get("has_sample_text_layer")
        page_count = pdf_meta.get("page_count")

    scores = {
        "invoice": invoice_hits / 6,
        "contract": contract_hits / 6,
        "textbook": textbook_hits / 6,
        "notice": notice_hits / 5,
    }
    best_profile, best_score = max(scores.items(), key=lambda item: item[1])
    if invoice_hits >= 3:
        best_profile = "invoice"
        best_score = scores["invoice"]
    elif fmt == "pdf" and has_text_layer is False:
        best_profile = "scanned_pdf"
        best_score = 0.8
    elif fmt == "pdf" and not compact_text(text):
        best_profile = "image_or_empty_pdf"
        best_score = 0.7
    elif best_score < 0.34:
        best_profile = "generic"
        best_score = 0.5

    strategy = {
        "invoice": "extract_fields_and_verify",
        "contract": "chunk_with_page_references",
        "textbook": "chunk_by_page_and_build_knowledge_pack",
        "notice": "extract_text_and_metadata",
        "scanned_pdf": "ocr_then_chunk",
        "image_or_empty_pdf": "ocr_required",
        "generic": "extract_best_text_and_chunk",
    }.get(best_profile, "extract_best_text_and_chunk")

    return {
        "profile": best_profile,
        "confidence": round(min(1.0, best_score), 3),
        "strategy": strategy,
        "signals": {
            "invoice_hits": invoice_hits,
            "contract_hits": contract_hits,
            "textbook_hits": textbook_hits,
            "notice_hits": notice_hits,
            "format": fmt,
            "page_count": page_count,
            "has_sample_text_layer": has_text_layer,
        },
    }


def write_invoice_fields(fields: dict[str, object], out_path: Path, output_format: str) -> Path:
    """Write invoice fields as JSON or Markdown."""
    out_path.parent.mkdir(parents=True, exist_ok=True)
    if output_format == "json":
        out_path.write_text(json.dumps(fields, ensure_ascii=False, indent=2), encoding="utf-8")
    else:
        validation = fields.get("validation") if isinstance(fields.get("validation"), dict) else {}
        lines = [
            "# Invoice Fields",
            "",
            f"- Source: `{fields.get('source_file')}`",
            f"- Parser: `{fields.get('parser')}`",
            f"- Quality: `{fields.get('quality_score')}` (`{fields.get('quality_label')}`)",
            f"- Validation: `{validation.get('status')}`",
            "",
            "| Field | Value |",
            "|---|---|",
        ]
        for key in INVOICE_SUMMARY_COLUMNS[4:16]:
            lines.append(f"| `{key}` | {fields.get(key) or ''} |")
        lines.extend(["", "## Items", ""])
        items = fields.get("items")
        if isinstance(items, list) and items:
            lines.extend(["| Name | Unit | Quantity | Unit Price | Amount | Tax Rate | Tax Amount |", "|---|---|---:|---:|---:|---|---:|"])
            for item in items:
                if isinstance(item, dict):
                    lines.append(
                        f"| {item.get('name') or ''} | {item.get('unit') or ''} | {item.get('quantity') or ''} | "
                        f"{item.get('unit_price') or ''} | {item.get('amount') or ''} | {item.get('tax_rate') or ''} | {item.get('tax_amount') or ''} |"
                    )
        else:
            lines.append("No line items extracted.")
        out_path.write_text("\n".join(lines), encoding="utf-8")
    return out_path


def invoice_summary_row(fields: dict[str, object]) -> dict[str, object]:
    """Flatten invoice fields for tabular exports."""
    validation = fields.get("validation") if isinstance(fields.get("validation"), dict) else {}
    missing = validation.get("missing_fields") if isinstance(validation, dict) else []
    row = {key: fields.get(key) for key in INVOICE_SUMMARY_COLUMNS}
    row["validation_status"] = validation.get("status") if isinstance(validation, dict) else None
    row["missing_fields"] = ",".join(missing) if isinstance(missing, list) else ""
    return row


def collect_input_files(path: Path, extensions: set[str], recursive: bool) -> list[Path]:
    """Collect one input file or files under a directory."""
    if path.is_file():
        return [path]
    if not path.exists() or not path.is_dir():
        raise FileNotFoundError(f"路径不存在：{path}")
    files: list[Path] = []
    for ext in sorted(extensions):
        iterator = path.rglob(f"*{ext}") if recursive else path.glob(f"*{ext}")
        files.extend(item for item in iterator if item.is_file())
    return sorted(set(files))


def write_invoice_xlsx(rows: list[dict[str, object]], out_path: Path) -> Path:
    """Write invoice summary and line items to XLSX."""
    if not module_exists("openpyxl"):
        raise ImportError("导出 XLSX 需要 openpyxl，运行: pip install openpyxl")
    from openpyxl import Workbook  # type: ignore

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "invoices"
    ws.append(INVOICE_SUMMARY_COLUMNS)
    for fields in rows:
        flat = invoice_summary_row(fields)
        ws.append([flat.get(column) for column in INVOICE_SUMMARY_COLUMNS])

    item_ws = wb.create_sheet("items")
    item_columns = ["source_file", "invoice_number", "name", "unit", "quantity", "unit_price", "amount", "tax_rate", "tax_amount"]
    item_ws.append(item_columns)
    for fields in rows:
        items = fields.get("items")
        if not isinstance(items, list):
            continue
        for item in items:
            if not isinstance(item, dict):
                continue
            item_ws.append([
                fields.get("source_file"),
                fields.get("invoice_number"),
                item.get("name"),
                item.get("unit"),
                item.get("quantity"),
                item.get("unit_price"),
                item.get("amount"),
                item.get("tax_rate"),
                item.get("tax_amount"),
            ])

    wb.save(out_path)
    return out_path


def write_json_file(path: Path, payload: object) -> Path:
    """Write a JSON file with UTF-8 encoding."""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def write_knowledge_pack(
    file_path: Path,
    out_dir: Path,
    parser_name: str = "auto",
    start_page: int = 1,
    max_pages: int | None = None,
    chunk_by: str = "page",
    chunk_size: int = 2000,
    overlap: int = 200,
    min_quality: float = 0.5,
) -> dict[str, object]:
    """Generate a traceable RAG/knowledge package for one document."""
    out_dir.mkdir(parents=True, exist_ok=True)
    metadata = file_metadata(file_path)
    text, result = extract_best_text(file_path, parser_name, start_page, max_pages)
    classification = classify_document(text, metadata)
    gate = quality_gate_status(result, min_quality, False)

    layout: dict[str, object] | None = None
    page_map: list[dict[str, object]] = []
    if file_path.suffix.lower() == ".pdf":
        try:
            layout = pdf_layout_json(file_path, start_page, max_pages)
            page_map = layout_to_page_map(layout)
        except Exception as exc:
            page_map = [{"error": repr(exc)}]

    if chunk_by == "page" and file_path.suffix.lower() == ".pdf" and result.parser in PAGE_CHUNK_PARSERS:
        page_extractor = get_pdf_page_extractor(result.parser)
        chunks = split_pages_into_chunks(page_extractor(file_path, start_page, max_pages) if page_extractor else [])
    else:
        chunks = split_text_into_chunks(text, chunk_size, overlap)

    source_md = out_dir / "source.md"
    source_txt = out_dir / "source.txt"
    source_md.write_text(text + "\n", encoding="utf-8")
    source_txt.write_text(text + "\n", encoding="utf-8")

    metadata_path = write_json_file(out_dir / "metadata.json", metadata)
    quality_path = write_json_file(out_dir / "quality_report.json", {"quality": asdict(result), "gate": gate})
    page_map_path = write_json_file(out_dir / "page_map.json", page_map)
    layout_path = write_json_file(out_dir / "layout.json", layout) if layout is not None else None
    chunk_paths = write_chunks(
        chunks,
        out_dir / "chunks",
        "all",
        {
            "version": __version__,
            "source_file": str(file_path),
            "parser": result.parser,
            "format": detect_format(file_path),
            "chunk_by": chunk_by,
        },
    )

    manifest = {
        "version": __version__,
        "type": "knowledge-pack",
        "source_file": str(file_path),
        "parser": result.parser,
        "classification": classification,
        "quality": asdict(result),
        "quality_gate": gate,
        "chunk_count": len(chunks),
        "outputs": {
            "source_md": str(source_md),
            "source_txt": str(source_txt),
            "metadata": str(metadata_path),
            "quality_report": str(quality_path),
            "page_map": str(page_map_path),
            "layout": str(layout_path) if layout_path else None,
            "chunks": [str(path) for path in chunk_paths],
        },
    }
    manifest_path = write_json_file(out_dir / "manifest.json", manifest)
    readme_lines = [
        "# Knowledge Pack",
        "",
        f"- Source: `{file_path}`",
        f"- Parser: `{result.parser}`",
        f"- Classification: `{classification['profile']}`",
        f"- Quality: `{result.quality_score}` (`{result.quality_label}`)",
        f"- Chunk count: `{len(chunks)}`",
        f"- Manifest: `{manifest_path}`",
    ]
    readme_path = out_dir / "README.md"
    readme_path.write_text("\n".join(readme_lines), encoding="utf-8")
    manifest["outputs"]["manifest"] = str(manifest_path)
    manifest["outputs"]["readme"] = str(readme_path)
    write_json_file(manifest_path, manifest)
    return manifest


# ---------------------------------------------------------------------------
# RAG/问答和文档差异辅助
# ---------------------------------------------------------------------------

SEARCH_STOPWORDS = {
    "the", "and", "or", "of", "to", "in", "for", "a", "an", "is", "are", "this", "that",
    "什么", "哪些", "如何", "是否", "可以", "这个", "那个", "里面", "文档", "文件",
    "的", "了", "和", "与", "或", "是", "在", "有", "我", "你", "他", "她", "它",
}


def tokenize_for_search(text: str) -> list[str]:
    """Tokenize mixed Chinese/English text for deterministic local retrieval."""
    tokens = [token.lower() for token in _RE_TOKEN.findall(text or "")]
    return [token for token in tokens if token and token not in SEARCH_STOPWORDS]


def chunk_snippet(text: str, query_tokens: list[str], max_chars: int = RAG_SNIPPET_CHARS) -> str:
    """Return a compact snippet around the first query-token hit."""
    normalized = normalize_text(text)
    if len(normalized) <= max_chars:
        return normalized
    lowered = normalized.lower()
    hit_positions = [lowered.find(token.lower()) for token in query_tokens if token and lowered.find(token.lower()) >= 0]
    center = min(hit_positions) if hit_positions else 0
    start = max(0, center - max_chars // 3)
    end = min(len(normalized), start + max_chars)
    start = max(0, end - max_chars)
    prefix = "..." if start > 0 else ""
    suffix = "..." if end < len(normalized) else ""
    return prefix + normalized[start:end] + suffix


def load_chunks_jsonl(path: Path) -> list[dict[str, object]]:
    chunks: list[dict[str, object]] = []
    for line in path.read_text(encoding="utf-8").splitlines():
        if not line.strip():
            continue
        payload = json.loads(line)
        if isinstance(payload, dict) and payload.get("text"):
            chunks.append(payload)
    return chunks


def load_chunks_json(path: Path) -> list[dict[str, object]]:
    payload = json.loads(path.read_text(encoding="utf-8"))
    if isinstance(payload, dict):
        metadata = payload.get("metadata") if isinstance(payload.get("metadata"), dict) else {}
        chunks = payload.get("chunks")
        if isinstance(chunks, list):
            return [{**metadata, **chunk} for chunk in chunks if isinstance(chunk, dict) and chunk.get("text")]
    if isinstance(payload, list):
        return [chunk for chunk in payload if isinstance(chunk, dict) and chunk.get("text")]
    return []


def load_chunks_from_manifest(manifest_path: Path) -> list[dict[str, object]]:
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    outputs = manifest.get("outputs") if isinstance(manifest, dict) else {}
    if not isinstance(outputs, dict):
        return []
    chunk_paths = outputs.get("chunks")
    if not isinstance(chunk_paths, list):
        return []
    chunks: list[dict[str, object]] = []
    for item in chunk_paths:
        path = Path(str(item))
        if not path.exists():
            continue
        if path.suffix.lower() == ".jsonl":
            chunks.extend(load_chunks_jsonl(path))
        elif path.suffix.lower() == ".json":
            chunks.extend(load_chunks_json(path))
    return chunks


def load_chunks_from_index(index_path: Path) -> list[dict[str, object]]:
    index = json.loads(index_path.read_text(encoding="utf-8"))
    manifests = index.get("manifests") if isinstance(index, dict) else []
    chunks: list[dict[str, object]] = []
    if isinstance(manifests, list):
        for item in manifests:
            manifest_path = Path(str(item))
            if manifest_path.exists():
                chunks.extend(load_chunks_from_manifest(manifest_path))
    return chunks


def collect_chunks_from_pack(path: Path) -> list[dict[str, object]]:
    """Load chunks from a knowledge pack, batch index, chunks file, or pack directory."""
    if path.is_file():
        lower_name = path.name.lower()
        if lower_name == "manifest.json":
            return load_chunks_from_manifest(path)
        if lower_name == "index.json":
            return load_chunks_from_index(path)
        if path.suffix.lower() == ".jsonl":
            return load_chunks_jsonl(path)
        if path.suffix.lower() == ".json":
            return load_chunks_json(path)
        return []

    candidates = [
        path / "chunks" / "chunks.jsonl",
        path / "chunks.jsonl",
        path / "manifest.json",
        path / "index.json",
    ]
    for candidate in candidates:
        if candidate.exists():
            loaded = collect_chunks_from_pack(candidate)
            if loaded:
                return loaded

    chunks: list[dict[str, object]] = []
    for candidate in sorted(path.rglob("chunks.jsonl")):
        chunks.extend(load_chunks_jsonl(candidate))
    return chunks


def collect_qa_chunks(
    source: Path,
    parser_name: str,
    start_page: int,
    max_pages: int | None,
    chunk_by: str,
    chunk_size: int,
    overlap: int,
) -> tuple[list[dict[str, object]], dict[str, object]]:
    """Collect searchable chunks from a knowledge pack or directly from a document."""
    if source.is_dir() or source.suffix.lower() in {".jsonl", ".json"}:
        chunks = collect_chunks_from_pack(source)
        if not chunks:
            raise RuntimeError(f"未在知识包或 chunks 文件中找到可检索文本：{source}")
        return chunks, {"source_type": "knowledge_pack_or_chunks", "source": str(source)}

    fmt = detect_format(source)
    if fmt is None:
        raise ValueError(f"不支持的输入：{source.suffix}")

    text, result = extract_best_text(source, parser_name, start_page, max_pages)
    if fmt == "pdf" and chunk_by == "page" and result.parser in PAGE_CHUNK_PARSERS:
        page_extractor = get_pdf_page_extractor(result.parser)
        chunks = split_pages_into_chunks(page_extractor(source, start_page, max_pages) if page_extractor else [])
    else:
        chunks = split_text_into_chunks(text, chunk_size, overlap)

    for chunk in chunks:
        chunk.setdefault("source_file", str(source))
        chunk.setdefault("parser", result.parser)
        chunk.setdefault("format", fmt)
        chunk.setdefault("chunk_by", chunk_by if fmt == "pdf" else "char")
    return chunks, {"source_type": "document", "source": str(source), "parser": result.parser, "quality": asdict(result)}


def rank_chunks(question: str, chunks: list[dict[str, object]], top_k: int) -> tuple[list[dict[str, object]], list[str]]:
    """Rank chunks with a transparent token-overlap score."""
    query_tokens = tokenize_for_search(question)
    query_set = set(query_tokens)
    scored: list[dict[str, object]] = []
    compact_question = "".join(query_tokens)
    for chunk in chunks:
        text = str(chunk.get("text") or "")
        if not text.strip():
            continue
        text_tokens = tokenize_for_search(text)
        if not text_tokens:
            continue
        token_counts: dict[str, int] = {}
        for token in text_tokens:
            token_counts[token] = token_counts.get(token, 0) + 1
        overlap = sum(min(3, token_counts.get(token, 0)) for token in query_set)
        density = overlap / max(1, len(set(text_tokens)) ** 0.5)
        phrase_bonus = 1.0 if compact_question and compact_question in "".join(text_tokens) else 0.0
        score = round(density + phrase_bonus, 4)
        if score <= 0:
            continue
        scored.append(
            {
                **chunk,
                "retrieval_score": score,
                "snippet": chunk_snippet(text, query_tokens),
            }
        )
    scored.sort(key=lambda item: (float(item.get("retrieval_score") or 0.0), int(item.get("chars") or 0)), reverse=True)
    return scored[:top_k], query_tokens


def extract_relevant_sentences(question: str, matches: list[dict[str, object]], max_sentences: int) -> list[dict[str, object]]:
    """Build an extractive answer from the best-scoring source sentences."""
    query_tokens = set(tokenize_for_search(question))
    candidates: list[dict[str, object]] = []
    for index, match in enumerate(matches, start=1):
        text = normalize_text(str(match.get("text") or match.get("snippet") or ""))
        sentences = [part.strip() for part in re.split(r"(?<=[。！？!?\.])\s+|\n+", text) if part.strip()]
        if not sentences and text:
            sentences = [text[:RAG_SNIPPET_CHARS]]
        for sentence in sentences:
            tokens = set(tokenize_for_search(sentence))
            score = len(query_tokens & tokens)
            if score > 0:
                candidates.append(
                    {
                        "citation": index,
                        "score": score,
                        "text": sentence[:RAG_SNIPPET_CHARS],
                        "source_file": match.get("source_file"),
                        "page": match.get("page"),
                        "chunk_id": match.get("chunk_id"),
                    }
                )
    candidates.sort(key=lambda item: (int(item["score"]), len(str(item["text"]))), reverse=True)
    return candidates[:max_sentences]


def source_citation(match: dict[str, object], index: int) -> dict[str, object]:
    return {
        "citation": index,
        "source_file": match.get("source_file"),
        "page": match.get("page"),
        "chunk_id": match.get("chunk_id"),
        "parser": match.get("parser"),
        "score": match.get("retrieval_score"),
        "snippet": match.get("snippet"),
    }


def build_qa_payload(
    source: Path,
    question: str,
    parser_name: str,
    start_page: int,
    max_pages: int | None,
    chunk_by: str,
    chunk_size: int,
    overlap: int,
    top_k: int,
    answer_sentences: int,
) -> dict[str, object]:
    chunks, source_info = collect_qa_chunks(source, parser_name, start_page, max_pages, chunk_by, chunk_size, overlap)
    matches, query_tokens = rank_chunks(question, chunks, top_k)
    answers = extract_relevant_sentences(question, matches, answer_sentences)
    citations = [source_citation(match, index) for index, match in enumerate(matches, start=1)]
    return {
        "version": __version__,
        "type": "qa",
        "method": "extractive_local_retrieval",
        "question": question,
        "query_tokens": query_tokens,
        "source": source_info,
        "chunk_count": len(chunks),
        "answer": {
            "status": "found" if answers else "no_direct_answer",
            "sentences": answers,
            "note": "This command retrieves and quotes relevant source text; it does not call an LLM.",
        },
        "citations": citations,
    }


def qa_to_markdown(payload: dict[str, object]) -> str:
    answer = payload.get("answer") if isinstance(payload.get("answer"), dict) else {}
    sentences = answer.get("sentences") if isinstance(answer, dict) else []
    citations = payload.get("citations") if isinstance(payload.get("citations"), list) else []
    lines = [
        "# QA Report",
        "",
        f"- Question: {payload.get('question')}",
        f"- Method: `{payload.get('method')}`",
        f"- Status: `{answer.get('status') if isinstance(answer, dict) else None}`",
        f"- Chunk count: `{payload.get('chunk_count')}`",
        "",
        "## Extractive Answer",
        "",
    ]
    if isinstance(sentences, list) and sentences:
        for item in sentences:
            if isinstance(item, dict):
                lines.append(f"- {item.get('text')} [{item.get('citation')}]")
    else:
        lines.append("No direct answer was found from the retrieved chunks.")
    lines.extend(["", "## Citations", "", "| Ref | Page | Chunk | Score | Source | Snippet |", "|---:|---:|---:|---:|---|---|"])
    for item in citations:
        if isinstance(item, dict):
            snippet = str(item.get("snippet") or "").replace("\n", " ")[:220]
            lines.append(
                f"| {item.get('citation')} | {item.get('page') or ''} | {item.get('chunk_id') or ''} | "
                f"{item.get('score') or ''} | `{item.get('source_file') or ''}` | {snippet} |"
            )
    return "\n".join(lines) + "\n"


def write_qa_output(payload: dict[str, object], out_path: Path, output_format: str) -> list[Path]:
    if output_format == "all":
        out_path.mkdir(parents=True, exist_ok=True)
        targets = {"json": out_path / "qa_report.json", "md": out_path / "qa_report.md"}
    else:
        out_path.parent.mkdir(parents=True, exist_ok=True)
        targets = {output_format: out_path}

    written: list[Path] = []
    for suffix, target in targets.items():
        if suffix == "json":
            write_json_file(target, payload)
        elif suffix == "md":
            target.write_text(qa_to_markdown(payload), encoding="utf-8")
        written.append(target)
    return written


def diff_line_stats(left_text: str, right_text: str) -> dict[str, object]:
    left_lines = left_text.splitlines()
    right_lines = right_text.splitlines()
    matcher = difflib.SequenceMatcher(None, left_lines, right_lines, autojunk=True)
    stats = {"equal": 0, "replace": 0, "delete": 0, "insert": 0}
    changed_blocks: list[dict[str, object]] = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        stats[tag] = stats.get(tag, 0) + max(i2 - i1, j2 - j1)
        if tag != "equal":
            changed_blocks.append(
                {
                    "type": tag,
                    "left_lines": [i1 + 1, i2],
                    "right_lines": [j1 + 1, j2],
                    "left_preview": "\n".join(left_lines[i1:min(i2, i1 + 3)])[:600],
                    "right_preview": "\n".join(right_lines[j1:min(j2, j1 + 3)])[:600],
                }
            )
    return {
        "left_line_count": len(left_lines),
        "right_line_count": len(right_lines),
        "stats": stats,
        "changed_block_count": len(changed_blocks),
        "changed_blocks": changed_blocks[:20],
    }


def limited_unified_diff(left_text: str, right_text: str, left_name: str, right_name: str, max_lines: int) -> list[str]:
    diff_iter = difflib.unified_diff(
        left_text.splitlines(),
        right_text.splitlines(),
        fromfile=left_name,
        tofile=right_name,
        lineterm="",
    )
    lines: list[str] = []
    for index, line in enumerate(diff_iter):
        if index >= max_lines:
            lines.append(f"... diff truncated after {max_lines} lines ...")
            break
        lines.append(line)
    return lines


def scalar_for_diff(value: object) -> str:
    if isinstance(value, float):
        return f"{value:.2f}"
    if value is None:
        return ""
    return str(value)


def diff_invoice_fields(left_fields: dict[str, object], right_fields: dict[str, object]) -> list[dict[str, object]]:
    keys = INVOICE_SUMMARY_COLUMNS[4:16]
    changes: list[dict[str, object]] = []
    for key in keys:
        left_value = scalar_for_diff(left_fields.get(key))
        right_value = scalar_for_diff(right_fields.get(key))
        if left_value != right_value:
            changes.append({"field": key, "left": left_value, "right": right_value})
    left_items = left_fields.get("items") if isinstance(left_fields.get("items"), list) else []
    right_items = right_fields.get("items") if isinstance(right_fields.get("items"), list) else []
    if len(left_items) != len(right_items):
        changes.append({"field": "items_count", "left": len(left_items), "right": len(right_items)})
    return changes


def build_diff_payload(
    left_path: Path,
    right_path: Path,
    parser_name: str,
    start_page: int,
    max_pages: int | None,
    profile: str,
    max_diff_lines: int,
) -> dict[str, object]:
    left_text, left_result = extract_best_text(left_path, parser_name, start_page, max_pages)
    right_text, right_result = extract_best_text(right_path, parser_name, start_page, max_pages)
    left_metadata = file_metadata(left_path)
    right_metadata = file_metadata(right_path)
    left_classification = classify_document(left_text, left_metadata)
    right_classification = classify_document(right_text, right_metadata)

    field_changes: list[dict[str, object]] = []
    should_extract_invoice = profile == "invoice" or (
        profile == "auto"
        and (left_classification.get("profile") == "invoice" or right_classification.get("profile") == "invoice")
    )
    if should_extract_invoice:
        left_fields = extract_invoice_fields_from_text(left_text, left_path, left_result.parser, left_result, strict=True)
        right_fields = extract_invoice_fields_from_text(right_text, right_path, right_result.parser, right_result, strict=True)
        field_changes = diff_invoice_fields(left_fields, right_fields)

    return {
        "version": __version__,
        "type": "diff-docs",
        "left": {"file": str(left_path), "parser": left_result.parser, "quality": asdict(left_result), "classification": left_classification},
        "right": {"file": str(right_path), "parser": right_result.parser, "quality": asdict(right_result), "classification": right_classification},
        "similarity": {
            "full_text_ratio": round(difflib.SequenceMatcher(None, left_text, right_text, autojunk=True).ratio(), 4),
            "first_50000_chars": similarity_score(left_text, right_text, 50000),
        },
        "line_diff": diff_line_stats(left_text, right_text),
        "unified_diff": limited_unified_diff(left_text, right_text, left_path.name, right_path.name, max_diff_lines),
        "field_changes": field_changes,
    }


def diff_to_markdown(payload: dict[str, object]) -> str:
    similarity = payload.get("similarity") if isinstance(payload.get("similarity"), dict) else {}
    line_diff = payload.get("line_diff") if isinstance(payload.get("line_diff"), dict) else {}
    stats = line_diff.get("stats") if isinstance(line_diff.get("stats"), dict) else {}
    field_changes = payload.get("field_changes") if isinstance(payload.get("field_changes"), list) else []
    unified = payload.get("unified_diff") if isinstance(payload.get("unified_diff"), list) else []
    lines = [
        "# Document Diff Report",
        "",
        f"- Left: `{payload.get('left', {}).get('file') if isinstance(payload.get('left'), dict) else ''}`",
        f"- Right: `{payload.get('right', {}).get('file') if isinstance(payload.get('right'), dict) else ''}`",
        f"- Full text similarity: `{similarity.get('full_text_ratio')}`",
        f"- First 50000 chars similarity: `{similarity.get('first_50000_chars')}`",
        f"- Changed blocks: `{line_diff.get('changed_block_count')}`",
        f"- Line stats: `{stats}`",
        "",
        "## Field Changes",
        "",
    ]
    if field_changes:
        lines.extend(["| Field | Left | Right |", "|---|---|---|"])
        for change in field_changes:
            if isinstance(change, dict):
                lines.append(f"| `{change.get('field')}` | {change.get('left') or ''} | {change.get('right') or ''} |")
    else:
        lines.append("No structured field changes detected.")
    lines.extend(["", "## Unified Diff", "", "```diff"])
    lines.extend(str(line) for line in unified)
    lines.extend(["```", ""])
    return "\n".join(lines)


def write_diff_output(payload: dict[str, object], out_path: Path, output_format: str) -> list[Path]:
    if output_format == "all":
        out_path.mkdir(parents=True, exist_ok=True)
        targets = {"json": out_path / "diff_report.json", "md": out_path / "diff_report.md"}
    else:
        out_path.parent.mkdir(parents=True, exist_ok=True)
        targets = {output_format: out_path}

    written: list[Path] = []
    for suffix, target in targets.items():
        if suffix == "json":
            write_json_file(target, payload)
        elif suffix == "md":
            target.write_text(diff_to_markdown(payload), encoding="utf-8")
        written.append(target)
    return written


# ---------------------------------------------------------------------------
# 报告生成
# ---------------------------------------------------------------------------

def make_report(
    file_path: Path,
    actual_path: Path,
    out_dir: Path,
    fmt: str,
    max_pages: int | None,
    start_page: int,
    results: list[ParseResult],
    normalized_texts: dict[str, str],
    similarity_chars: int,
    output_format: str,
    ocr_fallback: bool = False,
    min_quality: float = 0.0,
    fail_on_bad: bool = False,
) -> tuple[Path, Path]:
    """生成 JSON 和 Markdown 对比报告，返回两个报告文件路径。"""
    ok_results = [r for r in results if r.status == "ok"]

    pairwise: list[dict[str, object]] = []
    for i, left in enumerate(ok_results):
        for right in ok_results[i + 1:]:
            left_text = normalized_texts.get(left.parser, "")
            right_text = normalized_texts.get(right.parser, "")
            pairwise.append({
                "left": left.parser,
                "right": right.parser,
                "similarity_first_chars": similarity_chars,
                "similarity": similarity_score(left_text, right_text, similarity_chars),
            })

    report = {
        "version": __version__,
        "format": fmt,
        "source_file": str(file_path),
        "actual_file": str(actual_path),
        "start_page": start_page,
        "max_pages": max_pages,
        "output_format": output_format,
        "generated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "results": [asdict(r) for r in results],
        "pairwise_similarity": pairwise,
    }
    recommended = choose_recommended_result(results)
    diagnostics = build_diagnostics(results)
    ocr_notice = ocr_fallback_notice(ocr_fallback)
    if ocr_notice:
        diagnostics.append(ocr_notice)
    quality_gate = quality_gate_status(recommended, min_quality, fail_on_bad)
    report["recommendation"] = asdict(recommended) if recommended else None
    report["diagnostics"] = diagnostics
    report["quality_gate"] = quality_gate
    report["ocr_fallback_requested"] = ocr_fallback
    report["ocr_fallback_available"] = False
    json_path = out_dir / "compare_report.json"
    json_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    fmt_label = {"pdf": "PDF", "word": "Word", "ppt": "PPT", "excel": "Excel", "html": "HTML"}.get(fmt, fmt)
    md_lines = [
        f"# {fmt_label} 解析对比报告",
        "",
        f"- 版本：`{__version__}`",
        f"- 源文件：`{file_path}`",
        f"- 实际解析文件：`{actual_path}`",
        f"- 文件格式：`{fmt_label}`",
    ]
    if fmt == "pdf":
        md_lines.extend([
            f"- 起始页：`{start_page}`",
            f"- 页数限制：`{max_pages if max_pages is not None else '全量'}`",
        ])
    md_lines.extend([
        f"- 解析输出格式：`{output_format}`",
        f"- 输出目录：`{out_dir}`",
        "",
        "## 推荐",
        "",
    ])
    if recommended:
        md_lines.extend([
            f"- 推荐解析器：`{recommended.parser}`",
            f"- 质量评分：`{recommended.quality_score}` (`{recommended.quality_label}`)",
            f"- 非空白字符数：`{recommended.non_space_chars}`",
            f"- 耗时：`{recommended.seconds}s`",
            "",
        ])
    else:
        md_lines.extend(["- 未找到可推荐的成功解析结果。", ""])

    if quality_gate["enabled"]:
        md_lines.extend([
            "## 质量门禁",
            "",
            f"- 最低质量：`{quality_gate['min_quality']}`",
            f"- 失败标签拦截：`{quality_gate['fail_on_bad']}`",
            f"- 通过：`{quality_gate['passed']}`",
            f"- 原因：`{quality_gate['reason']}`",
            "",
        ])

    md_lines.extend(["## 质量诊断", ""])
    for item in diagnostics:
        md_lines.append(f"- {item}")

    md_lines.extend([
        "",
        "## 解析结果",
        "",
        "| 解析器 | 状态 | 质量 | 评分 | 耗时(s) | 字符数 | 非空白字符 | 行数 | 段落数 | 中文字符 | cid | 控制字符 | 输出 | 错误 |",
        "|---|---:|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---|---|",
    ])
    for r in results:
        output_name = Path(r.output_file).name if r.output_file else ""
        error = (r.error or "").replace("|", "\\|")
        md_lines.append(
            f"| {r.parser} | {r.status} | {r.quality_label} | {r.quality_score} | "
            f"{r.seconds} | {r.chars} | {r.non_space_chars} | {r.lines} | {r.paragraphs} | "
            f"{r.chinese_chars} | {r.cid_markers} | {r.control_chars} | {output_name} | {error} |"
        )

    md_lines.extend(["", "## 相似度", ""])
    if pairwise:
        md_lines.extend([
            f"> 仅比较每份结果的前 {similarity_chars} 个字符，适合快速判断抽取顺序和文本差异。",
            "",
            "| 左侧解析器 | 右侧解析器 | 相似度 |",
            "|---|---|---:|",
        ])
        for item in pairwise:
            md_lines.append(f"| {item['left']} | {item['right']} | {item['similarity']} |")
    else:
        md_lines.append("没有两个以上成功输出，无法计算相似度。")

    md_lines.extend(["", "## 预览", ""])
    for r in results:
        if r.status != "ok":
            continue
        md_lines.extend([
            f"### {r.parser}", "",
            "```text", r.preview, "```", "",
        ])

    md_path = out_dir / "compare_report.md"
    md_path.write_text("\n".join(md_lines), encoding="utf-8")

    return json_path, md_path


# ---------------------------------------------------------------------------
# 子命令：compare（多解析器对比）
# ---------------------------------------------------------------------------

def cmd_compare(args: argparse.Namespace) -> int:
    file_path = args.file.resolve()
    out_dir = args.out_dir.resolve() if args.out_dir else file_path.parent / f"{file_path.stem}_parse_output"
    max_pages = None if args.max_pages == 0 else args.max_pages
    start_page = args.start_page

    if not file_path.exists():
        print(f"错误：文件不存在：{file_path}", file=sys.stderr)
        return 1

    fmt = detect_format(file_path)
    if fmt is None:
        supported = ", ".join(sorted(EXTENSION_FORMAT_MAP.keys()))
        print(f"错误：不支持的文件格式：{file_path.suffix}。支持：{supported}", file=sys.stderr)
        return 1

    available_parsers = get_parsers_for_format(fmt)
    if args.parsers:
        parser_names = [resolve_parser_name(n.strip().lower()) for n in args.parsers.split(",") if n.strip()]
    else:
        parser_names = available_parsers

    valid_names = set(available_parsers)
    unknown = [n for n in parser_names if n not in valid_names]
    if unknown:
        print(f"警告：{fmt} 格式不支持以下解析器，将被跳过：{', '.join(unknown)}", file=sys.stderr)
        parser_names = [n for n in parser_names if n in valid_names]

    if not parser_names:
        print("错误：没有可用的解析器", file=sys.stderr)
        return 1

    fmt_label = {"pdf": "PDF", "word": "Word", "ppt": "PPT", "excel": "Excel", "html": "HTML"}.get(fmt, fmt)
    print(f"文件格式：{fmt_label}")
    print(f"解析器：{', '.join(parser_names)}")

    out_dir.mkdir(parents=True, exist_ok=True)

    actual_path = file_path
    temp_dir_obj: tempfile.TemporaryDirectory[str] | None = None
    markitdown_skip_reason: str | None = None
    if fmt == "pdf" and "markitdown" in parser_names:
        temp_dir_obj = tempfile.TemporaryDirectory(prefix="parse_compare_")
        actual_path, markitdown_skip_reason = prepare_markitdown_pdf_target(
            file_path, start_page, max_pages, Path(temp_dir_obj.name)
        )

    def run_parser(name: str) -> tuple[ParseResult, str]:
        if name == "markitdown" and markitdown_skip_reason:
            return ParseResult(parser=name, status="skipped", seconds=0, error=markitdown_skip_reason), ""

        extract_func = get_extractor(fmt, name)
        if extract_func is None:
            return ParseResult(parser=name, status="skipped", seconds=0, error=f"未注册的解析器：{name}"), ""

        target = actual_path if (name == "markitdown" and fmt == "pdf") else file_path

        start = time.perf_counter()
        try:
            raw_text = extract_func(target, start_page, max_pages)
            out_file, normalized = write_extracted_outputs(out_dir, name, raw_text, args.output_format)
            result = text_stats(name, "ok", time.perf_counter() - start, normalized, out_file)
            return result, normalized
        except Exception as exc:
            result = ParseResult(
                parser=name,
                status="failed",
                seconds=round(time.perf_counter() - start, 3),
                error=repr(exc),
            )
            return result, ""

    results: list[ParseResult] = []
    normalized_texts: dict[str, str] = {}
    try:
        if args.parallel:
            print("并行解析中...")
            with ThreadPoolExecutor(max_workers=min(len(parser_names), 4)) as pool:
                futures = {pool.submit(run_parser, name): name for name in parser_names}
                for future in as_completed(futures):
                    result, text = future.result()
                    results.append(result)
                    if text:
                        normalized_texts[result.parser] = text
                    icon = "OK" if result.status == "ok" else "FAIL" if result.status == "failed" else "SKIP"
                    print(f"  [{icon}] {result.parser}: {result.chars} chars, {result.seconds}s")
        else:
            for name in parser_names:
                print(f"解析中: {name}...", end=" ", flush=True)
                result, text = run_parser(name)
                results.append(result)
                if text:
                    normalized_texts[result.parser] = text
                if result.status == "ok":
                    print(f"OK {result.chars} chars, {result.seconds}s")
                elif result.status == "failed":
                    print(f"FAIL {result.error}")
                else:
                    print(f"SKIP {result.error}")

        json_path, md_path = make_report(
            file_path, actual_path, out_dir, fmt, max_pages, start_page,
            results, normalized_texts, args.similarity_chars, args.output_format, args.ocr_fallback,
            args.min_quality, args.fail_on_bad,
        )
        print(f"\n对比报告：{md_path}")
        print(f"机器可读报告：{json_path}")
    finally:
        if temp_dir_obj:
            temp_dir_obj.cleanup()

    recommended = choose_recommended_result(results)
    gate = quality_gate_status(recommended, args.min_quality, args.fail_on_bad)
    if gate["enabled"] and not gate["passed"]:
        print(f"质量门禁未通过：{gate['reason']}（best={gate['best_quality_score']} {gate['best_quality_label']}）", file=sys.stderr)
        return 2
    return 0


# ---------------------------------------------------------------------------
# 子命令：convert（文档转 md/txt/json）
# ---------------------------------------------------------------------------

def cmd_convert(args: argparse.Namespace) -> int:
    file_path = args.file.resolve()
    parser_name = resolve_parser_name(args.parser)

    if not file_path.exists():
        print(f"错误：文件不存在：{file_path}", file=sys.stderr)
        return 1

    fmt = detect_format(file_path)
    if fmt is None:
        supported = ", ".join(sorted(EXTENSION_FORMAT_MAP.keys()))
        print(f"错误：不支持的文件格式：{file_path.suffix}。支持：{supported}", file=sys.stderr)
        return 1

    available = get_parsers_for_format(fmt)
    if parser_name not in available:
        print(f"错误：{fmt} 格式不支持解析器 {parser_name}。可用：{', '.join(available)}", file=sys.stderr)
        return 1

    extract_func = get_extractor(fmt, parser_name)
    if extract_func is None:
        print(f"错误：未找到解析器 {parser_name}", file=sys.stderr)
        return 1

    output_format = args.format

    # 确定输出路径
    if args.output:
        out_path = Path(args.output).resolve()
    else:
        out_path = default_output_path(file_path, output_format)

    # PDF 特殊处理
    max_pages = None if args.max_pages == 0 else args.max_pages
    start_page = args.start_page

    print(f"解析器：{parser_name}")
    print(f"输入：{file_path}")
    print(f"输出格式：{output_format}")
    print(f"输出：{out_path}")
    if args.ocr_fallback:
        print("OCR fallback requested; use the `ocr` subcommand to run Tesseract OCR.", file=sys.stderr)

    start = time.perf_counter()
    temp_dir_obj: tempfile.TemporaryDirectory[str] | None = None
    try:
        target_path = file_path
        if fmt == "pdf" and parser_name == "markitdown":
            temp_dir_obj = tempfile.TemporaryDirectory(prefix="parse_convert_")
            target_path, skip_reason = prepare_markitdown_pdf_target(
                file_path, start_page, max_pages, Path(temp_dir_obj.name)
            )
            if skip_reason:
                print(f"错误：{skip_reason}", file=sys.stderr)
                return 1

        raw_text = extract_func(target_path, start_page, max_pages)
        written = write_single_conversion_output(out_path, parser_name, file_path, raw_text, output_format)
        elapsed = time.perf_counter() - start
        chars = len(normalize_text(raw_text))
        print(f"完成：{chars} 字符，{elapsed:.3f}s")
        for path in written:
            print(f"输出文件：{path}")
        return 0
    except ImportError as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 1
    except Exception as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 1
    finally:
        if temp_dir_obj:
            temp_dir_obj.cleanup()


# ---------------------------------------------------------------------------
# 子命令：batch（批量转换）
# ---------------------------------------------------------------------------

def cmd_batch(args: argparse.Namespace) -> int:
    input_dir = Path(args.dir).resolve()
    parser_name = resolve_parser_name(args.parser)
    out_dir = Path(args.output_dir).resolve() if args.output_dir else input_dir / "batch_output"
    output_format = args.format
    extensions = parse_extensions(args.ext)

    if not input_dir.exists() or not input_dir.is_dir():
        print(f"错误：目录不存在：{input_dir}", file=sys.stderr)
        return 1

    # 收集文件
    files: list[Path] = []
    for ext in sorted(extensions):
        files.extend(input_dir.glob(f"*{ext}"))
    files.sort()

    if not files:
        print(f"错误：目录下没有找到 {', '.join(extensions)} 文件", file=sys.stderr)
        return 1

    print(f"输入目录：{input_dir}")
    print(f"输出目录：{out_dir}")
    print(f"输出格式：{output_format}")
    print(f"解析器：{parser_name}")
    print(f"文件数：{len(files)}")

    out_dir.mkdir(parents=True, exist_ok=True)

    # 处理单个文件
    def process_one(file_path: Path) -> tuple[str, str, bool]:
        fmt = detect_format(file_path)
        if fmt is None:
            return file_path.name, f"不支持的格式: {file_path.suffix}", False

        available = get_parsers_for_format(fmt)
        if parser_name not in available:
            return file_path.name, f"解析器 {parser_name} 不支持 {fmt}", False

        extract_func = get_extractor(fmt, parser_name)
        if extract_func is None:
            return file_path.name, f"未找到解析器", False

        try:
            raw_text = extract_func(file_path, 1, None)
            normalized = normalize_text(raw_text)
            if output_format == "all":
                target = out_dir / file_path.stem
            else:
                target = out_dir / f"{file_path.stem}.{output_format}"
            written = write_single_conversion_output(target, parser_name, file_path, raw_text, output_format)
            return file_path.name, f"{len(normalized)} chars -> {', '.join(p.name for p in written)}", True
        except Exception as exc:
            return file_path.name, repr(exc), False

    # 执行
    success = 0
    fail = 0
    if args.parallel:
        with ThreadPoolExecutor(max_workers=min(len(files), 4)) as pool:
            futures = {pool.submit(process_one, f): f for f in files}
            for future in as_completed(futures):
                name, msg, ok = future.result()
                icon = "OK" if ok else "FAIL"
                print(f"  [{icon}] {name}: {msg}")
                if ok:
                    success += 1
                else:
                    fail += 1
    else:
        for f in files:
            name, msg, ok = process_one(f)
            icon = "OK" if ok else "FAIL"
            print(f"  [{icon}] {name}: {msg}")
            if ok:
                success += 1
            else:
                fail += 1

    print(f"\n完成：{success} 成功，{fail} 失败")
    print(f"输出目录：{out_dir}")
    return 0 if fail == 0 else 1


# ---------------------------------------------------------------------------
# 子命令：scan-dir（批量质量扫描）
# ---------------------------------------------------------------------------

def choose_scan_parser(fmt: str, requested_parser: str) -> str | None:
    """Choose a parser for quality scan; prefer fast native parsers for PDF."""
    available = get_parsers_for_format(fmt)
    if not available:
        return None
    requested = resolve_parser_name(requested_parser.strip().lower())
    if requested != "auto":
        return requested if requested in available else None

    preferences = {
        "pdf": ["pymupdf", "pdfplumber", "pypdf", "pdfminer", "markitdown", "liteparse"],
        "word": ["markitdown", "python-docx"],
        "ppt": ["markitdown", "python-pptx"],
        "excel": ["markitdown", "openpyxl"],
        "html": ["markitdown", "beautifulsoup4"],
    }
    for parser_name in preferences.get(fmt, available):
        if parser_name in available and parser_available(parser_name):
            return parser_name
    return available[0]


def scan_one_document(
    file_path: Path,
    requested_parser: str,
    start_page: int,
    max_pages: int | None,
    min_quality: float,
    fail_on_bad: bool,
) -> dict[str, object]:
    """Extract a small sample and return a quality row for scan-dir."""
    fmt = detect_format(file_path)
    if fmt is None:
        return {
            "source_file": str(file_path),
            "file_name": file_path.name,
            "format": None,
            "parser": None,
            "status": "skipped",
            "seconds": 0.0,
            "quality_score": 0.0,
            "quality_label": "unknown",
            "chars": 0,
            "non_space_chars": 0,
            "lines": 0,
            "chinese_chars": 0,
            "cid_markers": 0,
            "control_chars": 0,
            "needs_ocr": False,
            "gate_passed": False,
            "gate_reason": "unsupported_format",
            "error": f"不支持的格式：{file_path.suffix}",
        }

    parser_name = choose_scan_parser(fmt, requested_parser)
    if parser_name is None:
        return {
            "source_file": str(file_path),
            "file_name": file_path.name,
            "format": fmt,
            "parser": None,
            "status": "failed",
            "seconds": 0.0,
            "quality_score": 0.0,
            "quality_label": "unknown",
            "chars": 0,
            "non_space_chars": 0,
            "lines": 0,
            "chinese_chars": 0,
            "cid_markers": 0,
            "control_chars": 0,
            "needs_ocr": fmt == "pdf",
            "gate_passed": False,
            "gate_reason": "no_parser",
            "error": f"没有可用于 {fmt} 的解析器",
        }

    extract_func = get_extractor(fmt, parser_name)
    if extract_func is None:
        return {
            "source_file": str(file_path),
            "file_name": file_path.name,
            "format": fmt,
            "parser": parser_name,
            "status": "failed",
            "seconds": 0.0,
            "quality_score": 0.0,
            "quality_label": "unknown",
            "chars": 0,
            "non_space_chars": 0,
            "lines": 0,
            "chinese_chars": 0,
            "cid_markers": 0,
            "control_chars": 0,
            "needs_ocr": fmt == "pdf",
            "gate_passed": False,
            "gate_reason": "missing_extractor",
            "error": f"未找到解析器 {parser_name}",
        }

    temp_dir_obj: tempfile.TemporaryDirectory[str] | None = None
    start = time.perf_counter()
    try:
        target_path = file_path
        if fmt == "pdf" and parser_name == "markitdown" and max_pages is not None:
            temp_dir_obj = tempfile.TemporaryDirectory(prefix="parse_scan_")
            target_path, skip_reason = prepare_markitdown_pdf_target(
                file_path, start_page, max_pages, Path(temp_dir_obj.name)
            )
            if skip_reason:
                raise RuntimeError(skip_reason)
        raw_text = extract_func(target_path, start_page, max_pages)
        normalized = normalize_text(raw_text)
        result = text_stats(parser_name, "ok", time.perf_counter() - start, normalized, None)
        gate = quality_gate_status(result, min_quality, fail_on_bad)
        needs_ocr = (
            fmt == "pdf"
            and (
                result.quality_label in QUALITY_BAD_LABELS
                or (min_quality > 0 and result.quality_score < min_quality)
            )
        )
        return {
            "source_file": str(file_path),
            "file_name": file_path.name,
            "format": fmt,
            "parser": parser_name,
            "status": result.status,
            "seconds": result.seconds,
            "quality_score": result.quality_score,
            "quality_label": result.quality_label,
            "chars": result.chars,
            "non_space_chars": result.non_space_chars,
            "lines": result.lines,
            "chinese_chars": result.chinese_chars,
            "cid_markers": result.cid_markers,
            "control_chars": result.control_chars,
            "needs_ocr": needs_ocr,
            "gate_passed": gate["passed"],
            "gate_reason": gate["reason"],
            "error": None,
        }
    except Exception as exc:
        gate = quality_gate_status(None, min_quality, fail_on_bad)
        return {
            "source_file": str(file_path),
            "file_name": file_path.name,
            "format": fmt,
            "parser": parser_name,
            "status": "failed",
            "seconds": round(time.perf_counter() - start, 3),
            "quality_score": 0.0,
            "quality_label": "failed",
            "chars": 0,
            "non_space_chars": 0,
            "lines": 0,
            "chinese_chars": 0,
            "cid_markers": 0,
            "control_chars": 0,
            "needs_ocr": fmt == "pdf",
            "gate_passed": gate["passed"],
            "gate_reason": gate["reason"],
            "error": repr(exc),
        }
    finally:
        if temp_dir_obj:
            temp_dir_obj.cleanup()


def write_scan_reports(
    rows: list[dict[str, object]],
    out_dir: Path,
    input_dir: Path,
    requested_parser: str,
    extensions: set[str],
    recursive: bool,
    min_quality: float,
    fail_on_bad: bool,
) -> tuple[Path, Path, Path]:
    """Write scan-dir JSON, Markdown, and CSV reports."""
    out_dir.mkdir(parents=True, exist_ok=True)
    summary = {
        "version": __version__,
        "input_dir": str(input_dir),
        "extensions": sorted(extensions),
        "recursive": recursive,
        "requested_parser": requested_parser,
        "min_quality": min_quality,
        "fail_on_bad": fail_on_bad,
        "file_count": len(rows),
        "ok_count": sum(1 for row in rows if row["status"] == "ok"),
        "failed_count": sum(1 for row in rows if row["status"] == "failed"),
        "needs_ocr_count": sum(1 for row in rows if row["needs_ocr"]),
        "gate_failed_count": sum(1 for row in rows if not row["gate_passed"]),
        "generated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
    }

    json_path = out_dir / "scan_report.json"
    json_path.write_text(
        json.dumps({"summary": summary, "files": rows}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    csv_path = out_dir / "scan_report.csv"
    fieldnames = [
        "source_file", "file_name", "format", "parser", "status", "quality_label",
        "quality_score", "chars", "non_space_chars", "lines", "chinese_chars",
        "cid_markers", "control_chars", "needs_ocr", "gate_passed", "gate_reason",
        "seconds", "error",
    ]
    with csv_path.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow({key: row.get(key) for key in fieldnames})

    md_path = out_dir / "scan_report.md"
    md_lines = [
        "# Directory Quality Scan",
        "",
        f"- Version: `{__version__}`",
        f"- Input directory: `{input_dir}`",
        f"- Extensions: `{', '.join(sorted(extensions))}`",
        f"- Recursive: `{recursive}`",
        f"- Requested parser: `{requested_parser}`",
        f"- Minimum quality: `{min_quality}`",
        f"- Fail on bad: `{fail_on_bad}`",
        f"- Files: `{summary['file_count']}`",
        f"- OK: `{summary['ok_count']}`",
        f"- Failed: `{summary['failed_count']}`",
        f"- Needs OCR: `{summary['needs_ocr_count']}`",
        f"- Gate failed: `{summary['gate_failed_count']}`",
        "",
        "| File | Format | Parser | Status | Quality | Score | Chars | Needs OCR | Gate | Error |",
        "|---|---|---|---|---|---:|---:|---|---|---|",
    ]
    for row in rows:
        error = str(row.get("error") or "").replace("|", "\\|")
        md_lines.append(
            f"| {row['file_name']} | {row['format']} | {row['parser']} | {row['status']} | "
            f"{row['quality_label']} | {row['quality_score']} | {row['chars']} | "
            f"{row['needs_ocr']} | {row['gate_passed']} | {error} |"
        )
    md_path.write_text("\n".join(md_lines), encoding="utf-8")
    return json_path, md_path, csv_path


def cmd_scan_dir(args: argparse.Namespace) -> int:
    input_dir = Path(args.dir).resolve()
    out_dir = Path(args.out_dir).resolve() if args.out_dir else input_dir / "scan_output"
    extensions = parse_extensions(args.ext)
    max_pages = None if args.max_pages == 0 else args.max_pages

    if not input_dir.exists() or not input_dir.is_dir():
        print(f"错误：目录不存在：{input_dir}", file=sys.stderr)
        return 1

    files: list[Path] = []
    for ext in sorted(extensions):
        iterator = input_dir.rglob(f"*{ext}") if args.recursive else input_dir.glob(f"*{ext}")
        files.extend(path for path in iterator if path.is_file())
    files = sorted(set(files))

    if not files:
        print(f"错误：目录下没有找到 {', '.join(sorted(extensions))} 文件", file=sys.stderr)
        return 1

    print(f"扫描目录：{input_dir}")
    print(f"文件数：{len(files)}")
    print(f"解析器：{args.parser}")
    print(f"输出目录：{out_dir}")

    rows: list[dict[str, object]] = []
    for file_path in files:
        print(f"扫描: {file_path.name}...", end=" ", flush=True)
        row = scan_one_document(
            file_path,
            args.parser,
            args.start_page,
            max_pages,
            args.min_quality,
            args.fail_on_bad,
        )
        rows.append(row)
        print(f"{row['status']} {row['quality_label']} {row['quality_score']}")

    json_path, md_path, csv_path = write_scan_reports(
        rows,
        out_dir,
        input_dir,
        args.parser,
        extensions,
        args.recursive,
        args.min_quality,
        args.fail_on_bad,
    )
    print(f"扫描报告：{md_path}")
    print(f"机器可读报告：{json_path}")
    print(f"CSV：{csv_path}")

    gate_failed = [row for row in rows if not row["gate_passed"]]
    if args.fail_on_bad and gate_failed:
        print(f"质量门禁未通过：{len(gate_failed)} 个文件", file=sys.stderr)
        return 2
    return 0


# ---------------------------------------------------------------------------
# 子命令：tables（表格提取）
# ---------------------------------------------------------------------------

def cmd_tables(args: argparse.Namespace) -> int:
    file_path = args.file.resolve()

    if not file_path.exists():
        print(f"错误：文件不存在：{file_path}", file=sys.stderr)
        return 1

    if file_path.suffix.lower() != ".pdf":
        print("错误：表格提取仅支持 PDF 文件", file=sys.stderr)
        return 1

    # 获取页数
    page_count = get_pdf_page_count(file_path)
    print(f"PDF 总页数：{page_count}")

    # 解析页码
    if args.pages:
        page_indices = parse_page_spec(args.pages, page_count)
    else:
        page_indices = list(range(page_count))

    print(f"提取范围：第 {page_indices[0] + 1} 页 ~ 第 {page_indices[-1] + 1} 页（共 {len(page_indices)} 页）")

    # 提取
    start = time.perf_counter()
    try:
        tables = extract_tables_from_pdf(file_path, page_indices)
    except ImportError as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 1
    except Exception as exc:
        print(f"错误：提取失败：{exc}", file=sys.stderr)
        return 1

    elapsed = time.perf_counter() - start
    print(f"提取到 {len(tables)} 个表格，耗时 {elapsed:.3f}s")

    if not tables:
        print("未找到任何表格")
        return 0

    # 输出
    fmt = args.format or "md"
    if fmt == "all":
        out_dir = Path(args.output).resolve() if args.output else file_path.parent / f"{file_path.stem}_tables"
        out_dir.mkdir(parents=True, exist_ok=True)
        outputs = {
            "md": out_dir / f"{file_path.stem}_tables.md",
            "csv": out_dir / f"{file_path.stem}_tables.csv",
            "json": out_dir / f"{file_path.stem}_tables.json",
        }
    else:
        suffix = fmt
        if args.output:
            outputs = {suffix: Path(args.output).resolve()}
        else:
            outputs = {suffix: file_path.parent / f"{file_path.stem}_tables.{suffix}"}

    for suffix, out_path in outputs.items():
        if suffix == "csv":
            content = tables_to_csv(tables)
        elif suffix == "json":
            content = tables_to_json(tables)
        else:
            content = tables_to_markdown(tables)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(content, encoding="utf-8")
        print(f"输出：{out_path}")
    return 0


# ---------------------------------------------------------------------------
# 子命令：doctor（依赖自检）
# ---------------------------------------------------------------------------

def cmd_doctor(args: argparse.Namespace) -> int:
    rows = parser_dependency_rows(args.format)
    missing_packages = sorted({str(row["package"]) for row in rows if not row["available"]})
    ocr_rows = OCR_DEPENDENCIES if getattr(args, "ocr", False) else []
    missing_ocr_python = [
        str(row["package"])
        for row in ocr_rows
        if not row["available"] and row["kind"] == "python"
    ]
    missing_ocr_system = [
        str(row["package"])
        for row in ocr_rows
        if not row["available"] and row["kind"] == "system"
    ]
    opendataloader_rows = OPENLOADER_DEPENDENCIES if getattr(args, "opendataloader", False) else []
    missing_opendataloader_python = [
        str(row["package"])
        for row in opendataloader_rows
        if not row["available"] and row["kind"] == "python"
    ]
    missing_opendataloader_system = [
        str(row["package"])
        for row in opendataloader_rows
        if not row["available"] and row["kind"] == "system"
    ]

    if args.json:
        payload = {
            "version": __version__,
            "python": sys.version.split()[0],
            "format": args.format,
            "rows": rows,
            "missing_packages": missing_packages,
            "ocr_rows": ocr_rows,
            "missing_ocr_python_packages": missing_ocr_python,
            "missing_ocr_system_packages": missing_ocr_system,
            "opendataloader_rows": opendataloader_rows,
            "missing_opendataloader_python_packages": missing_opendataloader_python,
            "missing_opendataloader_system_packages": missing_opendataloader_system,
            "install_command": f"pip install {' '.join(missing_packages)}" if missing_packages else None,
            "ocr_install_hint": ocr_install_hint(missing_ocr_python, missing_ocr_system),
            "opendataloader_install_hint": opendataloader_install_hint(missing_opendataloader_python, missing_opendataloader_system),
        }
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        strict_missing = missing_packages or missing_ocr_python or missing_ocr_system or missing_opendataloader_python or missing_opendataloader_system
        return 1 if strict_missing and args.strict else 0

    print(f"parse_pdf_compare {__version__}")
    print(f"Python: {sys.version.split()[0]}")
    print("")
    print("| 格式 | 解析器 | 状态 | 模块 | 安装包 |")
    print("|---|---|---|---|---|")
    for row in rows:
        modules = ", ".join(str(module) for module in row["modules"])
        print(f"| {row['format']} | {row['parser']} | {row['status']} | {modules} | {row['package']} |")

    if missing_packages:
        print("")
        print("缺失依赖安装建议：")
        print(f"pip install {' '.join(missing_packages)}")

    if ocr_rows:
        print("")
        print("| OCR 依赖 | 类型 | 状态 | 安装包 |")
        print("|---|---|---|---|")
        for row in ocr_rows:
            status = "ok" if row["available"] else "missing"
            print(f"| {row['name']} | {row['kind']} | {status} | {row['package']} |")
        if missing_ocr_python or missing_ocr_system:
            print("")
            print("OCR 依赖安装建议：")
            if missing_ocr_python:
                print(f"pip install {' '.join(missing_ocr_python)}")
            if missing_ocr_system:
                print("安装系统 Tesseract OCR，并确保 tesseract 在 PATH 中。")

    if opendataloader_rows:
        print("")
        print("| opendataloader-pdf 依赖 | 类型 | 状态 | 安装包 |")
        print("|---|---|---|---|")
        for row in opendataloader_rows:
            status = "ok" if row["available"] else "missing"
            print(f"| {row['name']} | {row['kind']} | {status} | {row['package']} |")
        if missing_opendataloader_python or missing_opendataloader_system:
            print("")
            print("opendataloader-pdf 依赖安装建议：")
            if missing_opendataloader_python:
                print(f"pip install {' '.join(missing_opendataloader_python)}")
            if missing_opendataloader_system:
                print("需要系统安装 Java (JDK/JRE)，并确保 java 命令在 PATH 中。")

    if missing_packages or missing_ocr_python or missing_ocr_system or missing_opendataloader_python or missing_opendataloader_system:
        return 1 if args.strict else 0

    print("")
    print("所有检查范围内的解析器依赖均可导入。")
    return 0


# ---------------------------------------------------------------------------
# 子命令：metadata（元数据提取）
# ---------------------------------------------------------------------------

def cmd_metadata(args: argparse.Namespace) -> int:
    file_path = args.file.resolve()
    if not file_path.exists():
        print(f"错误：文件不存在：{file_path}", file=sys.stderr)
        return 1

    try:
        metadata = file_metadata(file_path)
    except Exception as exc:
        print(f"错误：元数据读取失败：{exc}", file=sys.stderr)
        return 1

    output_format = args.format
    if args.output:
        out_path = Path(args.output).resolve()
    else:
        out_path = file_path.parent / f"{file_path.stem}_metadata.{output_format}"

    out_path.parent.mkdir(parents=True, exist_ok=True)
    if output_format == "json":
        out_path.write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")
    else:
        out_path.write_text(metadata_to_markdown(metadata), encoding="utf-8")
    print(f"输出：{out_path}")
    return 0


# ---------------------------------------------------------------------------
# 子命令：chunk（分块输出）
# ---------------------------------------------------------------------------

def cmd_chunk(args: argparse.Namespace) -> int:
    file_path = args.file.resolve()
    parser_name = resolve_parser_name(args.parser)

    if not file_path.exists():
        print(f"错误：文件不存在：{file_path}", file=sys.stderr)
        return 1

    fmt = detect_format(file_path)
    if fmt is None:
        supported = ", ".join(sorted(EXTENSION_FORMAT_MAP.keys()))
        print(f"错误：不支持的文件格式：{file_path.suffix}。支持：{supported}", file=sys.stderr)
        return 1

    if parser_name not in get_parsers_for_format(fmt):
        print(f"错误：{fmt} 格式不支持解析器 {parser_name}", file=sys.stderr)
        return 1

    extract_func = get_extractor(fmt, parser_name)
    if extract_func is None:
        print(f"错误：未找到解析器 {parser_name}", file=sys.stderr)
        return 1

    max_pages = None if args.max_pages == 0 else args.max_pages
    start_page = args.start_page
    temp_dir_obj: tempfile.TemporaryDirectory[str] | None = None

    if args.ocr_fallback:
        print("OCR fallback requested; use the `ocr` subcommand to run Tesseract OCR.", file=sys.stderr)

    try:
        target_path = file_path
        if fmt == "pdf" and parser_name == "markitdown":
            temp_dir_obj = tempfile.TemporaryDirectory(prefix="parse_chunk_")
            target_path, skip_reason = prepare_markitdown_pdf_target(
                file_path, start_page, max_pages, Path(temp_dir_obj.name)
            )
            if skip_reason:
                print(f"错误：{skip_reason}", file=sys.stderr)
                return 1

        if args.chunk_by == "page":
            if fmt != "pdf":
                print("错误：--chunk-by page 仅支持 PDF", file=sys.stderr)
                return 1
            page_extractor = get_pdf_page_extractor(parser_name)
            if page_extractor is None:
                supported = ", ".join(sorted(PAGE_CHUNK_PARSERS))
                print(f"错误：解析器 {parser_name} 不支持分页分块。可用：{supported}", file=sys.stderr)
                return 1
            pages = page_extractor(file_path, start_page, max_pages)
            chunks = split_pages_into_chunks(pages)
        else:
            raw_text = extract_func(target_path, start_page, max_pages)
            chunks = split_text_into_chunks(raw_text, args.chunk_size, args.overlap)
    except Exception as exc:
        print(f"错误：分块失败：{exc}", file=sys.stderr)
        return 1
    finally:
        if temp_dir_obj:
            temp_dir_obj.cleanup()

    if args.output:
        out_path = Path(args.output).resolve()
    else:
        suffix = "jsonl" if args.format == "jsonl" else args.format
        out_path = file_path.parent / f"{file_path.stem}_chunks.{suffix}"
        if args.format == "all":
            out_path = file_path.parent / f"{file_path.stem}_chunks"

    meta = {
        "version": __version__,
        "source_file": str(file_path),
        "parser": parser_name,
        "format": fmt,
        "chunk_by": args.chunk_by,
        "chunk_size": args.chunk_size,
        "overlap": args.overlap,
        "ocr_fallback_requested": args.ocr_fallback,
        "ocr_fallback_available": False,
    }
    written = write_chunks(chunks, out_path, args.format, meta)
    print(f"分块数：{len(chunks)}")
    for path in written:
        print(f"输出：{path}")
    return 0


# ---------------------------------------------------------------------------
# 子命令：render-pages（PDF 页面截图）
# ---------------------------------------------------------------------------

def cmd_render_pages(args: argparse.Namespace) -> int:
    file_path = args.file.resolve()
    if not file_path.exists():
        print(f"错误：文件不存在：{file_path}", file=sys.stderr)
        return 1
    if file_path.suffix.lower() != ".pdf":
        print("错误：render-pages 仅支持 PDF 文件", file=sys.stderr)
        return 1

    try:
        page_count = get_pdf_page_count(file_path)
        page_indices = resolve_page_indices(args.pages, page_count, default_all=False)
        out_dir = args.out_dir.resolve() if args.out_dir else file_path.parent / f"{file_path.stem}_pages"
        rendered = render_pdf_pages(file_path, page_indices, out_dir, args.dpi)
    except Exception as exc:
        print(f"错误：页面截图失败：{exc}", file=sys.stderr)
        return 1

    manifest = {
        "version": __version__,
        "source_file": str(file_path),
        "page_count": page_count,
        "dpi": args.dpi,
        "pages": rendered,
    }
    manifest_path = out_dir / "render_pages.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"输出目录：{out_dir}")
    print(f"页面数：{len(rendered)}")
    for item in rendered:
        print(f"输出：{item['file']}")
    print(f"清单：{manifest_path}")
    return 0


# ---------------------------------------------------------------------------
# 子命令：ocr（真实 OCR，可选依赖）
# ---------------------------------------------------------------------------

def cmd_ocr(args: argparse.Namespace) -> int:
    file_path = args.file.resolve()
    if not file_path.exists():
        print(f"错误：文件不存在：{file_path}", file=sys.stderr)
        return 1
    if file_path.suffix.lower() != ".pdf":
        print("错误：ocr 当前仅支持 PDF 文件", file=sys.stderr)
        return 1

    try:
        page_count = get_pdf_page_count(file_path)
        page_indices = resolve_page_indices(args.pages, page_count, default_all=False)
        pages = ocr_pdf_pages(file_path, page_indices, args.lang, args.dpi)
    except Exception as exc:
        print(f"错误：OCR 失败：{exc}", file=sys.stderr)
        print("依赖提示：pip install pymupdf pytesseract pillow；并安装系统 Tesseract OCR。", file=sys.stderr)
        return 1

    if args.output:
        out_path = Path(args.output).resolve()
    else:
        out_path = file_path.parent / f"{file_path.stem}_ocr.{args.format}"
    written = write_ocr_output(out_path, file_path, pages, args.format, args.lang, args.dpi)
    print(f"OCR 页数：{len(pages)}")
    print(f"输出：{written}")
    return 0


# ---------------------------------------------------------------------------
# 子命令：auto（智能流水线）
# ---------------------------------------------------------------------------

def cmd_auto(args: argparse.Namespace) -> int:
    file_path = args.file.resolve()
    if not file_path.exists():
        print(f"错误：文件不存在：{file_path}", file=sys.stderr)
        return 1

    fmt = detect_format(file_path)
    if fmt is None:
        supported = ", ".join(sorted(EXTENSION_FORMAT_MAP.keys()))
        print(f"错误：不支持的文件格式：{file_path.suffix}。支持：{supported}", file=sys.stderr)
        return 1

    out_dir = args.out_dir.resolve() if args.out_dir else file_path.parent / f"{file_path.stem}_auto_output"
    out_dir.mkdir(parents=True, exist_ok=True)
    max_pages = None if args.max_pages == 0 else args.max_pages

    try:
        metadata = file_metadata(file_path)
        text, result = extract_best_text(file_path, args.parser, args.start_page, max_pages)
        text, result, ocr_fallback = maybe_ocr_fallback(
            file_path,
            text,
            result,
            args.auto_ocr,
            args.min_quality,
            args.ocr_pages,
            args.ocr_lang,
            args.ocr_dpi,
        )
        classification = classify_document(text, metadata)
        gate = quality_gate_status(result, args.min_quality, args.fail_on_bad)
    except Exception as exc:
        print(f"错误：auto 流水线失败：{exc}", file=sys.stderr)
        return 1

    best_md = out_dir / "best.md"
    best_txt = out_dir / "best.txt"
    best_json = out_dir / "best.json"
    best_md.write_text(text + "\n", encoding="utf-8")
    best_txt.write_text(text + "\n", encoding="utf-8")
    best_json.write_text(
        json.dumps(
            {
                "version": __version__,
                "source_file": str(file_path),
                "parser": result.parser,
                "quality": asdict(result),
                "text": text,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )

    metadata_path = out_dir / "metadata.json"
    metadata_path.write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")

    layout_path: Path | None = None
    page_map_path: Path | None = None
    if fmt == "pdf" and args.layout:
        try:
            layout = pdf_layout_json(file_path, args.start_page, max_pages)
            layout_path = write_json_file(out_dir / "layout.json", layout)
            page_map_path = write_json_file(out_dir / "page_map.json", layout_to_page_map(layout))
        except Exception as exc:
            write_json_file(out_dir / "layout_error.json", {"error": repr(exc)})

    chunks: list[dict[str, object]]
    if fmt == "pdf" and args.chunk_by == "page" and result.parser in PAGE_CHUNK_PARSERS:
        page_extractor = get_pdf_page_extractor(result.parser)
        chunks = split_pages_into_chunks(page_extractor(file_path, args.start_page, max_pages) if page_extractor else [])
    else:
        chunks = split_text_into_chunks(text, args.chunk_size, args.overlap)
    chunk_meta = {
        "version": __version__,
        "source_file": str(file_path),
        "parser": result.parser,
        "format": fmt,
        "chunk_by": args.chunk_by if fmt == "pdf" else "char",
    }
    chunk_paths = write_chunks(chunks, out_dir / "chunks", "all", chunk_meta)

    fields_path: str | None = None
    fields_payload: dict[str, object] | None = None
    if args.profile == "invoice" or (args.profile == "auto" and classification["profile"] == "invoice"):
        fields_payload = extract_invoice_fields_from_text(text, file_path, result.parser, result, strict=True)
        fields_path_obj = write_invoice_fields(fields_payload, out_dir / "fields.json", "json")
        write_invoice_fields(fields_payload, out_dir / "fields.md", "md")
        fields_path = str(fields_path_obj)

    decision = {
        "version": __version__,
        "source_file": str(file_path),
        "format": fmt,
        "parser": result.parser,
        "quality_score": result.quality_score,
        "quality_label": result.quality_label,
        "classification": classification,
        "quality_gate": gate,
        "needs_ocr": result.quality_label in QUALITY_BAD_LABELS or result.quality_score < args.min_quality,
        "ocr_fallback": ocr_fallback,
        "outputs": {
            "best_md": str(best_md),
            "best_txt": str(best_txt),
            "best_json": str(best_json),
            "metadata": str(metadata_path),
            "layout": str(layout_path) if layout_path else None,
            "page_map": str(page_map_path) if page_map_path else None,
            "chunks": [str(path) for path in chunk_paths],
            "fields": fields_path,
        },
    }
    decision_path = out_dir / "auto_report.json"
    decision_path.write_text(json.dumps(decision, ensure_ascii=False, indent=2), encoding="utf-8")

    report_lines = [
        "# Auto Parse Report",
        "",
        f"- Source: `{file_path}`",
        f"- Format: `{fmt}`",
        f"- Parser: `{result.parser}`",
        f"- Quality: `{result.quality_score}` (`{result.quality_label}`)",
        f"- Classification: `{classification['profile']}` (`{classification['confidence']}`)",
        f"- Needs OCR: `{decision['needs_ocr']}`",
        f"- OCR fallback: `{ocr_fallback}`",
        f"- Output directory: `{out_dir}`",
        "",
        "## Outputs",
        "",
        f"- Best Markdown: `{best_md}`",
        f"- Best TXT: `{best_txt}`",
        f"- Best JSON: `{best_json}`",
        f"- Metadata: `{metadata_path}`",
        f"- Auto JSON: `{decision_path}`",
    ]
    if fields_path:
        report_lines.append(f"- Fields: `{fields_path}`")
    if layout_path:
        report_lines.append(f"- Layout: `{layout_path}`")
    report_path = out_dir / "auto_report.md"
    report_path.write_text("\n".join(report_lines), encoding="utf-8")

    print(f"自动解析完成：{out_dir}")
    print(f"推荐解析器：{result.parser}，质量：{result.quality_score} {result.quality_label}")
    print(f"报告：{report_path}")
    if gate["enabled"] and not gate["passed"]:
        print(f"质量门禁未通过：{gate['reason']}", file=sys.stderr)
        return 2
    return 0


# ---------------------------------------------------------------------------
# 子命令：extract-fields（结构化字段抽取）
# ---------------------------------------------------------------------------

def cmd_extract_fields(args: argparse.Namespace) -> int:
    file_path = args.file.resolve()
    if not file_path.exists():
        print(f"错误：文件不存在：{file_path}", file=sys.stderr)
        return 1
    if args.profile != "invoice":
        print(f"错误：当前仅支持 profile=invoice，收到：{args.profile}", file=sys.stderr)
        return 1

    max_pages = None if args.max_pages == 0 else args.max_pages
    try:
        fields = extract_invoice_fields_from_file(file_path, args.parser, args.start_page, max_pages)
    except Exception as exc:
        print(f"错误：字段抽取失败：{exc}", file=sys.stderr)
        return 1

    if args.output:
        out_path = Path(args.output).resolve()
    else:
        out_path = file_path.parent / f"{file_path.stem}_fields.{args.format}"
    written = write_invoice_fields(fields, out_path, args.format)
    print(f"字段输出：{written}")
    validation = fields.get("validation")
    if isinstance(validation, dict):
        print(f"校验状态：{validation.get('status')}")
    return 0


# ---------------------------------------------------------------------------
# 子命令：export-xlsx（结构化批量导出）
# ---------------------------------------------------------------------------

def cmd_export_xlsx(args: argparse.Namespace) -> int:
    input_path = Path(args.path).resolve()
    if args.profile != "invoice":
        print(f"错误：当前仅支持 profile=invoice，收到：{args.profile}", file=sys.stderr)
        return 1

    extensions = parse_extensions(args.ext)
    max_pages = None if args.max_pages == 0 else args.max_pages
    try:
        files = collect_input_files(input_path, extensions, args.recursive)
    except Exception as exc:
        print(f"错误：收集文件失败：{exc}", file=sys.stderr)
        return 1

    if not files:
        print(f"错误：未找到输入文件，扩展名：{', '.join(sorted(extensions))}", file=sys.stderr)
        return 1

    rows: list[dict[str, object]] = []
    failures: list[dict[str, object]] = []
    for file_path in files:
        print(f"抽取: {file_path.name}...", end=" ", flush=True)
        try:
            fields = extract_invoice_fields_from_file(file_path, args.parser, args.start_page, max_pages)
            rows.append(fields)
            validation = fields.get("validation") if isinstance(fields.get("validation"), dict) else {}
            print(validation.get("status", "ok"))
        except Exception as exc:
            failures.append({"source_file": str(file_path), "error": repr(exc)})
            print(f"FAIL {exc}")

    if args.output:
        out_path = Path(args.output).resolve()
    else:
        if input_path.is_dir():
            out_path = input_path / "invoice_summary.xlsx"
        else:
            out_path = input_path.parent / f"{input_path.stem}_summary.xlsx"

    try:
        written = write_invoice_xlsx(rows, out_path)
    except Exception as exc:
        print(f"错误：XLSX 导出失败：{exc}", file=sys.stderr)
        return 1

    manifest_path = written.with_suffix(".json")
    manifest_path.write_text(
        json.dumps(
            {
                "version": __version__,
                "profile": args.profile,
                "input": str(input_path),
                "file_count": len(files),
                "success_count": len(rows),
                "failure_count": len(failures),
                "failures": failures,
                "output": str(written),
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    print(f"导出：{written}")
    print(f"清单：{manifest_path}")
    return 0 if not failures else 1


# ---------------------------------------------------------------------------
# 子命令：layout-json（版面结构输出）
# ---------------------------------------------------------------------------

def cmd_layout_json(args: argparse.Namespace) -> int:
    file_path = args.file.resolve()
    if not file_path.exists():
        print(f"错误：文件不存在：{file_path}", file=sys.stderr)
        return 1
    if file_path.suffix.lower() != ".pdf":
        print("错误：layout-json 当前仅支持 PDF", file=sys.stderr)
        return 1
    max_pages = None if args.max_pages == 0 else args.max_pages
    try:
        layout = pdf_layout_json(file_path, args.start_page, max_pages)
    except Exception as exc:
        print(f"错误：layout-json 失败：{exc}", file=sys.stderr)
        return 1
    out_path = Path(args.output).resolve() if args.output else file_path.parent / f"{file_path.stem}_layout.json"
    write_json_file(out_path, layout)
    print(f"输出：{out_path}")
    return 0


# ---------------------------------------------------------------------------
# 子命令：verify-fields（字段校验）
# ---------------------------------------------------------------------------

def cmd_verify_fields(args: argparse.Namespace) -> int:
    path = Path(args.path).resolve()
    if not path.exists():
        print(f"错误：文件不存在：{path}", file=sys.stderr)
        return 1
    if args.profile != "invoice":
        print(f"错误：当前仅支持 profile=invoice，收到：{args.profile}", file=sys.stderr)
        return 1

    try:
        if path.suffix.lower() == ".json":
            fields = json.loads(path.read_text(encoding="utf-8"))
            fields["validation"] = validate_invoice_fields(fields, strict=args.strict)
        else:
            max_pages = None if args.max_pages == 0 else args.max_pages
            fields = extract_invoice_fields_from_file(path, args.parser, args.start_page, max_pages, strict=args.strict)
    except Exception as exc:
        print(f"错误：字段校验失败：{exc}", file=sys.stderr)
        return 1

    validation = fields.get("validation") if isinstance(fields.get("validation"), dict) else {}
    out_path = Path(args.output).resolve() if args.output else path.parent / f"{path.stem}_verify.json"
    write_json_file(out_path, fields)
    print(f"校验状态：{validation.get('status')}")
    print(f"输出：{out_path}")
    return 0 if validation.get("status") == "ok" else 2


# ---------------------------------------------------------------------------
# 子命令：classify（自动分类和策略）
# ---------------------------------------------------------------------------

def cmd_classify(args: argparse.Namespace) -> int:
    file_path = args.file.resolve()
    if not file_path.exists():
        print(f"错误：文件不存在：{file_path}", file=sys.stderr)
        return 1
    max_pages = None if args.max_pages == 0 else args.max_pages
    try:
        metadata = file_metadata(file_path)
        text, result = extract_best_text(file_path, args.parser, args.start_page, max_pages)
        classification = classify_document(text, metadata)
        payload = {
            "version": __version__,
            "source_file": str(file_path),
            "parser": result.parser,
            "quality": asdict(result),
            "classification": classification,
        }
    except Exception as exc:
        print(f"错误：分类失败：{exc}", file=sys.stderr)
        return 1

    if args.output:
        out_path = Path(args.output).resolve()
        write_json_file(out_path, payload)
        print(f"输出：{out_path}")
    else:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0


# ---------------------------------------------------------------------------
# 子命令：knowledge-pack（可追溯 RAG 包）
# ---------------------------------------------------------------------------

def cmd_knowledge_pack(args: argparse.Namespace) -> int:
    file_path = args.file.resolve()
    if not file_path.exists():
        print(f"错误：文件不存在：{file_path}", file=sys.stderr)
        return 1
    max_pages = None if args.max_pages == 0 else args.max_pages
    out_dir = args.out_dir.resolve() if args.out_dir else file_path.parent / f"{file_path.stem}_knowledge_pack"
    try:
        manifest = write_knowledge_pack(
            file_path,
            out_dir,
            args.parser,
            args.start_page,
            max_pages,
            args.chunk_by,
            args.chunk_size,
            args.overlap,
            args.min_quality,
        )
    except Exception as exc:
        print(f"错误：knowledge-pack 失败：{exc}", file=sys.stderr)
        return 1
    print(f"知识包：{out_dir}")
    print(f"manifest：{manifest['outputs']['manifest']}")
    return 0


# ---------------------------------------------------------------------------
# 子命令：batch-knowledge（批量知识库生产）
# ---------------------------------------------------------------------------

def cmd_batch_knowledge(args: argparse.Namespace) -> int:
    input_dir = Path(args.dir).resolve()
    extensions = parse_extensions(args.ext)
    max_pages = None if args.max_pages == 0 else args.max_pages
    try:
        files = collect_input_files(input_dir, extensions, args.recursive)
    except Exception as exc:
        print(f"错误：收集文件失败：{exc}", file=sys.stderr)
        return 1
    if not files:
        print(f"错误：没有找到 {', '.join(sorted(extensions))} 文件", file=sys.stderr)
        return 1

    out_dir = args.out_dir.resolve() if args.out_dir else input_dir / "knowledge_packs"
    out_dir.mkdir(parents=True, exist_ok=True)
    manifests: list[dict[str, object]] = []
    failures: list[dict[str, object]] = []
    for file_path in files:
        print(f"知识包: {file_path.name}...", end=" ", flush=True)
        try:
            pack_dir = out_dir / file_path.stem
            manifest = write_knowledge_pack(
                file_path,
                pack_dir,
                args.parser,
                args.start_page,
                max_pages,
                args.chunk_by,
                args.chunk_size,
                args.overlap,
                args.min_quality,
            )
            manifests.append(manifest)
            print("ok")
        except Exception as exc:
            failures.append({"source_file": str(file_path), "error": repr(exc)})
            print(f"FAIL {exc}")

    index = {
        "version": __version__,
        "type": "batch-knowledge",
        "input_dir": str(input_dir),
        "file_count": len(files),
        "success_count": len(manifests),
        "failure_count": len(failures),
        "manifests": [manifest.get("outputs", {}).get("manifest") for manifest in manifests],
        "failures": failures,
    }
    index_path = write_json_file(out_dir / "index.json", index)
    print(f"索引：{index_path}")
    return 0 if not failures else 1


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    all_parsers = list_all_parsers()

    parser = argparse.ArgumentParser(
        prog="parse_pdf_compare",
        description="多格式文档解析工具：对比、转换、批量处理、表格提取。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")

    subparsers = parser.add_subparsers(dest="command", help="子命令")

    # --- compare ---
    p_compare = subparsers.add_parser("compare", help="多解析器对比（默认子命令）",
                                      formatter_class=argparse.RawDescriptionHelpFormatter)
    p_compare.add_argument("file", type=Path, help="文件路径")
    p_compare.add_argument("--out-dir", type=Path, default=None, help="输出目录")
    p_compare.add_argument("--max-pages", type=int, default=30, help="（仅 PDF）前 N 页，0=全量。默认 30")
    p_compare.add_argument("--start-page", type=int, default=1, help="（仅 PDF）起始页。默认 1")
    p_compare.add_argument("--parsers", default=None,
                           help=f"逗号分隔解析器。默认全部：{','.join(all_parsers)}")
    p_compare.add_argument("--output-format", choices=sorted(TEXT_OUTPUT_FORMATS), default="md",
                           help="解析器输出格式：md/txt/json/all。默认 md")
    p_compare.add_argument("--similarity-chars", type=int, default=50000, help="相似度比较字符数")
    p_compare.add_argument("--min-quality", type=float, default=0.0,
                           help="最低质量评分门槛，0-1。低于该值返回退出码 2")
    p_compare.add_argument("--fail-on-bad", action="store_true",
                           help="当最佳结果标签为 empty/bad 时返回退出码 2")
    p_compare.add_argument("--ocr-fallback", action="store_true",
                           help="标记需要 OCR 回退；真实 OCR 请使用 ocr 子命令")
    p_compare.add_argument("--parallel", action="store_true", help="并行执行")

    # --- convert ---
    p_convert = subparsers.add_parser("convert", help="文档转指定格式")
    p_convert.add_argument("file", type=Path, help="文件路径")
    p_convert.add_argument("--parser", default="markitdown", help="解析器（默认 markitdown）")
    p_convert.add_argument("--format", choices=sorted(TEXT_OUTPUT_FORMATS), default="md",
                           help="输出格式：md/txt/json/all。默认 md")
    p_convert.add_argument("-o", "--output", default=None,
                           help="输出文件路径；--format all 时为输出目录。默认按格式生成")
    p_convert.add_argument("--max-pages", type=int, default=30, help="（仅 PDF）前 N 页，0=全量")
    p_convert.add_argument("--start-page", type=int, default=1, help="（仅 PDF）起始页")
    p_convert.add_argument("--ocr-fallback", action="store_true",
                           help="标记需要 OCR 回退；真实 OCR 请使用 ocr 子命令")

    # --- batch ---
    p_batch = subparsers.add_parser("batch", help="批量转换目录下所有文档")
    p_batch.add_argument("dir", type=str, help="输入目录")
    p_batch.add_argument("--parser", default="markitdown", help="解析器（默认 markitdown）")
    p_batch.add_argument("--output-dir", default=None, help="输出目录（默认 <输入目录>/batch_output）")
    p_batch.add_argument("--ext", default=".pdf", help="文件扩展名，逗号分隔（默认 .pdf）")
    p_batch.add_argument("--format", choices=sorted(TEXT_OUTPUT_FORMATS), default="md",
                         help="输出格式：md/txt/json/all。默认 md")
    p_batch.add_argument("--parallel", action="store_true", help="并行处理")

    # --- scan-dir ---
    p_scan = subparsers.add_parser("scan-dir", help="批量扫描目录内文档解析质量")
    p_scan.add_argument("dir", type=str, help="输入目录")
    p_scan.add_argument("--parser", default="auto", help="解析器，默认 auto")
    p_scan.add_argument("--out-dir", type=Path, default=None, help="输出目录（默认 <输入目录>/scan_output）")
    p_scan.add_argument("--ext", default=".pdf", help="文件扩展名，逗号分隔（默认 .pdf）")
    p_scan.add_argument("--recursive", action="store_true", help="递归扫描子目录")
    p_scan.add_argument("--max-pages", type=int, default=3, help="（仅 PDF）抽样前 N 页，0=全量。默认 3")
    p_scan.add_argument("--start-page", type=int, default=1, help="（仅 PDF）起始页。默认 1")
    p_scan.add_argument("--min-quality", type=float, default=0.5, help="最低质量评分门槛，默认 0.5")
    p_scan.add_argument("--fail-on-bad", action="store_true", help="存在门禁失败文件时返回退出码 2")

    # --- tables ---
    p_tables = subparsers.add_parser("tables", help="从 PDF 提取表格")
    p_tables.add_argument("file", type=Path, help="PDF 文件路径")
    p_tables.add_argument("--pages", default=None, help="页码范围，如 1-5 或 1,3,5（默认全部）")
    p_tables.add_argument("--format", choices=["md", "csv", "json", "all"], default="md",
                          help="输出格式：md/csv/json/all。默认 md")
    p_tables.add_argument("-o", "--output", default=None,
                          help="输出文件路径；--format all 时为输出目录")

    # --- doctor ---
    p_doctor = subparsers.add_parser("doctor", help="检查解析器依赖是否可用")
    p_doctor.add_argument("--format", choices=sorted(FORMAT_PARSERS), default=None,
                          help="只检查某一类格式")
    p_doctor.add_argument("--json", action="store_true", help="输出 JSON")
    p_doctor.add_argument("--ocr", action="store_true", help="同时检查 OCR 可选依赖")
    p_doctor.add_argument("--opendataloader", action="store_true", help="同时检查 opendataloader-pdf 依赖（需要 Java）")
    p_doctor.add_argument("--strict", action="store_true",
                          help="存在缺失依赖时返回非零退出码")

    # --- metadata ---
    p_metadata = subparsers.add_parser("metadata", help="提取文件元数据")
    p_metadata.add_argument("file", type=Path, help="文件路径")
    p_metadata.add_argument("--format", choices=["md", "json"], default="json",
                            help="输出格式：md/json。默认 json")
    p_metadata.add_argument("-o", "--output", default=None, help="输出文件路径")

    # --- chunk ---
    p_chunk = subparsers.add_parser("chunk", help="分块输出文档文本")
    p_chunk.add_argument("file", type=Path, help="文件路径")
    p_chunk.add_argument("--parser", default="markitdown", help="解析器（默认 markitdown）")
    p_chunk.add_argument("--chunk-by", choices=["char", "page"], default="char",
                         help="分块方式：char/page。默认 char")
    p_chunk.add_argument("--format", choices=["jsonl", "json", "md", "txt", "all"], default="jsonl",
                         help="输出格式：jsonl/json/md/txt/all。默认 jsonl")
    p_chunk.add_argument("--chunk-size", type=int, default=2000, help="每块字符数。默认 2000")
    p_chunk.add_argument("--overlap", type=int, default=200, help="块间重叠字符数。默认 200")
    p_chunk.add_argument("--max-pages", type=int, default=30, help="（仅 PDF）前 N 页，0=全量")
    p_chunk.add_argument("--start-page", type=int, default=1, help="（仅 PDF）起始页")
    p_chunk.add_argument("--ocr-fallback", action="store_true",
                         help="标记需要 OCR 回退；真实 OCR 请使用 ocr 子命令")
    p_chunk.add_argument("-o", "--output", default=None,
                         help="输出文件路径；--format all 时为输出目录")

    # --- render-pages ---
    p_render = subparsers.add_parser("render-pages", help="把 PDF 页面渲染为 PNG 图片")
    p_render.add_argument("file", type=Path, help="PDF 文件路径")
    p_render.add_argument("--pages", default="1", help="页码范围，如 1-5、1,3,5、all。默认 1")
    p_render.add_argument("--dpi", type=int, default=200, help="渲染 DPI。默认 200")
    p_render.add_argument("--out-dir", type=Path, default=None, help="输出目录（默认 <文件名>_pages）")

    # --- ocr ---
    p_ocr = subparsers.add_parser("ocr", help="对 PDF 页面执行真实 OCR（Tesseract，可选依赖）")
    p_ocr.add_argument("file", type=Path, help="PDF 文件路径")
    p_ocr.add_argument("--pages", default="1", help="页码范围，如 1-5、1,3,5、all。默认 1")
    p_ocr.add_argument("--lang", default="chi_sim+eng", help="Tesseract 语言。默认 chi_sim+eng")
    p_ocr.add_argument("--dpi", type=int, default=200, help="渲染 DPI。默认 200")
    p_ocr.add_argument("--format", choices=["txt", "md", "json"], default="txt", help="输出格式。默认 txt")
    p_ocr.add_argument("-o", "--output", default=None, help="输出文件路径")

    # --- auto ---
    p_auto = subparsers.add_parser("auto", help="智能解析流水线：元数据、最佳文本、分块、字段和报告")
    p_auto.add_argument("file", type=Path, help="文件路径")
    p_auto.add_argument("--parser", default="auto", help="解析器，默认 auto")
    p_auto.add_argument("--profile", choices=["auto", "invoice"], default="auto", help="结构化 profile，默认 auto")
    p_auto.add_argument("--out-dir", type=Path, default=None, help="输出目录（默认 <文件名>_auto_output）")
    p_auto.add_argument("--max-pages", type=int, default=30, help="（仅 PDF）前 N 页，0=全量。默认 30")
    p_auto.add_argument("--start-page", type=int, default=1, help="（仅 PDF）起始页。默认 1")
    p_auto.add_argument("--min-quality", type=float, default=0.5, help="最低质量评分门槛，默认 0.5")
    p_auto.add_argument("--fail-on-bad", action="store_true", help="质量不达标时返回退出码 2")
    p_auto.add_argument("--chunk-by", choices=["char", "page"], default="char", help="分块方式，默认 char")
    p_auto.add_argument("--chunk-size", type=int, default=2000, help="字符分块大小。默认 2000")
    p_auto.add_argument("--overlap", type=int, default=200, help="字符分块重叠。默认 200")
    p_auto.add_argument("--auto-ocr", action="store_true", help="质量低于阈值时尝试 Tesseract OCR 回退")
    p_auto.add_argument("--ocr-pages", default="1", help="OCR 回退页码范围，如 1-3、all。默认 1")
    p_auto.add_argument("--ocr-lang", default="chi_sim+eng", help="OCR 语言。默认 chi_sim+eng")
    p_auto.add_argument("--ocr-dpi", type=int, default=200, help="OCR 渲染 DPI。默认 200")
    p_auto.add_argument("--layout", action="store_true", help="同时输出 layout.json 和 page_map.json")

    # --- extract-fields ---
    p_fields = subparsers.add_parser("extract-fields", help="从文档抽取结构化字段")
    p_fields.add_argument("file", type=Path, help="文件路径")
    p_fields.add_argument("--profile", choices=["invoice"], default="invoice", help="字段 profile。默认 invoice")
    p_fields.add_argument("--parser", default="auto", help="解析器，默认 auto")
    p_fields.add_argument("--format", choices=["json", "md"], default="json", help="输出格式。默认 json")
    p_fields.add_argument("--max-pages", type=int, default=3, help="（仅 PDF）前 N 页，0=全量。默认 3")
    p_fields.add_argument("--start-page", type=int, default=1, help="（仅 PDF）起始页。默认 1")
    p_fields.add_argument("-o", "--output", default=None, help="输出文件路径")

    # --- export-xlsx ---
    p_xlsx = subparsers.add_parser("export-xlsx", help="批量抽取结构化字段并导出 XLSX")
    p_xlsx.add_argument("path", type=str, help="文件或目录路径")
    p_xlsx.add_argument("--profile", choices=["invoice"], default="invoice", help="字段 profile。默认 invoice")
    p_xlsx.add_argument("--parser", default="auto", help="解析器，默认 auto")
    p_xlsx.add_argument("--ext", default=".pdf", help="目录输入时的扩展名，逗号分隔。默认 .pdf")
    p_xlsx.add_argument("--recursive", action="store_true", help="递归扫描目录")
    p_xlsx.add_argument("--max-pages", type=int, default=3, help="（仅 PDF）前 N 页，0=全量。默认 3")
    p_xlsx.add_argument("--start-page", type=int, default=1, help="（仅 PDF）起始页。默认 1")
    p_xlsx.add_argument("-o", "--output", default=None, help="输出 XLSX 路径")

    # --- layout-json ---
    p_layout = subparsers.add_parser("layout-json", help="输出 PDF 页面、块、行、span 和坐标 JSON")
    p_layout.add_argument("file", type=Path, help="PDF 文件路径")
    p_layout.add_argument("--max-pages", type=int, default=30, help="前 N 页，0=全量。默认 30")
    p_layout.add_argument("--start-page", type=int, default=1, help="起始页。默认 1")
    p_layout.add_argument("-o", "--output", default=None, help="输出 JSON 路径")

    # --- verify-fields ---
    p_verify = subparsers.add_parser("verify-fields", help="校验结构化字段")
    p_verify.add_argument("path", type=str, help="字段 JSON 或原始文档路径")
    p_verify.add_argument("--profile", choices=["invoice"], default="invoice", help="字段 profile。默认 invoice")
    p_verify.add_argument("--parser", default="auto", help="原始文档解析器，默认 auto")
    p_verify.add_argument("--strict", action="store_true", help="启用严格校验")
    p_verify.add_argument("--max-pages", type=int, default=3, help="原始 PDF 前 N 页，0=全量。默认 3")
    p_verify.add_argument("--start-page", type=int, default=1, help="原始 PDF 起始页。默认 1")
    p_verify.add_argument("-o", "--output", default=None, help="输出 JSON 路径")

    # --- classify ---
    p_classify = subparsers.add_parser("classify", help="自动分类文档并推荐处理策略")
    p_classify.add_argument("file", type=Path, help="文件路径")
    p_classify.add_argument("--parser", default="auto", help="解析器，默认 auto")
    p_classify.add_argument("--max-pages", type=int, default=3, help="前 N 页，0=全量。默认 3")
    p_classify.add_argument("--start-page", type=int, default=1, help="起始页。默认 1")
    p_classify.add_argument("-o", "--output", default=None, help="输出 JSON 路径；默认打印到 stdout")

    # --- knowledge-pack ---
    p_pack = subparsers.add_parser("knowledge-pack", help="生成可追溯 RAG/知识包")
    p_pack.add_argument("file", type=Path, help="文件路径")
    p_pack.add_argument("--parser", default="auto", help="解析器，默认 auto")
    p_pack.add_argument("--out-dir", type=Path, default=None, help="输出目录（默认 <文件名>_knowledge_pack）")
    p_pack.add_argument("--max-pages", type=int, default=30, help="前 N 页，0=全量。默认 30")
    p_pack.add_argument("--start-page", type=int, default=1, help="起始页。默认 1")
    p_pack.add_argument("--chunk-by", choices=["char", "page"], default="page", help="分块方式，默认 page")
    p_pack.add_argument("--chunk-size", type=int, default=2000, help="字符分块大小。默认 2000")
    p_pack.add_argument("--overlap", type=int, default=200, help="字符分块重叠。默认 200")
    p_pack.add_argument("--min-quality", type=float, default=0.5, help="最低质量评分门槛，默认 0.5")

    # --- batch-knowledge ---
    p_batch_pack = subparsers.add_parser("batch-knowledge", help="批量生成知识包")
    p_batch_pack.add_argument("dir", type=str, help="输入目录")
    p_batch_pack.add_argument("--parser", default="auto", help="解析器，默认 auto")
    p_batch_pack.add_argument("--out-dir", type=Path, default=None, help="输出目录（默认 <目录>/knowledge_packs）")
    p_batch_pack.add_argument("--ext", default=".pdf,.docx,.pptx,.xlsx,.html", help="扩展名，逗号分隔")
    p_batch_pack.add_argument("--recursive", action="store_true", help="递归扫描目录")
    p_batch_pack.add_argument("--max-pages", type=int, default=30, help="前 N 页，0=全量。默认 30")
    p_batch_pack.add_argument("--start-page", type=int, default=1, help="起始页。默认 1")
    p_batch_pack.add_argument("--chunk-by", choices=["char", "page"], default="page", help="分块方式，默认 page")
    p_batch_pack.add_argument("--chunk-size", type=int, default=2000, help="字符分块大小。默认 2000")
    p_batch_pack.add_argument("--overlap", type=int, default=200, help="字符分块重叠。默认 200")
    p_batch_pack.add_argument("--min-quality", type=float, default=0.5, help="最低质量评分门槛，默认 0.5")

    return parser


def main() -> int:
    parser = build_parser()

    # 向后兼容：无子命令时等同于 compare
    if len(sys.argv) > 1 and sys.argv[1] not in (
        "compare", "convert", "batch", "scan-dir", "tables", "doctor", "metadata",
        "chunk", "render-pages", "ocr", "auto", "extract-fields", "export-xlsx",
        "layout-json", "verify-fields", "classify", "knowledge-pack", "batch-knowledge",
        "--version", "-h", "--help",
    ):
        sys.argv.insert(1, "compare")

    args = parser.parse_args()

    if args.command is None:
        parser.print_help()
        return 0

    dispatch = {
        "compare": cmd_compare,
        "convert": cmd_convert,
        "batch": cmd_batch,
        "scan-dir": cmd_scan_dir,
        "tables": cmd_tables,
        "doctor": cmd_doctor,
        "metadata": cmd_metadata,
        "chunk": cmd_chunk,
        "render-pages": cmd_render_pages,
        "ocr": cmd_ocr,
        "auto": cmd_auto,
        "extract-fields": cmd_extract_fields,
        "export-xlsx": cmd_export_xlsx,
        "layout-json": cmd_layout_json,
        "verify-fields": cmd_verify_fields,
        "classify": cmd_classify,
        "knowledge-pack": cmd_knowledge_pack,
        "batch-knowledge": cmd_batch_knowledge,
    }

    handler = dispatch.get(args.command)
    if handler is None:
        parser.print_help()
        return 1

    return handler(args)


if __name__ == "__main__":
    raise SystemExit(main())

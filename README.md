# Document Skill - 多格式文档智能解析工具

这是一个可发布的 Codex skill，也可以作为独立命令行工具使用。它用于对同一个文档运行多个解析器，比较文本抽取质量，并帮助判断是否需要 OCR 或更干净的源文件。

支持格式：PDF、Word、PPT、Excel、HTML、TXT、CSV、JSON、XML、图片、音频、EPub、ZIP。

## 能力

| 格式 | 扩展名 | 解析器 |
|---|---|---|
| PDF | `.pdf` | `markitdown`, `pymupdf`, `pypdf`, `pdfplumber`, `pdfminer`, `liteparse`, `opendataloader` (需 Java) |
| Word | `.docx`, `.doc` | `markitdown`, `python-docx` |
| PPT | `.pptx`, `.ppt` | `markitdown`, `python-pptx` |
| Excel | `.xlsx`, `.xls` | `markitdown`, `openpyxl` |
| HTML | `.html`, `.htm` | `markitdown`, `beautifulsoup4` |
| 文本 | `.txt`, `.csv`, `.json`, `.xml`, `.md` | `markitdown` |
| 图片 | `.jpg`, `.png`, `.gif`, `.bmp`, `.tiff`, `.webp` | `markitdown` (EXIF + OCR，需 Tesseract) |
| 音频 | `.mp3`, `.wav` | `markitdown` (EXIF + 语音转录，需 whisper) |
| 电子书 | `.epub` | `markitdown` |
| 压缩包 | `.zip` | `markitdown` (遍历内部文件) |

主要命令：

- `auto`：智能流水线，一次生成 metadata、最佳文本、分块、分类、字段和报告。
- `auto --auto-ocr`：质量低于阈值时自动尝试 OCR 回退。
- `compare`：多解析器对比，生成 Markdown/JSON 报告，并可输出 `md/txt/json/all` 格式的解析结果。
- `convert`：用指定解析器把单个文档转成 `md/txt/json/all`。
- `batch`：批量转换目录里的文档。
- `scan-dir`：批量扫描目录里的文档解析质量，输出 Markdown/JSON/CSV 报告。
- `tables`：从 PDF 提取表格，输出 `md/csv/json/all`。
- `doctor`：检查解析器依赖是否可导入，并给出安装建议。
- `metadata`：提取文件元数据；PDF 会额外输出页数、内嵌 metadata、目录预览、抽样文本层判断。
- `chunk`：把抽取文本切成 `jsonl/json/md/txt/all` 分块，支持按字符或按 PDF 页分块。
- `render-pages`：把 PDF 指定页面渲染成 PNG 截图。
- `ocr`：通过 Tesseract 对 PDF 页面执行真正 OCR，输出 `txt/md/json`。
- `extract-fields`：结构化字段抽取；v3.6 支持 `--profile invoice`。
- `verify-fields`：结构化字段校验；发票支持金额、税额、明细合计、必填字段和严格格式校验。
- `export-xlsx`：批量抽取发票字段并导出 XLSX。
- `layout-json`：输出 PDF 页面、块、行、span、字体和坐标。
- `classify`：自动识别文档类型并给出处理策略。
- `knowledge-pack`：生成可追溯 RAG/知识包。
- `batch-knowledge`：批量生产知识包。
- `--min-quality / --fail-on-bad`：质量门禁，适合批处理或自动化流程。
- `--ocr-fallback`：标记需要 OCR 回退；真正 OCR 使用 `ocr` 子命令。

版本路线见 [docs/ROADMAP.md](docs/ROADMAP.md)。

## Codex Skill 结构

```text
pdf-skill/
├── SKILL.md
├── agents/
│   └── openai.yaml
├── scripts/
│   └── parse_document_compare.py
├── references/
├── parse_pdf_compare.py
├── requirements.txt
└── README.md
```

`parse_pdf_compare.py` 是兼容入口，核心实现位于 `scripts/parse_document_compare.py`。

## 触发方式

Codex 显式触发：

```text
$document-skill
```

对话示例：

```text
使用 $document-skill 对比这个 PDF 的解析质量，并推荐最佳解析器。
```

```text
使用 $document-skill，把 D:\documents\report.pdf 用 pymupdf 解析成 json。
```

```text
使用 $document-skill，提取这个 PDF 的 metadata，并判断是否有文本层。
```

```text
使用 $document-skill，把这个 PDF 按 2000 字分块输出 jsonl。
```

```text
使用 $document-skill，把这个 PDF 按页分块输出 jsonl，保留页码。
```

```text
使用 $document-skill，扫描 D:\documents 目录里的 PDF，质量低于 0.6 就失败。
```

```text
使用 $document-skill，把这个 PDF 的第 1-3 页渲染成图片。
```

```text
使用 $document-skill，对这个 PDF 第 1 页做 OCR，输出 Markdown。
```

```text
使用 $document-skill，自动解析这个 PDF，判断类型并输出最佳文本和字段。
```

```text
使用 $document-skill，抽取这张发票的结构化字段，输出 JSON。
```

```text
使用 $document-skill，把这个目录里的发票汇总导出 Excel。
```

```text
使用 $document-skill，自动解析这个 PDF，质量差就 OCR，并输出 layout。
```

```text
使用 $document-skill，校验这个发票字段 JSON 是否金额一致。
```

```text
使用 $document-skill，识别这个文档类型并推荐处理策略。
```

```text
使用 $document-skill，为这个 PDF 生成可追溯知识包，后面做 RAG。
```

```text
使用 $document-skill，检查 PDF 解析器依赖是否都安装。
```

自然语言也可以触发，例如：

- “帮我对比这个 PDF 用哪种解析器效果最好”
- “这个 PDF 抽出来都是 `(cid:...)`，看看要不要 OCR”
- “把这个 DOCX 转成 JSON”
- “提取这个 PDF 里的表格，输出 CSV 和 JSON”
- “把这个目录里的 PDF 批量转成 TXT”
- “提取 PDF 元数据和页数”
- “把文档切成 JSONL 分块，后面我要做检索”
- “按页分块这个 PDF，输出 JSONL”
- “批量扫描这个目录，看看哪些 PDF 需要 OCR”
- “解析质量低于 0.6 就返回失败”
- “把 PDF 第 1 页截图出来看看”
- “对这个扫描版 PDF 做真正 OCR”
- “自动解析这个文档，能抽字段就抽字段”
- “提取这张发票的发票号码、金额、税额”
- “把这些发票批量导出到 Excel”
- “解析不好就自动 OCR”
- “输出这个 PDF 的 layout 坐标”
- “校验发票字段是否一致”
- “判断这个文档是什么类型”
- “生成 RAG 知识包”
- “批量生产知识库材料”

注意：`/pdf-compare` 是 Claude Code 兼容方式；Codex 标准触发用 `$document-skill`。

## 使用

在仓库根目录运行：

```powershell
python scripts\parse_document_compare.py compare "D:\documents\report.pdf" --max-pages 30 --output-format md
```

也可以用旧入口：

```powershell
python parse_pdf_compare.py "D:\documents\report.pdf" --max-pages 30
```

示例：

```powershell
# PDF 抽样对比
python scripts\parse_document_compare.py compare "D:\documents\report.pdf" --start-page 100 --max-pages 10

# 智能流水线
python scripts\parse_document_compare.py auto "D:\documents\report.pdf" --profile auto

# 智能流水线 + OCR 回退 + layout
python scripts\parse_document_compare.py auto "D:\documents\report.pdf" --auto-ocr --ocr-pages 1-3 --layout

# 发票字段抽取
python scripts\parse_document_compare.py extract-fields "D:\documents\invoice.pdf" --profile invoice --format json

# 发票字段校验
python scripts\parse_document_compare.py verify-fields "D:\documents\invoice_fields.json" --profile invoice --strict

# 发票批量汇总导出 Excel
python scripts\parse_document_compare.py export-xlsx "D:\documents\invoices" --profile invoice --recursive -o "D:\documents\invoice_summary.xlsx"

# PDF layout 坐标
python scripts\parse_document_compare.py layout-json "D:\documents\report.pdf" --max-pages 3 -o "D:\documents\layout.json"

# 文档分类和策略
python scripts\parse_document_compare.py classify "D:\documents\report.pdf" --max-pages 3

# 单文件知识包
python scripts\parse_document_compare.py knowledge-pack "D:\documents\report.pdf" --chunk-by page --out-dir "D:\documents\report_pack"

# 批量知识包
python scripts\parse_document_compare.py batch-knowledge "D:\documents" --recursive --out-dir "D:\documents\packs"

# Word 转 JSON
python scripts\parse_document_compare.py convert "D:\documents\report.docx" --parser markitdown --format json -o "D:\documents\report.json"

# 批量转换
python scripts\parse_document_compare.py batch "D:\documents" --ext .pdf,.docx --parser markitdown --format txt

# 批量质量扫描
python scripts\parse_document_compare.py scan-dir "D:\documents" --ext .pdf --max-pages 3 --min-quality 0.6 --fail-on-bad

# PDF 表格提取，三种格式全出
python scripts\parse_document_compare.py tables "D:\documents\report.pdf" --pages 1-5 --format all

# 检查 PDF 解析器依赖和 OCR 依赖
python scripts\parse_document_compare.py doctor --format pdf --ocr --json

# 提取 PDF 元数据
python scripts\parse_document_compare.py metadata "D:\documents\report.pdf" --format json

# 输出 JSONL 分块
python scripts\parse_document_compare.py chunk "D:\documents\report.pdf" --parser pymupdf --format jsonl --chunk-size 2000 --overlap 200

# 按 PDF 页输出 JSONL 分块
python scripts\parse_document_compare.py chunk "D:\documents\report.pdf" --parser pymupdf --chunk-by page --format jsonl

# 渲染 PDF 页面为 PNG
python scripts\parse_document_compare.py render-pages "D:\documents\report.pdf" --pages 1-3 --dpi 150

# 真正 OCR
python scripts\parse_document_compare.py ocr "D:\documents\report.pdf" --pages 1 --lang chi_sim+eng --format md

# 记录 OCR 回退意图
python scripts\parse_document_compare.py compare "D:\documents\report.pdf" --max-pages 3 --ocr-fallback
```

## 输出

`compare` 默认在输入文件同级目录生成 `<文件名>_parse_output/`：

| 文件 | 说明 |
|---|---|
| `compare_report.md` | 人类可读的对比报告 |
| `compare_report.json` | 机器可读报告 |
| `<parser>.md` | Markdown 输出，默认生成 |
| `<parser>.txt` | 纯文本输出，使用 `--output-format txt` 或 `all` |
| `<parser>.json` | 结构化文本输出，使用 `--output-format json` 或 `all` |
| `auto_report.md/json` | `auto` 智能流水线决策报告 |
| `best.md/txt/json` | `auto` 推荐解析器的最佳文本输出 |
| `fields.json/md` | `auto` 或 `extract-fields` 的结构化字段输出 |
| `invoice_summary.xlsx` | `export-xlsx` 发票汇总工作簿 |
| `layout.json` | `layout-json` 或 `auto --layout` 的 PDF 坐标结构 |
| `page_map.json` | 页面预览、页码和字符数索引 |
| `manifest.json` | `knowledge-pack` 知识包清单 |
| `quality_report.json` | 知识包质量报告 |
| `*_metadata.json` | `metadata` 默认输出 |
| `*_chunks.jsonl` | `chunk` 默认输出，每行一个分块 |
| `scan_report.md/json/csv` | `scan-dir` 批量质量扫描报告 |
| `render_pages.json` | `render-pages` 页面截图清单 |
| `*_ocr.txt/md/json` | `ocr` 识别结果 |

报告包含解析器状态、耗时、字符数、行数、段落数、中文字符数、相似度和文本预览。

报告还包含：

- 推荐解析器
- 质量评分和质量标签
- `(cid:...)`、控制字符、近空输出等诊断提示

Markdown 是默认输出，适合 LLM 阅读；TXT 适合检索和简单文本处理；JSON 适合后续程序化处理。

`auto` 适合用户只说“帮我解析这个文档”的场景。它会自动选解析器、输出最佳文本、生成分块、判断类型；如果识别为发票，还会额外生成结构化字段。

`auto --auto-ocr` 会在普通抽取低于质量阈值时尝试 Tesseract OCR；如果 OCR 质量更好，会使用 OCR 文本继续后续流水线。

`extract-fields --profile invoice` 会输出发票号码、开票日期、购销方名称和税号、项目明细、金额、税额、价税合计、开票人，并做金额 + 税额 = 价税合计校验。

`verify-fields --profile invoice --strict` 会校验必填字段、金额税额关系、明细合计关系、发票号和税号格式。

`export-xlsx --profile invoice` 会生成两个 sheet：`invoices` 汇总主字段，`items` 汇总明细行。

`layout-json` 和 `knowledge-pack` 面向可追溯 RAG：chunk 可以保留来源文件、解析器、页码、质量报告和页面预览。

`--min-quality` 会根据最佳解析结果的 `quality_score` 判断是否达标；`--fail-on-bad` 会把 `empty/bad` 标签视为失败。`compare` 和 `scan-dir` 都支持质量门禁，失败时返回退出码 `2`。

`--ocr-fallback` 不直接执行 OCR。它用于在报告或分块 metadata 中明确标记“如果质量差，下一步应进入 OCR 流程”。

## 真正 OCR 是什么

真正 OCR 指的是：先把 PDF 页面渲染成图片，再用 OCR 引擎识别图片里的文字。它和普通 PDF 文本层抽取不同：

- 普通解析：读取 PDF 内部已有的文字对象，速度快，但遇到扫描版、字体映射异常、`(cid:...)` 或乱码时可能失败。
- 真正 OCR：看页面图片并重新识别文字，能处理扫描版或坏文本层，但更慢，也依赖 OCR 引擎和语言包。

当前 `ocr` 子命令使用 Tesseract，需要：

- Python 包：`pymupdf`, `pytesseract`, `pillow`
- 系统程序：Tesseract OCR，并确保 `tesseract` 在 PATH 中
- 中文识别：需要安装 `chi_sim` 等对应语言数据

## 依赖

按需安装：

```powershell
pip install -r requirements.txt
```

或只安装某类文件需要的库：

```powershell
pip install markitdown pymupdf pypdf pdfplumber pdfminer.six liteparse
pip install opendataloader-pdf  # 需要系统安装 Java
pip install python-docx python-pptx openpyxl beautifulsoup4
pip install pytesseract pillow
```

缺失的解析器会在报告中跳过，不会导致整个对比任务失败。

**opendataloader-pdf 特别说明**：该解析器基于 Java JAR，`pip install` 会安装 Python 包装器和 JAR 文件，但运行时需要系统有 `java` 命令（JDK/JRE）。

OCR 的系统 Tesseract 需要单独安装；仅安装 Python 包还不能完成真正 OCR。

## 兼容说明

`.claude/skills/pdf-compare.md` 仍保留 Claude Code 风格说明，方便其它环境复用。Codex 发布以 `SKILL.md` 和 `agents/openai.yaml` 为准。

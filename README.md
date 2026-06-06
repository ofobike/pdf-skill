# PDF Parse Skill - PDF 优先的多格式文档解析工具

这是一个可发布的 Codex skill，也可以作为独立命令行工具使用。它用于对同一个文档运行多个解析器，比较文本抽取质量，并帮助判断是否需要 OCR 或更干净的源文件。

定位：PDF 优先，兼容 Word、PPT、Excel、HTML、TXT、CSV、JSON、XML、图片、音频、EPub、ZIP 等格式的解析、转换和结构化提取。

## 触发规则

Codex 对话里可以显式触发：

```text
$pdf-parse-skill
```

也可以用自然语言触发。只要请求涉及本地文档解析、PDF 抽取质量、OCR、表格、分块、元数据、发票字段、知识包或文档差异，Codex 都应优先使用这个 skill。

典型触发语：

- “帮我解析这个 PDF”
- “对比这个 PDF 哪个解析器效果最好”
- “多个解析器投票，选出最好的 PDF 解析结果给客户”
- “检查这些 PDF 解析器是不是真的能跑，不只是已安装”
- “给客户一份最可靠的 PDF 转 Markdown/TXT/JSON 结果”
- “对这张发票 PDF 做多解析器投票，校验后给客户 Markdown/TXT/JSON”
- “这个 PDF 抽出来是 `(cid:...)` / 乱码 / 空白，看看要不要 OCR”
- “把 OCR 也纳入 PDF 投票，普通解析器和 OCR 结果一起选最佳”
- “这个 PDF 有复杂表格、跨页表格、合并单元格或无边框表格，请做表格质量评分和多表格解析器投票”
- “给每个发票字段标出来源 parser、页码和置信度”
- “把这个目录里的 PDF 批量生成 customer-pack，并给总索引”
- “docling 太慢，请加解析器超时和健康缓存”
- “按合同/银行流水/报价单/采购单/报表/年报 profile 识别并选择解析策略”
- “把这个 DOCX/PPTX/XLSX/HTML 转成 Markdown/JSON/TXT”
- “提取 PDF 表格 / metadata / layout 坐标 / 页面截图”
- “把文档按页或按字符切成 JSONL 分块”
- “生成 RAG 知识包”
- “提取或校验发票字段”
- “对比两个文档版本差异”

命令行触发：

```powershell
python scripts\parse_document_compare.py <command> <path> [options]
```

旧入口兼容：

```powershell
python parse_pdf_compare.py <path> [options]
```

无子命令时默认等同于 `compare`：

```powershell
python scripts\parse_document_compare.py "D:\documents\report.pdf" --max-pages 3
```

客户交付场景建议使用 `vote`，先抽样投票，确认质量后再全量投票：

```powershell
python scripts\parse_document_compare.py vote "D:\documents\report.pdf" --max-pages 3 --format all
python scripts\parse_document_compare.py vote "D:\documents\report.pdf" --probe-before-vote --max-pages 0 --format all
```

发票客户交付建议启用发票 profile 和客户输出：

```powershell
python scripts\parse_document_compare.py vote "D:\documents\invoice.pdf" --probe-before-vote --max-pages 0 --format all --profile invoice --customer
```

## 功能总览

- 解析质量诊断：多解析器对比 PDF 文本抽取效果，推荐最佳解析器，并诊断 `(cid:...)`、乱码、空白、控制字符、低质量文本层。
- 解析器运行时体检：`probe` 会用真实文件逐个测试解析器，区分依赖缺失、运行失败、质量不达标和可交付解析器。
- 多解析器投票交付：对同一个 PDF 跑多个解析器，按质量、共识、覆盖度、结构、重复惩罚投票，输出 `best.md/txt/json` 和 `vote_report.md/json`；可用 `--probe-before-vote` 先体检并自动剔除不可用或低质量解析器。
- OCR 入投票：`ocr-tesseract` 已注册为 PDF parser，可和 `markitdown`、`pymupdf4llm`、`docling`、`pymupdf` 等普通解析器一起进入 `vote/probe/customer-pack`，适合扫描件或坏文本层 PDF。
- 复杂 PDF 客户交付包：`customer-pack` 一次性生成最佳文本、客户稿、表格投票、layout 坐标、metadata、预检报告、投票报告和 `manifest.json`，适合公司资料、合同、报表和带复杂表格的 PDF。
- 批量客户交付包：`batch-customer-pack` 可一次处理目录内多个 PDF，每个 PDF 生成独立客户包，并输出 `index.json/index.md` 总索引。
- 文档格式转换：把 PDF、Word、PPT、Excel、HTML、文本、图片、音频、EPub、ZIP 转成 Markdown、TXT 或 JSON。
- PDF Markdown 增强：支持 `pymupdf4llm`、`docling`、`pspdfkit`/`pdf-to-markdown` CLI 这类更偏 Markdown/RAG 的解析器。
- OCR 路由：判断普通文本层抽取是否失败，可用 Tesseract 对 PDF 页面执行真正 OCR；`vote` 中也可直接让 OCR 参与竞争。
- PDF 表格与版面：提取表格，输出页面、块、行、span、字体和坐标，支持页面渲染截图。
- 表格增强：`table-vote` 同时运行 `pdfplumber`、`pdfplumber-text`、`pymupdf-text`，按表格密度、行列一致性、表头完整度、跨方法共识投票，输出 `table_vote_report.*` 和 `best_tables.*`。
- 元数据提取：读取文件 metadata；PDF 会额外输出页数、内嵌 metadata、目录预览和文本层抽样判断。
- 分块与知识包：按字符或 PDF 页分块，生成可追溯 RAG/知识包、manifest、page map、质量报告。
- 批处理与质量门禁：批量转换目录、扫描解析质量，支持 `--min-quality` 和 `--fail-on-bad`。
- 字段级置信度：发票客户交付 JSON/Markdown/TXT 会给每个关键字段附带来源 parser、支持 parser 数、估算页码、坐标占位和置信度。
- 结构化字段：支持发票字段抽取、金额/税额/明细校验、明细重复检查、批量导出 Excel；`vote --profile invoice --customer` 可输出 `customer_best.md/txt/json`。
- 业务 profile：发票最强；另外可识别合同、银行流水、报价单、采购单、报表、公司年报等 profile，用于分类、策略和客户包处理。
- 解析器超时与健康缓存：`vote/probe/customer-pack/batch-customer-pack` 支持 `--timeout` 杀掉超时子进程，支持 `--parser-health-cache` 跳过 24 小时内同文件同页段近期失败或超时的解析器。
- 本地问答与差异：对知识包、chunks 或文档执行本地抽取式 QA；对比两个文档版本的文本、分类和字段差异。
- 依赖自检：检查 Python 解析器、OCR、Java、外部 CLI 等依赖，并给出安装建议。

## 支持格式与解析器

| 格式 | 扩展名 | 解析器 |
|---|---|---|
| PDF | `.pdf` | `markitdown`, `pymupdf4llm`, `docling`, `pspdfkit`/`pdf-to-markdown` CLI, `pymupdf`, `pypdf`, `pdfplumber`, `pdfminer`, `liteparse`, `opendataloader` (需 Java), `ocr-tesseract` (需 Tesseract) |
| Word | `.docx`, `.doc` | `markitdown`, `python-docx` |
| PPT | `.pptx`, `.ppt` | `markitdown`, `python-pptx` |
| Excel | `.xlsx`, `.xls` | `markitdown`, `openpyxl` |
| HTML | `.html`, `.htm` | `markitdown`, `beautifulsoup4` |
| 文本 | `.txt`, `.csv`, `.json`, `.xml`, `.md` | `markitdown` |
| 图片 | `.jpg`, `.png`, `.gif`, `.bmp`, `.tiff`, `.webp` | `markitdown` (EXIF + OCR，需 Tesseract) |
| 音频 | `.mp3`, `.wav` | `markitdown` (EXIF + 语音转录，需 whisper) |
| 电子书 | `.epub` | `markitdown` |
| 压缩包 | `.zip` | `markitdown` (遍历内部文件) |

## 命令选择

- `auto`：智能流水线，一次生成 metadata、最佳文本、分块、分类、字段和报告。
- `auto --auto-ocr`：质量低于阈值时自动尝试 OCR 回退。
- `compare`：多解析器对比，生成 Markdown/JSON 报告，并可输出 `md/txt/json/all` 格式的解析结果。
- `vote`：PDF 多解析器投票，输出最佳文本和投票报告；支持 `--profile invoice` 用发票校验加权，支持 `--customer` 额外生成客户交付稿，支持 `--probe-before-vote` 先体检解析器再投票。`ocr-tesseract` 可作为 OCR parser 参与投票。
- `customer-pack`：复杂 PDF 客户交付包，一次输出 `best.*`、`customer_best.*`、`tables/*`、`layout.json`、`metadata.json`、`manifest.json` 和预检/投票报告；表格默认走 `table-vote`。
- `batch-customer-pack`：批量生成多个 PDF 的客户交付包，每个文件一个目录，额外写 `index.json/index.md` 总索引。
- `probe`：用真实文件逐个测试解析器运行状态，输出 `probe_report.md/json`，适合发现 `docling` 运行时拉模型失败这类问题。
- `convert`：用指定解析器把单个文档转成 `md/txt/json/all`。
- `batch`：批量转换目录里的文档。
- `scan-dir`：批量扫描目录里的文档解析质量，输出 Markdown/JSON/CSV 报告。
- `tables`：从 PDF 提取表格，输出 `md/csv/json/all`。
- `table-vote`：多表格解析器投票，输出表格质量评分、解析器共识和最佳表格 `best_tables.md/csv/json`。
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
- `qa`：基于知识包、chunks 文件或文档执行本地抽取式问答，输出引用和片段。
- `diff-docs`：对比两个文档版本的文本差异、分类差异和可选发票字段差异。
- `--min-quality / --fail-on-bad`：质量门禁，适合批处理或自动化流程。
- `--ocr-fallback`：标记需要 OCR 回退；真正 OCR 使用 `ocr` 子命令。
- `--timeout`：限制单个解析器运行秒数，超时会杀掉子进程并标记为 `timeout`。
- `--parser-health-cache / --health-cache`：启用解析器健康缓存，跳过 24 小时内同文件同页段近期失败或超时的解析器。

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

## 自然语言交互提示语

下面这些提示语可以直接在 Codex 对话里使用；带文件路径时把示例路径换成实际文件或目录即可。

### 可靠客户交付

- “给客户一份最可靠的这个 PDF 转 Markdown/TXT/JSON 结果。”
- “先体检解析器，再多解析器投票，选出最好的 PDF 解析结果给客户。”
- “请对这个 PDF 先做解析器运行时体检，自动剔除不可用或低质量解析器，然后全量多解析器投票，输出最可靠的 Markdown、TXT 和 JSON 结果。”
- “这个 PDF 用所有可用解析器跑一遍，自动剔除不可用解析器，然后输出 best.md、best.txt、best.json。”
- “这个 PDF 里有复杂表格和公司数据，请生成客户交付包，包含最佳文本、表格、layout 坐标、metadata 和 manifest。”
- “这个 PDF 里有复杂表格和公司数据，请先做表格投票，再把最佳文本、最佳表格、layout、metadata 打成客户交付包。”
- “把这个目录里的 PDF 批量生成客户交付包，每个 PDF 一个目录，再给我总索引。”
- “请用 `$pdf-parse-skill` 对这个 PDF 做 `--probe-before-vote` 投票交付。”
- “我需要能交付给客户的结果，不要只给解析预览，要给 Markdown/TXT/JSON 文件。”

### 发票 PDF

- “给客户一份最可靠的这张发票 PDF 转 Markdown/TXT/JSON 结果。”
- “对这张发票 PDF 先做解析器体检，再投票，最后输出结构化客户交付稿。”
- “这张发票用发票校验加权，金额、税额、价税合计要校验，通过后给 customer_best.md/txt/json。”
- “帮我抽取这张发票的发票号码、开票日期、购销方、金额、税额、价税合计和明细。”
- “这张发票解析结果可能重复，请做重复惩罚后再选最佳解析器。”
- “这张发票请给每个字段标出来源 parser、页码、置信度和支持解析器数量。”
- “把这个目录里的发票批量抽字段并导出 Excel 汇总。”

### PDF 解析质量诊断

- “对比这个 PDF 哪个解析器效果最好，并说明为什么。”
- “这个 PDF 抽出来是 `(cid:...)`、乱码或空白，帮我判断是不是需要 OCR。”
- “检查这个 PDF 有没有文本层，普通解析能不能用。”
- “用前 3 页做抽样对比，质量没问题再全量解析。”
- “请生成 compare_report.md/json，并告诉我推荐解析器和质量分。”

### 解析器依赖与运行时体检

- “检查 PDF 解析器依赖是否都安装好了。”
- “这些解析器不只是要能导入，还要用这个真实 PDF 跑一下体检。”
- “帮我 probe 这个 PDF，看看 markitdown、pymupdf4llm、docling、pymupdf、pdfplumber、liteparse 哪些 ready。”
- “刚安装了 docling/pymupdf4llm/Tesseract，请用真实文件验证能不能跑。”
- “docling 有时很慢，请 probe/vote 时给每个解析器设置 120 秒超时，并启用健康缓存。”
- “这次投票请跳过最近失败或超时过的解析器。”
- “PSPDFKit/pdf-to-markdown 在这台机器能用吗？请 doctor 和 probe 一下。”

### 转换格式

- “把这个 PDF 转成 Markdown。”
- “把这个 DOCX/PPTX/XLSX/HTML 转成 Markdown/JSON/TXT。”
- “用 pymupdf4llm 把这个 PDF 转成适合 LLM/RAG 的 Markdown。”
- “试一下 docling 的结构化 Markdown 输出。”
- “把这个目录里的 PDF 和 DOCX 批量转成 TXT。”

### OCR 与扫描件

- “这个扫描版 PDF 做真正 OCR，输出 Markdown。”
- “先普通解析，质量低于 0.5 就自动尝试 OCR。”
- “对 PDF 第 1-3 页做 OCR，语言用中文和英文。”
- “把 PDF 页面先渲染成图片，再 OCR。”
- “解析不好就走 OCR，并在报告里说明普通解析为什么失败。”
- “把 OCR 也当成一个解析器参与 vote，不要只在最后回退 OCR。”

### 表格、版面和页面截图

- “提取这个 PDF 里的表格，输出 CSV 和 JSON。”
- “这个 PDF 的表格可能跨页、合并单元格、无边框，请做表格质量评分和多表格解析器投票。”
- “分别用 pdfplumber、pdfplumber-text、pymupdf-text 抽表格，然后投票选最佳表格。”
- “输出这个 PDF 的 layout 坐标，包括 block、line、span、字体和页面位置。”
- “这个 PDF 同时提取表格和版面坐标，并和最佳文本一起打成客户交付包。”
- “把 PDF 第 1-3 页截图出来看看。”
- “提取 PDF metadata、页数、目录和文本层抽样信息。”
- “我要核对版面来源，请生成 layout.json 和 page_map.json。”

### 分块、知识包和本地问答

- “把这个 PDF 按 2000 字分块输出 JSONL。”
- “把这个 PDF 按页分块，保留页码，输出 JSONL。”
- “为这个 PDF 生成可追溯知识包，后面做 RAG。”
- “基于这个知识包回答问题，并给出引用片段。”
- “批量生产这个目录的知识包。”

### 文档分类与差异对比

- “判断这个文档是什么类型，并推荐处理策略。”
- “按合同、银行流水、报价单、采购单、报表、年报这些 profile 识别这个 PDF。”
- “自动解析这个文档，能抽字段就抽字段。”
- “对比这两个 PDF 版本的文本差异。”
- “对比两张发票的关键字段差异。”
- “校验这个发票字段 JSON 是否金额一致。”

## 对话示例

显式触发：

```text
使用 $pdf-parse-skill 对比这个 PDF 的解析质量，并推荐最佳解析器。
```

```text
使用 $pdf-parse-skill，把 D:\documents\report.pdf 用 pymupdf 解析成 json。
```

```text
使用 $pdf-parse-skill，提取这个 PDF 的 metadata，并判断是否有文本层。
```

```text
使用 $pdf-parse-skill，把这个 PDF 按 2000 字分块输出 jsonl。
```

```text
使用 $pdf-parse-skill，把这个 PDF 按页分块输出 jsonl，保留页码。
```

```text
使用 $pdf-parse-skill，扫描 D:\documents 目录里的 PDF，质量低于 0.6 就失败。
```

```text
使用 $pdf-parse-skill，把这个 PDF 的第 1-3 页渲染成图片。
```

```text
使用 $pdf-parse-skill，对这个 PDF 第 1 页做 OCR，输出 Markdown。
```

```text
使用 $pdf-parse-skill，自动解析这个 PDF，判断类型并输出最佳文本和字段。
```

```text
使用 $pdf-parse-skill，抽取这张发票的结构化字段，输出 JSON。
```

```text
使用 $pdf-parse-skill，把这个目录里的发票汇总导出 Excel。
```

```text
使用 $pdf-parse-skill，自动解析这个 PDF，质量差就 OCR，并输出 layout。
```

```text
使用 $pdf-parse-skill，校验这个发票字段 JSON 是否金额一致。
```

```text
使用 $pdf-parse-skill，识别这个文档类型并推荐处理策略。
```

```text
使用 $pdf-parse-skill，为这个 PDF 生成可追溯知识包，后面做 RAG。
```

```text
使用 $pdf-parse-skill，检查 PDF 解析器依赖是否都安装。
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
- “这张发票 PDF 多解析器投票后给客户 Markdown/TXT/JSON”
- “把这些发票批量导出到 Excel”
- “解析不好就自动 OCR”
- “输出这个 PDF 的 layout 坐标”
- “校验发票字段是否一致”
- “判断这个文档是什么类型”
- “生成 RAG 知识包”
- “批量生产知识库材料”

注意：`/pdf-compare` 是旧版/Claude Code 兼容叫法；Codex 标准触发用 `$pdf-parse-skill`。

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

# PDF 多解析器投票，输出客户交付稿
python scripts\parse_document_compare.py vote "D:\documents\report.pdf" --probe-before-vote --max-pages 3 --format all

# PDF 多解析器投票，OCR 也作为 parser 参与，并限制慢解析器超时
python scripts\parse_document_compare.py vote "D:\documents\report.pdf" --probe-before-vote --max-pages 0 --format all --parsers markitdown,pymupdf4llm,docling,pymupdf,pdfplumber,ocr-tesseract --timeout 120 --parser-health-cache

# 复杂 PDF 客户交付包：最佳文本 + 表格 + layout + metadata + manifest
python scripts\parse_document_compare.py customer-pack "D:\documents\report.pdf" --max-pages 0 --table-pages all --layout-max-pages 30

# 批量复杂 PDF 客户交付包：每个 PDF 一个包 + 总索引
python scripts\parse_document_compare.py batch-customer-pack "D:\documents\pdfs" --recursive --max-pages 0 --table-pages all --layout-max-pages 30 --timeout 120 --parser-health-cache

# 解析器运行时体检
python scripts\parse_document_compare.py probe "D:\documents\report.pdf" --max-pages 1 --keep-outputs

# 解析器运行时体检 + 超时 + 健康缓存
python scripts\parse_document_compare.py probe "D:\documents\report.pdf" --max-pages 1 --timeout 120 --parser-health-cache --keep-outputs

# 发票 PDF 投票 + 校验加权 + 自动客户交付
python scripts\parse_document_compare.py vote "D:\documents\invoice.pdf" --probe-before-vote --max-pages 0 --format all --profile invoice --customer

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

# 本地抽取式问答（不调用 LLM）
python scripts\parse_document_compare.py qa "D:\documents\report_pack" "合同金额是多少？" --format md

# 文档版本差异对比
python scripts\parse_document_compare.py diff-docs "D:\documents\old.pdf" "D:\documents\new.pdf" --format all

# Word 转 JSON
python scripts\parse_document_compare.py convert "D:\documents\report.docx" --parser markitdown --format json -o "D:\documents\report.json"

# 批量转换
python scripts\parse_document_compare.py batch "D:\documents" --ext .pdf,.docx --parser markitdown --format txt

# 批量质量扫描
python scripts\parse_document_compare.py scan-dir "D:\documents" --ext .pdf --max-pages 3 --min-quality 0.6 --fail-on-bad

# PDF 表格提取，三种格式全出
python scripts\parse_document_compare.py tables "D:\documents\report.pdf" --pages 1-5 --format all

# PDF 表格质量评分 + 多表格解析器投票
python scripts\parse_document_compare.py table-vote "D:\documents\report.pdf" --pages all --format all

# 检查 PDF 解析器、OCR 和 opendataloader-pdf/Java 依赖
python scripts\parse_document_compare.py doctor --format pdf --ocr --opendataloader --json

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
| `vote_report.md/json` | `vote` 多解析器投票决策报告；使用 `--probe-before-vote` 时会记录预检摘要 |
| `probe_report.md/json` | `probe` 解析器运行时体检报告 |
| `preflight_probe/probe_report.md/json` | `vote --probe-before-vote` 自动生成的投票前解析器体检报告 |
| `best.md/txt/json` | `vote` 或 `auto` 推荐给客户/下游使用的最佳原始文本 |
| `customer_best.md/txt/json` | `vote --customer` 的客户交付稿；发票 profile 下是结构化发票字段和校验结果 |
| `manifest.json` / `README.md` | `customer-pack` 客户交付包清单和说明 |
| `tables/table_vote_report.md/json` | `customer-pack` 或 `table-vote` 的表格解析器投票报告 |
| `tables/best_tables.md/csv/json` | `customer-pack` 选出的最佳表格输出 |
| `*_tables.md/csv/json` | `tables --format all` 的单解析器 PDF 表格输出 |
| `index.json` / `index.md` | `batch-customer-pack` 批量客户包总索引 |
| `auto_report.md/json` | `auto` 智能流水线决策报告 |
| `fields.json/md` | `auto` 或 `extract-fields` 的结构化字段输出 |
| `invoice_summary.xlsx` | `export-xlsx` 发票汇总工作簿 |
| `layout.json` | `layout-json` 或 `auto --layout` 的 PDF 坐标结构 |
| `page_map.json` | 页面预览、页码和字符数索引 |
| `manifest.json` | `knowledge-pack` 知识包清单 |
| `quality_report.json` | 知识包质量报告 |
| `qa_report.md/json` | `qa` 本地抽取式问答报告 |
| `diff_report.md/json` | `diff-docs` 文档差异报告 |
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

`vote` 会用多个 PDF 解析器同时抽取文本并打分。默认 `--profile auto` 会自动识别是否是发票；`--profile none` 只做通用文本投票；`--profile invoice` 会把发票字段完整度、金额税额校验、明细合计、明细重复检查纳入投票。`--profile contract/bank_statement/quotation/purchase_order/report/annual_report` 可显式指定业务类型，目前主要用于分类、策略和报告。`--probe-before-vote` 会先跑真实文件小样本体检，只让 `ready` 解析器进入正式投票，适合客户交付前避开不可用、运行失败或低质量解析器。`--customer` 会额外输出 `customer_best.md/txt/json`，发票场景下优先给结构化、校验后的客户交付稿。

`ocr-tesseract` 是 PDF parser。把它加入 `--parsers` 后，OCR 文本会和普通文本层解析器一起进入 `probe/vote/customer-pack`。这适合扫描件、坏文本层、`(cid:...)` 或普通解析结果近空的 PDF。

`vote --profile invoice --customer` 的 `customer_best.json` 会包含 `field_confidence`：每个发票字段都有 winning value、来源 parser、支持 parser 数、估算页码、坐标占位和置信度。坐标只有在后续接入字段到 layout 的精确匹配时才会填充；当前无法可靠定位时为 `null`。

`table-vote` 会对同一 PDF 跑多个表格解析策略：`pdfplumber` 偏有边框/标准表格，`pdfplumber-text` 偏无边框或文本对齐表格，`pymupdf-text` 是文本行启发式兜底。它按表格密度、行列一致性、表头完整度、覆盖度和跨方法共识评分，输出 `table_vote_report.md/json` 和 `best_tables.md/csv/json`。复杂跨页表格仍建议人工抽样核查，但报告会明确哪个方法胜出。

`customer-pack` 默认把表格提取升级为表格投票，因此交付包中的 `tables/` 目录会包含 `table_vote_report.*` 和 `best_tables.*`。`batch-customer-pack` 会对目录内 PDF 批量执行同样流程，并写 `index.json/index.md`。

`--timeout` 会把单个解析器放进子进程运行，到秒数后杀掉该子进程并在报告里标为 `timeout`。`--parser-health-cache` 会记录同文件、同页段、同解析器的近期失败/超时，24 小时内自动跳过，适合 docling、OCR 或外部 CLI 这类慢解析器。

`probe` 适合安装依赖后或交付前验证解析器是否真的可用。它会用真实文件跑解析器小样本，把结果分为 `ready`、`dependency_missing`、`runtime_failed` 和 `quality_failed`；这能发现“依赖能导入但运行时需要下载模型/外部命令失败”的情况。

`extract-fields --profile invoice` 会输出发票号码、开票日期、购销方名称和税号、项目明细、金额、税额、价税合计、开票人，并做金额 + 税额 = 价税合计校验。

`verify-fields --profile invoice --strict` 会校验必填字段、金额税额关系、明细合计关系、发票号和税号格式。

`export-xlsx --profile invoice` 会生成两个 sheet：`invoices` 汇总主字段，`items` 汇总明细行。

`layout-json` 和 `knowledge-pack` 面向可追溯 RAG：chunk 可以保留来源文件、解析器、页码、质量报告和页面预览。

`qa` 是本地抽取式检索问答，不调用 LLM；它会返回匹配片段、答案句和引用信息，适合对知识包或 `chunks.jsonl` 做快速定位。

`diff-docs` 会解析两个文档，输出文本相似度、行级 diff、分类信息；当 profile 为 `auto` 且识别为发票时，会额外比较关键发票字段。

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
pip install markitdown pymupdf4llm docling pymupdf pypdf pdfplumber pdfminer.six liteparse
pip install opendataloader-pdf  # 需要系统安装 Java
npm install -g @pspdfkit/pdf-to-markdown  # PSPDFKit/Nutrient CLI，当前主要支持 macOS/Linux
pip install python-docx python-pptx openpyxl beautifulsoup4
pip install pytesseract pillow
```

缺失的解析器会在报告中跳过，不会导致整个对比任务失败。

**opendataloader-pdf 特别说明**：该解析器基于 Java JAR，`pip install` 会安装 Python 包装器和 JAR 文件，但运行时需要系统有 `java` 命令（JDK/JRE）。

OCR 的系统 Tesseract 需要单独安装；仅安装 Python 包还不能完成真正 OCR。

## 验证与发布打包

基础验证：

```powershell
python -m py_compile scripts\parse_document_compare.py parse_pdf_compare.py scripts\package_skill.py
python -m unittest discover -s tests
```

生成干净的可发布 skill 目录：

```powershell
python scripts\package_skill.py --force
```

默认输出到 `dist\pdf-parse-skill`，只包含 `SKILL.md`、`agents/`、`scripts/`、`references/`、兼容入口和依赖清单。

## 兼容说明

旧版 `/pdf-compare` 说明只作为历史兼容语义保留。Codex 发布以 `SKILL.md` 和 `agents/openai.yaml` 为准。

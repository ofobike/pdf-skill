---
description: >
  多格式文档解析工具：对比解析器、文档转 Markdown、批量转换、表格提取。
  支持格式：PDF (.pdf)、Word (.docx)、PPT (.pptx)、Excel (.xlsx)、HTML (.html)。
  触发条件：用户要求"对比解析器"、"测试文档提取"、"比较解析库"、"文档转 Markdown"、
  "批量转换 PDF"、"提取 PDF 表格"、"PDF/Word/PPT/Excel 解析"。
  用法：/pdf-compare compare|convert|batch|tables <文件> [选项]
  跳过：单纯的文档阅读、文档创建、与解析无关的操作。
---

# 多格式文档解析工具

你正在执行文档解析任务。本工具有四个子命令，根据用户意图选择合适的子命令。

## 子命令速查

| 子命令 | 用途 | 示例 |
|--------|------|------|
| `compare` | 多解析器对比（默认） | `python script.py compare report.pdf` |
| `convert` | 文档转 Markdown | `python script.py convert report.pdf --parser pymupdf` |
| `batch` | 批量转换目录 | `python script.py batch ./pdfs/ --parser pymupdf` |
| `tables` | PDF 表格提取 | `python script.py tables report.pdf --pages 1-5` |

## 判断用户意图

- 用户说"对比"、"比较"、"哪个好" → `compare`
- 用户说"转 Markdown"、"提取文本"、"转为 md" → `convert`
- 用户说"批量"、"全部转换"、"整个目录" → `batch`
- 用户说"表格"、"提取表格"、"导出表格" → `tables`
- 没有明确意图，默认 `compare`

## 支持格式

| 格式 | 扩展名 | 可用解析器 |
|------|--------|-----------|
| PDF | .pdf | markitdown, pymupdf, pypdf, pdfplumber, pdfminer, liteparse |
| Word | .docx | markitdown, python-docx |
| PPT | .pptx | markitdown, python-pptx |
| Excel | .xlsx | markitdown, openpyxl |
| HTML | .html | markitdown, beautifulsoup4 |

## 执行步骤

### 1. 定位脚本

按以下顺序查找：
- 当前项目下的 `pdf-skill/parse_pdf_compare.py`
- 用户主目录下的 `~/pdf-skill/parse_pdf_compare.py`
- 当前工作目录下的 `parse_pdf_compare.py`

### 2. 验证输入

- 确认文件/目录存在
- 确认格式受支持
- 确认 Python 环境可用

### 3. 安装依赖（按需）

```bash
pip install markitdown pymupdf pypdf pdfplumber pdfminer.six liteparse  # PDF
pip install markitdown python-docx                                       # Word
pip install markitdown python-pptx                                       # PPT
pip install markitdown openpyxl                                          # Excel
pip install markitdown beautifulsoup4                                    # HTML
```

### 4. 运行对应子命令

```bash
# 对比
python "<脚本>" compare "<文件>" [--parsers pymupdf,pdfplumber] [--parallel]

# 转换
python "<脚本>" convert "<文件>" --parser pymupdf -o output.md

# 批量
python "<脚本>" batch "<目录>" --parser pymupdf --output-dir ./out/

# 表格
python "<脚本>" tables "<PDF>" --pages 1-5 --format csv
```

### 5. 展示结果

- compare：展示对比报告，分析各解析器优劣
- convert：展示输出文件路径和字符数
- batch：展示成功/失败统计
- tables：展示提取到的表格数量和输出路径

### 6. 结果分析（compare 模式）

- 哪个解析器提取字符数最多
- 哪个解析器速度最快
- 各解析器相似度如何
- 针对该文件类型的推荐建议

## 约束

- 不要修改原始文件
- 如果某个库未安装，报告为缺失依赖，不要自动安装
- 保持客观，不偏向任何解析器
- PDF 默认只解析前 30 页（compare/convert），batch 默认全量

# mineru-pdf2docx

基于 [MinerU](https://github.com/opendatalab/MinerU) 的 **PDF → Word (.docx) 导出工具**。

MinerU 解决了 PDF 解析的四大难点（双栏/竖排版面重排、无边框表格、矢量图表、公式识别），但它只输出 markdown；从 markdown 到 Word 的最后一公里存在三个断点，导致**表格丢失、大图被拆碎、公式降级**。本项目把这条路补通了。

## 基于的原项目

| | |
|---|---|
| 原项目 | [opendatalab/MinerU](https://github.com/opendatalab/MinerU)（PDF/文档智能解析引擎） |
| 原项目规模 | 82k+ stars，Apache 2.0 + 附加条款（MinerU Open Source License） |
| 依赖 | `mineru>=4.0,<5`、`pymupdf`、`pypandoc-binary` |

> ⚠️ 许可证说明：MinerU 采用自有的 MinerU Open Source License（非纯 Apache 2.0）。个人/内部使用无障碍，**商用前请阅读其仓库根目录 `LICENSE.md` 附加条款**。本项目自身代码以 MIT 发布。

## 我们做了什么（二创部分）

没有修改 MinerU 源码，全部做在外围集成层：

### 1. 命令行封装（`parse_pdf.py`）
MinerU 只有裸的 `mineru-kit parse`，本项目提供：
- 批量输入：单文件 / 目录 / 列表文件（`--list`）
- 关键词排除（`--exclude`）
- 解析档位（`--tier`）、页码范围（`-p`）透传
- 自动解压、按文件名建目录、块类型统计
- `--docx` 一键导出 Word

### 2. 修复 pandoc 链路的三个断点（核心）

| # | 断点 | 现象 | 修复 |
|---|---|---|---|
| 1 | **HTML 表格被 pandoc 丢弃** | MinerU 的无线框表格输出 HTML `<table>`，pandoc 的 markdown 模式静默丢弃表格结构（只留文字）。实测某篇 8 个表全丢 | `html_tables_to_pipe()`：HTML 表格 → 管道符表格，处理 colspan/rowspan/单元格内换行/`\|` 转义 |
| 2 | **多面板大图被拆碎** | 版面模型把一张 Fig.1 的 4 个面板识别为 4 个独立 image 块，Word 里变成一堆小图 | `merge_split_figures()`：依据 `structured_content.json` 的 bbox，将相邻图块（左右间距 <5% 页宽且垂直重叠 >50%，或上下间距 <3% 页高且水平重叠 >50%，中间无文字块）合并，从原 PDF 以 200 DPI 重渲染为整图 |
| 3 | **pandoc 不认旧式 LaTeX 宏** | `\AA`、`\tt`、`\sf`、`\cal`、`\cfrac`、`\tag{N}`、`\sp{}`、`\vphantom`、`\tiny` 等导致公式降级为纯文本 | `_fix_tex_macros()`：12 种宏兼容替换（`\tag{N}` 的方程编号以 `(N)` 形式保留） |

### 3. 工程化配套
- 独立 venv 安装指南（不污染科研环境）
- `refix_batch.py`：对已解析输出批量重应用修复（无需重跑 MinerU）
- `check_docx.py` / `verify_fix.py`：docx 结构抽检、修复效果核对
- `AGENTS.md`：选型理由与踩坑记录

## 安装

```bash
# 1. 创建独立环境（Python >= 3.10, < 3.15）
python -m venv .venv

# 2. 安装依赖（mineru 基础包 = ONNX + llama.cpp，CPU 可跑）
.venv/Scripts/pip install "mineru>=4.0,<5" pymupdf pypandoc-binary   # Windows
# .venv/bin/pip install "mineru>=4.0,<5" pymupdf pypandoc-binary     # Linux/macOS
```

无 NVIDIA 独显的机器用默认 `basic` 档（约 7~10 秒/页）；有 8GB+ 显存可用 `--tier standard` 提升质量。

## 使用

```bash
# 单文件：解析 + 导出 Word
python parse_pdf.py paper.pdf --docx

# 整个目录批量（排除扫描书等）
python parse_pdf.py ./pdfs --docx --exclude 扫描 --exclude 发票

# 指定页码 / 档位
python parse_pdf.py paper.pdf --docx -p 1-5 --tier standard

# 列表文件批量（每行一个 PDF 路径）
python parse_pdf.py --list papers.txt --docx -o ./output
```

## 输出结构

```
output/
└── <论文名>/
    ├── markdown.md            # MinerU 解析结果（修复后）
    ├── markdown.docx          # 最终 Word 文档
    ├── markdown_original.md   # 修复前原始 markdown（refix_batch.py 生成时备份）
    ├── images/                # 提取/合并的图片（fig_p*_*.png 为合并整图）
    ├── middle_json.json       # 带坐标的结构化数据（二次开发用）
    └── structured_content.json
```

**排版保留情况**：公式 → Word 原生方程（可编辑）；表格 → Word 原生表格（含合并单元格）；图片 → 内嵌并按原文顺序插入；标题层级、上下标、粗斜体、阅读顺序完整。双栏论文内容重排为单栏（Word 无法还原物理分栏，顺序不丢）。

## 实测数据

15 篇 DFT/化学论文（229 页，`basic` 档 CPU）：

| 论文 | 原生方程 | Word 表格 | 图片 |
|---|---|---|---|
| OC22 Dataset（50页） | 319 | 20 | 14 |
| becker 双电层（26页） | 676 | 0 | 35 |
| CHGNet | 174 | 0 | 8 |
| DeepGlycanSite | 68 | 1 | 7 |

- 表格修复：28 → **44/44** 全部保留
- 大图合并：图片数下降 40~50%
- pandoc 转换零警告

## 已知限制

- `--pages` 仅适用于 PDF；Office 文档仅支持 flash 档
- 纯文本文件（.txt/.md）MinerU 不解析，直接读取
- 扫描件可 OCR，但 CPU 上约 10 秒/页（507 页扫描书约需 1.5 小时+）
- Windows GPU 加速需单独安装 GPU 版 PyTorch

## 文件说明

| 文件 | 作用 |
|---|---|
| `parse_pdf.py` | 工具本体：封装 MinerU + 三个修复 + docx 导出 |
| `refix_batch.py` | 批量修复入口（对已解析输出重新应用修复） |
| `check_docx.py` | docx 结构抽检（方程/表格/图片计数） |
| `verify_fix.py` | 表格/图片修复效果核对 |
| `AGENTS.md` | 选型理由与踩坑记录 |

## 致谢

- [MinerU](https://github.com/opendatalab/MinerU) — 底层 PDF 解析引擎
- [pandoc](https://pandoc.org/) — markdown → docx 转换

## License

本项目代码以 [MIT](LICENSE) 发布；依赖的 MinerU 遵循其自有的 MinerU Open Source License。

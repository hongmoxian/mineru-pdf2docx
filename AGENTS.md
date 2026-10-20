# PDF 解析工具 — AGENTS.md

## 项目目标

解决 PDF 解析的四个硬骨头问题（在文本型 PDF 上）：

1. **复杂版面乱序** — 双栏论文 `get_text("text")` 左右栏文字交错；竖排文字被拆成单字
2. **无边框表格 / 合并单元格** — pdfplumber 靠线条判断表格，无线条的表格失效
3. **矢量图表提取** — `get_images()` 只能取嵌入位图，矢量图需整页渲染裁剪
4. **公式识别** — PDF 里公式只是带坐标的字符，直接提取是乱码

## 选型结论：基于 MinerU 二创

**选定项目**：[opendatalab/MinerU](https://github.com/opendatalab/MinerU)

调研数据（2026-10-20，GitHub API 实测）：

| 项目 | Stars | 最近提交 | 许可证 | 四能力覆盖 |
|---|---|---|---|---|
| **MinerU** | 82,079 | 2026-10-19 | MinerU Open Source License（基于 Apache 2.0 + 附加条款） | 全部 ✅ |
| marker | 40,549 | 2026-10-19 | Apache-2.0 | 全部 ✅ |
| PaddleOCR | 91,397 | 2026-09-16 | Apache-2.0 | 3.5/4 |
| unstructured | 15,578 | 2026-10-19 | Apache-2.0 | 2/4 |
| nougat | 10,085 | 2025-02-21 | MIT | 2/4 |
| table-transformer | 2,948 | 2024-06-24 | MIT | 已归档，排除 |

**为什么选 MinerU**：
- 四个问题（分栏阅读顺序、无边框表格、图像/图表、公式 LaTeX）原生全覆盖，无需自己拼装
- 第三方 benchmark 前列：OmniDocBench（arXiv 2412.07626）开源前二；INFINITY-Parser 论文对比表总分 96.6
- 82k star、昨天仍在提交，生态最活跃；四档质量可选（flash/basic/standard/advanced）
- 无 NVIDIA 独显的机器可走 CPU 路径（ONNX + llama.cpp 基础包）

**为什么排除其他**：
- `table-transformer`：仓库已归档，只做表格检测，能力已被上述项目内置方案取代
- `nougat`：一年半未提交（2025-02 后停更），只针对学术论文，不管表格和图片
- `unstructured`：公式识别和矢量图能力缺失，四个问题只覆盖一半
- `marker`：纯 Apache-2.0 且 benchmark 略高，但 open issues 494 个、CPU 路径支持弱于 MinerU
- `PaddleOCR`：能力足够强，但版面/表格/公式三件套需自行组合，工程复杂度高于 MinerU 的一键方案

## 使用约束

- **个人 / 内部学习使用无障碍；商用前必须阅读仓库根目录 `LICENSE.md` 附加条款**（许可证为 MinerU Open Source License，非纯 Apache 2.0，也非 AGPL）
- 隐私：默认本地解析，不传远程；远程解析需显式 `--remote`
- 本机无 NVIDIA GPU：使用 `basic`（ONNX 基础包，CPU）或 `standard`（llama.cpp + Vulkan）档；不使用 `mineru[full]`（需 8GB+ 显存）

## 环境

- 独立 venv：`e:/pdf-tool/.venv`（避免污染 miniconda 科研环境）
- Python >= 3.10, < 3.15
- 安装：`pip install "mineru>=4.0,<5"`（基础包：ONNX + llama.cpp，模型约 0.8GB）

## 目录结构

```
e:/pdf-tool/
├── AGENTS.md          # 本文件
├── parse_pdf.py       # 封装 CLI：PDF → Markdown + 图片 + 表格 + 公式
├── .venv/             # 独立虚拟环境
└── output/            # 解析输出（按 PDF 名分子目录）
```

## Word 导出（PDF -> docx）

流水线：MinerU 解析（zip）-> `markdown.md` -> pandoc -> `.docx`

- 公式 -> **Word 原生方程**（oMath），可编辑
- 表格 -> **Word 原生表格**（w:tbl）
- 图片/矢量图表 -> 内嵌图片，按原文顺序插入
- 阅读顺序、标题层级、上下标、粗斜体均保留
- 双栏论文内容重排为单栏（Word 无法还原原版面分栏，但顺序不丢）

pandoc 不识别 MinerU 输出的部分旧式 LaTeX 宏，`parse_pdf.py` 的 `_fix_tex_macros()` 做兼容替换：`\AA`→Å、`\bf`→`\mathbf`、`\tt`/`\sf`→`\mathrm`、`\cal`→`\mathcal`、`\cfrac`→`\frac`、`\tag{N}`→`\ \text{(N)}`（保留方程编号）、删除 `\vphantom`/`\tiny`/`\scriptsize`、`\sp{x}`→`^{x}`。修复后 pandoc 转换零警告。

### 已踩坑并修复的两个问题（2026-10-20）

1. **HTML 表格被 pandoc 丢弃**：MinerU 的 markdown 中表格有两种形态——HTML `<table>`（无线框表格识别输出）和管道符表格。pandoc 的 markdown 模式会静默丢弃 HTML 表格结构（只留文字，实测 tran 那篇 8 个表全丢）。修复：`html_tables_to_pipe()` 把 HTML 表格转成管道符表格（处理 colspan/rowspan/单元格内换行/`|` 转义）。修复后 15 篇论文 44/44 表格全部成为 Word 原生表格。
2. **多面板大图被拆成多个小图**：版面模型把一张大图的各面板识别为独立 image 块（如 CHGNet Fig.1 被拆成 4 块）。修复：`merge_split_figures()` 依据 `structured_content.json` 的 bbox，把相邻（左右间距<5%页宽且垂直重叠>50%，或上下间距<3%页高且水平重叠>50%）且中间无文字块隔开的 image/chart 块合并，从原 PDF 按坐标以 200 DPI 重渲染为整图，并改写 markdown 引用。修复后图片数下降约 40~50%（becker 68→35、tetrazine 27→12）。

批量修复入口：`refix_batch.py`（对已解析输出重新应用两个修复并再生成 docx，原 markdown 备份为 `markdown_original.md`）。

## 实测验证记录（2026-10-20）

| 测试文件 | 结果 |
|---|---|
| CHGNet 论文（双栏，4页） | 阅读顺序完美；摘要/正文公式全部转 LaTeX（如 `$\mathrm{Li_{x}MnO_{2}}$`）；Fig.1 面板 4 张位图 + Fig.3 矢量图表 2 个 chart 块均渲染为 JPG 并在 markdown 中引用 |
| 工发票（竖排+表格） | 竖排"购买方信息"完整不再拆字；表格输出 HTML `<table>`（colspan 合并单元格正确）；额外识别出发票印章文字 |
| DeepGlycanSite 论文 p4 | Table 1 输出为标准 Markdown 表格（表头、±、显著性上标、表脚完整） |

CPU 速度参考（basic 档，模型缓存后）：约 0.1~0.14 页/秒（单页 7~10 秒）。

## DFT 论文批量转换结果（2026-10-20）

`e:/download/` 下 15 篇论文（229 页）全部转换为 docx，输出在 `e:/pdf-tool/output/dft/<论文名>/markdown.docx`：

| 论文 | 原生方程 | 表格 | 图片 | 大小 |
|---|---|---|---|---|
| OC22 Dataset (50页) | 319 | 20 | 14 | 2.3 MB |
| CHGNet | 174 | 0 | 8 | 1.4 MB |
| becker EDL (26页) | 676 | 0 | 35 | 1.3 MB |
| DeepGlycanSite | 68 | 1 | 7 | 1.9 MB |
| 其余 11 篇 | — | — | — | 0.4~4.2 MB |

（表格/图片数为修复 HTML 表格丢弃、大图拆分两个问题后的最终值；原始 markdown 备份在各目录 `markdown_original.md`）

pandoc 宏修复后转换零警告。校验脚本：`check_docx.py`（docx 结构抽检）、`verify_fix.py`（表格/图片修复效果核对）。

## 已知限制（来自 MinerU README）

- `--pages` 参数仅适用于 PDF
- Office 文档仅支持 flash 质量档
- 纯文本文件（.txt/.md 等）不解析，直接读取
- Windows GPU 加速需单独安装 GPU 版 PyTorch（本机无独显，不涉及）

# -*- coding: utf-8 -*-
"""
parse_pdf.py — PDF 解析小工具（基于 MinerU 二创）

解决四个硬骨头问题：
  1. 双栏/竖排等复杂版面的阅读顺序重排
  2. 无边框表格 / 合并单元格 -> Markdown 或 HTML 表格
  3. 矢量图表 -> 渲染为图片导出
  4. 公式 -> LaTeX

用法：
  python parse_pdf.py <pdf文件或目录> [选项]

选项：
  -o, --output DIR   输出目录（默认 ./output）
      --tier STR     解析档位 flash/basic/standard/advanced（默认 basic，CPU 友好）
  -p, --pages STR    页码，如 '1-5,8'（默认全部）
      --docx         同时导出 Word (.docx)，公式转为 Word 原生方程，表格为 Word 原生表格
      --list FILE    从文本文件读取 PDF 路径列表（每行一个）
      --exclude STR  排除文件名包含该关键词的 PDF（可多次指定）
      --keep-zip     保留 MinerU 原始 zip 包
  -h, --help         显示帮助

示例：
  python parse_pdf.py paper.pdf --docx
  python parse_pdf.py ./pdfs --tier standard -p 1-3
  python parse_pdf.py ./pdfs --docx --exclude 扫描 --exclude 发票
"""
import argparse
import json
import shutil
import subprocess
import sys
import time
import zipfile
from pathlib import Path

VENV_PYTHON = Path(__file__).parent / ".venv" / "Scripts" / "python.exe"
MINERU_KIT = Path(__file__).parent / ".venv" / "Scripts" / "mineru-kit.exe"
DEFAULT_OUTPUT = Path(__file__).parent / "output"


def find_pdfs(input_path: Path, excludes: list[str], list_file: Path | None) -> list[Path]:
    """收集待解析的 PDF：单文件 / 多文件 / 目录 / 列表文件（可按关键词排除）"""
    if list_file:
        pdfs = [Path(line.strip()) for line in list_file.read_text(encoding="utf-8").splitlines()
                if line.strip() and not line.strip().startswith("#")]
    elif input_path.is_file():
        pdfs = [input_path]
    elif input_path.is_dir():
        pdfs = sorted(input_path.glob("*.pdf"))
    else:
        print(f"[错误] 路径不存在: {input_path}")
        sys.exit(1)
    if excludes:
        pdfs = [p for p in pdfs if not any(k in p.name for k in excludes)]
    return pdfs


def _fix_tex_macros(md: str) -> str:
    """pandoc 的 TeX 数学解析器不支持部分 LaTeX 宏，转换为等价写法。

    MinerU 公式识别输出中常见：\\AA（埃）、\\bf、\\tt、\\sf（旧字体宏）、
    \\tag{N}（方程编号）、\\cfrac、\\textless、\\sp{}^{}（上标简写）、
    \\vphantom、\\tiny/\\scriptsize（字号宏）。
    """
    import re

    def remove_macro_with_arg(text: str, macro: str) -> str:
        """删除 \\macro{...}，用括号计数正确处理嵌套花括号"""
        token = "\\" + macro
        while True:
            i = text.find(token)
            if i < 0:
                return text
            j = i + len(token)
            while j < len(text) and text[j] in " \t":
                j += 1
            if j >= len(text) or text[j] != "{":
                # 宏后面没有参数组，只删宏名本身
                text = text[:i] + text[j:]
                continue
            depth, k = 0, j
            while k < len(text):
                if text[k] == "{":
                    depth += 1
                elif text[k] == "}":
                    depth -= 1
                    if depth == 0:
                        break
                k += 1
            text = text[:i] + text[j + 1:k] + text[k + 1:]

    md = md.replace("\\AA", "\u00c5")            # 埃单位 -> Unicode
    md = md.replace("\\bf", "\\mathbf")           # 粗体
    md = md.replace("\\tt", "\\mathrm")           # 旧打字机体 -> 罗马体
    md = md.replace("\\sf", "\\mathrm")           # 旧 sans-serif -> 罗马体
    md = md.replace("\\cal", "\\mathcal")         # 旧花体宏
    md = md.replace("\\cfrac", "\\frac")          # 连续分式 -> 普通分式
    md = md.replace("\\textless", "<")            # HTML 实体名
    md = re.sub(r"\\tag\s*\{(\d+)\}", r"\\ \\text{(\1)}", md)  # 方程编号保留
    for macro in ("vphantom", "tiny", "scriptsize"):
        md = remove_macro_with_arg(md, macro)
    # \sp{x} 是上标简写：\sp{a} -> ^{a}（同样支持嵌套）
    while True:
        i = md.find("\\sp")
        if i < 0:
            break
        j = i + 3
        while j < len(md) and md[j] in " \t":
            j += 1
        if j >= len(md) or md[j] != "{":
            md = md[:i] + md[j:]
            continue
        depth, k = 0, j
        while k < len(md):
            if md[k] == "{":
                depth += 1
            elif md[k] == "}":
                depth -= 1
                if depth == 0:
                    break
            k += 1
        md = md[:i] + "^{" + md[j + 1:k] + "}" + md[k + 1:]
    md = md.replace("`", "")                      # 反引号在数学模式中非法
    return md


def html_tables_to_pipe(md: str) -> str:
    """把 MinerU markdown 中的 HTML <table> 转成管道符表格。

    pandoc 的 markdown 模式会丢弃 HTML 表格结构（只保留文字），
    管道符表格才能转成 Word 原生表格。处理 colspan / rowspan / 单元格内换行。
    """
    import re
    from html.parser import HTMLParser

    class TableParser(HTMLParser):
        def __init__(self):
            super().__init__()
            self.tables = []          # 每个元素是 list[list[dict]]（行 -> 单元格）
            self.cur_table = None
            self.cur_row = None
            self.cur_cell = None
            self.buf = []

        def handle_starttag(self, tag, attrs):
            if tag == "table":
                self.cur_table = []
            elif tag == "tr" and self.cur_table is not None:
                self.cur_row = []
            elif tag in ("td", "th") and self.cur_row is not None:
                d = dict(attrs)
                self.cur_cell = {
                    "text": "",
                    "colspan": int(d.get("colspan", 1)),
                    "rowspan": int(d.get("rowspan", 1)),
                }
                self.buf = []
            elif tag == "br" and self.cur_cell is not None:
                self.buf.append(" ")

        def handle_endtag(self, tag):
            if tag == "table" and self.cur_table is not None:
                self.tables.append(self.cur_table)
                self.cur_table = None
            elif tag == "tr" and self.cur_row is not None:
                if self.cur_table is not None:
                    self.cur_table.append(self.cur_row)
                self.cur_row = None
            elif tag in ("td", "th") and self.cur_cell is not None:
                self.cur_cell["text"] = "".join(self.buf).strip()
                self.cur_row.append(self.cur_cell)
                self.cur_cell = None

        def handle_data(self, data):
            if self.cur_cell is not None:
                self.buf.append(data)

    def table_to_pipe(rows: list[list[dict]]) -> str:
        # 展开 rowspan / colspan，补齐为整齐网格
        grid: list[list[str]] = []
        pending: dict[int, tuple[str, int]] = {}   # 列 -> (文字, 剩余行数)
        for row in rows:
            line: list[str] = []
            col = 0
            while col in pending:                  # 上方单元格延续下来
                text, rem = pending[col]
                line.append(text)
                if rem > 1:
                    pending[col] = (text, rem - 1)
                else:
                    del pending[col]
                col += 1
            for cell in row:
                for k in range(cell["colspan"]):
                    line.append(cell["text"] if k == 0 else "")
                    if cell["rowspan"] > 1 and k == 0:
                        pending[col] = (cell["text"], cell["rowspan"] - 1)
                    col += 1
            grid.append(line)

        width = max((len(r) for r in grid), default=0)
        grid = [r + [""] * (width - len(r)) for r in grid]

        def esc(s: str) -> str:
            return s.replace("|", "\\|").replace("\n", " ").strip()

        out = ["| " + " | ".join(esc(c) for c in grid[0]) + " |",
               "| " + " | ".join(["---"] * width) + " |"]
        out += ["| " + " | ".join(esc(c) for c in r) + " |" for r in grid[1:]]
        return "\n".join(out)

    parser = TableParser()
    parser.feed(md)

    for rows in parser.tables:
        if not rows:
            continue
        pipe = table_to_pipe(rows)
        # 用正则找回这段 HTML 原文并替换（按 <table> 出现顺序）
        m = re.search(r"<table.*?</table>", md, re.S)
        if m:
            md = md[:m.start()] + "\n\n" + pipe + "\n\n" + md[m.end():]
    return md


def merge_split_figures(paper_dir: Path, pdf_path: Path, dpi: int = 200) -> str:
    """合并被版面模型拆散的同一张图（多面板大图），从原 PDF 按坐标重渲染为整图。

    依据 structured_content.json 中 image/chart 块的 bbox：
    - 相邻（左右间距 <5% 页宽且垂直重叠 >50%，或上下间距 <3% 页高且水平重叠 >50%）
      且中间没有文字块隔开的 image 块视为同一张图
    返回改写后的 markdown（image 引用指向合并后的整图）。
    """
    import json

    import pymupdf

    import re

    struct_file = paper_dir / "structured_content.json"
    md_file = paper_dir / "markdown.md"
    if not struct_file.exists() or not md_file.exists() or not pdf_path.exists():
        return md_file.read_text(encoding="utf-8") if md_file.exists() else ""

    data = json.loads(struct_file.read_text(encoding="utf-8"))
    doc = pymupdf.open(pdf_path)

    def bbox_of(b):
        return b["bbox"]

    def overlap_1d(a0, a1, b0, b1):
        return max(0.0, min(a1, b1) - max(a0, b0))

    def adjacent(prev, cur):
        px0, py0, px1, py1 = bbox_of(prev)
        cx0, cy0, cx1, cy1 = bbox_of(cur)
        h_gap = max(cx0 - px1, px0 - cx1)          # 水平间距（负=重叠）
        v_gap = cy0 - py1                           # 垂直间距（cur 在 prev 下方）
        v_overlap = overlap_1d(py0, py1, cy0, cy1)
        h_overlap = overlap_1d(px0, px1, cx0, cx1)
        min_h = min(py1 - py0, cy1 - cy0)
        min_w = min(px1 - px0, cx1 - cx0)
        side_by_side = h_gap < 0.05 and v_overlap > 0.5 * min_h
        stacked = 0 <= v_gap < 0.03 and h_overlap > 0.5 * min_w
        return side_by_side or stacked

    # 每页：把连续的 image/chart 块（中间无其他块）聚成组
    groups = {}          # image_source -> (page_idx, group_id, union_bbox)
    for pg in data.get("pages", []):
        page_idx = pg.get("page_idx", 0)
        blocks = pg.get("blocks", [])
        cur_group, gid = [], 0
        runs = []
        for b in blocks:
            if b.get("type") in ("image", "chart") and b.get("image_source"):
                cur_group.append(b)
            else:
                if cur_group:
                    runs.append(cur_group)
                cur_group = []
        if cur_group:
            runs.append(cur_group)

        for run in runs:
            # 组内按相邻关系再细分
            sub = [[run[0]]]
            for b in run[1:]:
                if adjacent(sub[-1][-1], b):
                    sub[-1].append(b)
                else:
                    sub.append([b])
            for gid, grp in enumerate(sub):
                xs0 = min(b["bbox"][0] for b in grp)
                ys0 = min(b["bbox"][1] for b in grp)
                xs1 = max(b["bbox"][2] for b in grp)
                ys1 = max(b["bbox"][3] for b in grp)
                for b in grp:
                    groups[b["image_source"]] = (page_idx, gid,
                                                 (xs0, ys0, xs1, ys1), len(grp))

    # 重渲染多块组成的组（单块组保留 MinerU 提取的原图）
    img_dir = paper_dir / "images"
    img_dir.mkdir(exist_ok=True)
    composite_name = {}
    rendered = set()
    for src, (page_idx, gid, bbox, size) in groups.items():
        if size < 2 or (page_idx, gid) in rendered:
            continue
        rendered.add((page_idx, gid))
        if page_idx >= len(doc):
            continue
        page = doc[page_idx]
        rect = page.rect
        clip = pymupdf.Rect(bbox[0] * rect.width, bbox[1] * rect.height,
                            bbox[2] * rect.width, bbox[3] * rect.height)
        if clip.is_empty or clip.width < 10 or clip.height < 10:
            continue
        pix = page.get_pixmap(clip=clip, dpi=dpi)
        if pix.n > 4:
            pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
        name = f"fig_p{page_idx + 1}_{gid}.png"
        pix.save(str(img_dir / name))
        pix = None
        composite_name[(page_idx, gid)] = name
    doc.close()

    # 改写 markdown：同一组的连续 image 行合并为一行
    md = md_file.read_text(encoding="utf-8")
    lines = md.splitlines()
    out, emitted, i = [], set(), 0
    img_re = re.compile(r"^!\[\]\(images/[^)]+\)$")
    while i < len(lines):
        m = img_re.match(lines[i].strip())
        if not m:
            out.append(lines[i])
            i += 1
            continue
        src = lines[i].strip()[4:-1]                    # ![](images/xxx.jpg) -> images/xxx.jpg
        key = groups.get(src)
        if key is None:
            out.append(lines[i])
            i += 1
            continue
        page_idx, gid, _, _ = key
        comp = composite_name.get((page_idx, gid))
        if comp is None:                                # 单块组：保留原引用
            out.append(lines[i])
            i += 1
            continue
        if (page_idx, gid) in emitted:                  # 组内后续图片行：跳过
            i += 1
            continue
        emitted.add((page_idx, gid))
        out.append(f"![](images/{comp})")
        i += 1
    return "\n".join(out) + "\n"


def md_to_docx(md_file: Path, pdf_path: Path | None = None) -> Path | None:
    """markdown -> docx（pandoc）：公式转 Word 原生方程，表格转 Word 原生表格"""
    try:
        import pypandoc
    except ImportError:
        print("    [跳过 docx] 未安装 pypandoc-binary，请运行: pip install pypandoc-binary")
        return None

    # 1. 合并被拆散的大图（多面板图按坐标重渲染为整图）
    if pdf_path is not None:
        md = merge_split_figures(md_file.parent, pdf_path)
    else:
        md = md_file.read_text(encoding="utf-8")
    # 2. HTML 表格 -> 管道符表格（pandoc 否则丢弃表格结构）
    md = html_tables_to_pipe(md)
    # 3. LaTeX 宏兼容修复
    md = _fix_tex_macros(md)
    docx_file = md_file.with_suffix(".docx")
    pypandoc.convert_text(
        md, "docx", format="markdown",
        outputfile=str(docx_file),
        extra_args=["--resource-path", str(md_file.parent)],
    )
    return docx_file


def run_mineru(pdf: Path, out_dir: Path, tier: str, pages: str | None) -> Path:
    """调用 mineru-kit parse（zip 格式），返回 zip 路径"""
    cmd = [
        str(MINERU_KIT), "parse", str(pdf),
        "-o", str(out_dir), "-f", "zip",
        "--tier", tier,
    ]
    if pages:
        cmd += ["-p", pages]

    t0 = time.time()
    proc = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    elapsed = time.time() - t0

    zip_path = out_dir / f"{pdf.stem}.zip"
    if not zip_path.exists():
        print(f"[失败] {pdf.name} ({elapsed:.0f}s)")
        # 打印关键错误行
        for line in (proc.stdout + proc.stderr).splitlines():
            if any(k in line for k in ("Error", "error", "Traceback", "Exception")):
                print("   ", line.strip()[:200])
        return None

    print(f"[完成] {pdf.name} ({elapsed:.0f}s)")
    return zip_path


def unpack(zip_path: Path, dest: Path, keep_zip: bool) -> dict:
    """解压 zip 到目标目录，返回统计信息"""
    with zipfile.ZipFile(zip_path) as z:
        names = z.namelist()
        z.extractall(dest)

    if not keep_zip:
        zip_path.unlink()

    md_file = dest / "markdown.md"
    stats = {"images": sum(1 for n in names if n.startswith("images/")), "size_mb": 0}
    if md_file.exists():
        stats["size_mb"] = round(md_file.stat().st_size / 1024, 1)

    # 统计块类型（公式/表格/图片/图表）
    mid = dest / "middle_json.json"
    if mid.exists():
        data = json.loads(mid.read_text(encoding="utf-8"))
        counts = {}
        for pg in data.get("pages", []):
            for b in pg.get("blocks", []):
                counts[b.get("type", "?")] = counts.get(b.get("type", "?"), 0) + 1
        stats["blocks"] = counts
    return stats


def main():
    ap = argparse.ArgumentParser(description="PDF 解析小工具（基于 MinerU）")
    ap.add_argument("input", nargs="?", default=".", help="PDF 文件或包含 PDF 的目录")
    ap.add_argument("-o", "--output", default=str(DEFAULT_OUTPUT), help="输出目录")
    ap.add_argument("--tier", default="basic",
                    choices=["flash", "basic", "standard", "advanced"],
                    help="解析档位（默认 basic，无独显的 CPU 环境推荐）")
    ap.add_argument("-p", "--pages", default=None, help="页码，如 '1-5,8'（默认全部）")
    ap.add_argument("--docx", action="store_true", help="同时导出 Word (.docx)")
    ap.add_argument("--exclude", action="append", default=[],
                    help="排除文件名包含该关键词的 PDF（可多次指定）")
    ap.add_argument("--keep-zip", action="store_true", help="保留原始 zip 包")
    ap.add_argument("--list", default=None, help="从文本文件读取 PDF 路径列表（每行一个）")
    args = ap.parse_args()

    if not MINERU_KIT.exists():
        print(f"[错误] 未找到 mineru-kit: {MINERU_KIT}")
        print("请先运行: .venv\\Scripts\\pip.exe install \"mineru>=4.0,<5\"")
        sys.exit(1)

    pdfs = find_pdfs(Path(args.input), args.exclude,
                     Path(args.list) if args.list else None)
    if not pdfs:
        print("[错误] 没有找到 PDF 文件（或全被排除）")
        sys.exit(1)

    out_root = Path(args.output)
    raw_dir = out_root / "_raw"          # mineru-kit 的 zip 临时落点
    raw_dir.mkdir(parents=True, exist_ok=True)

    print(f"待解析 {len(pdfs)} 个 PDF，档位 {args.tier}"
          f"{'，导出 docx' if args.docx else ''}，输出到 {out_root}\n")
    ok, fail = 0, 0
    for pdf in pdfs:
        zip_path = run_mineru(pdf, raw_dir, args.tier, args.pages)
        if zip_path is None:
            fail += 1
            continue
        dest = out_root / pdf.stem
        if dest.exists():
            shutil.rmtree(dest)
        stats = unpack(zip_path, dest, args.keep_zip)
        blocks = stats.get("blocks", {})
        summary = ", ".join(f"{k}×{v}" for k, v in sorted(blocks.items()))
        print(f"    -> {dest}")
        print(f"       图片 {stats['images']} 张 | markdown {stats['size_mb']} KB | {summary}")

        if args.docx:
            docx_file = md_to_docx(dest / "markdown.md")
            if docx_file:
                print(f"       docx {round(docx_file.stat().st_size / 1024)} KB -> {docx_file.name}")
        ok += 1

    print(f"\n全部完成：成功 {ok}，失败 {fail}")


if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-
"""验证修复效果：docx 表格数 vs middle_json 表格块数；图片合并情况"""
import json
import re
import zipfile
from pathlib import Path

ROOT = Path(r"e:/pdf-tool/output/dft")

print(f"{'paper':38s} {'tbl_mid':>7s} {'tbl_docx':>8s} {'img_before':>10s} {'img_after':>9s} {'figs':>5s}")
total_mid = total_docx = 0
for d in sorted(p for p in ROOT.glob("*/") if p.is_dir()):
    md = d / "markdown.md"
    orig = d / "markdown_original.md"
    docx = d / "markdown.docx"
    mid = d / "middle_json.json"
    if not md.exists() or not docx.exists():
        continue

    xml = zipfile.ZipFile(docx).read("word/document.xml").decode()
    tbl_docx = xml.count("<w:tbl>")
    tbl_mid = 0
    if mid.exists():
        data = json.loads(mid.read_text(encoding="utf-8"))
        tbl_mid = sum(1 for pg in data.get("pages", []) for b in pg.get("blocks", [])
                      if b.get("type") == "table")
    total_mid += tbl_mid
    total_docx += tbl_docx

    img_after = len(re.findall(r"^!\[\]\(", md.read_text(encoding="utf-8"), re.M))
    img_before = len(re.findall(r"^!\[\]\(", orig.read_text(encoding="utf-8"), re.M)) if orig.exists() else -1
    figs = len(list((d / "images").glob("fig_p*.png")))
    flag = "OK" if tbl_docx >= tbl_mid else "LOST!"
    print(f"{d.name[:38]:38s} {tbl_mid:7d} {tbl_docx:8d} {img_before:10d} {img_after:9d} {figs:5d}  {flag}")

print(f"\n合计: middle_json 表格块 {total_mid} -> docx 表格 {total_docx}")

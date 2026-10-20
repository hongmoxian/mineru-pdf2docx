# -*- coding: utf-8 -*-
"""抽检 docx 内部结构：Word 原生方程 / 表格 / 图片数量"""
import zipfile
from pathlib import Path

ROOT = Path(r"e:/pdf-tool/output/dft")
CHECK = [
    "2022The Open Catalyst 2022 (OC22) Dataset and",
    "2024CHGNet as a pretrained universal neural network potential for charge-informed atomistic modelling",
    "2024Highly accurate carbohydrate-binding site prediction with DeepGlycanSite",
    "becker-et-al-2023-multiscale-modeling-of-aqueous-electric-double-layers",
]

total_omath = total_tbl = total_img = 0
for name in CHECK:
    p = ROOT / name / "markdown.docx"
    z = zipfile.ZipFile(p)
    xml = z.read("word/document.xml").decode("utf-8")
    omath = xml.count("<m:oMath")
    tbl = xml.count("<w:tbl>")
    img = len([n for n in z.namelist() if "media" in n])
    total_omath += omath
    total_tbl += tbl
    total_img += img
    print(f"{name[:42]:42s}  equations={omath:4d}  tables={tbl:3d}  images={img:3d}  size={p.stat().st_size//1024}KB")

print(f"\n抽检合计: equations={total_omath}, tables={total_tbl}, images={total_img}")

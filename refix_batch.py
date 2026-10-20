# -*- coding: utf-8 -*-
"""
refix_batch.py — 批量修复已解析输出的两个问题：
  1. HTML 表格被 pandoc 丢弃 -> 转为管道符表格
  2. 多面板大图被拆成多个小图 -> 按 bbox 合并重渲染为整图

用法: python refix_batch.py [输出根目录] [PDF 源目录]
"""
import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from parse_pdf import html_tables_to_pipe, md_to_docx, merge_split_figures


def main():
    root = Path(sys.argv[1] if len(sys.argv) > 1 else r"e:/pdf-tool/output/dft")
    src_dir = Path(sys.argv[2] if len(sys.argv) > 2 else r"e:/download")

    for paper_dir in sorted(p for p in root.glob("*/") if p.is_dir()):
        md_file = paper_dir / "markdown.md"
        if not md_file.exists():
            continue
        pdf_path = src_dir / f"{paper_dir.name}.pdf"

        # 备份原始 markdown（仅第一次）
        backup = paper_dir / "markdown_original.md"
        if not backup.exists():
            shutil.copy2(md_file, backup)

        # 修复 1：合并拆散的大图；修复 2：HTML 表格 -> 管道符表格
        md = merge_split_figures(paper_dir, pdf_path)
        md = html_tables_to_pipe(md)
        md_file.write_text(md, encoding="utf-8")

        # 重新生成 docx（pdf_path=None：不再重复合并，直接用修复后的 markdown）
        docx = md_to_docx(md_file, None)
        size_kb = docx.stat().st_size // 1024 if docx else -1
        print(f"[修复] {paper_dir.name[:50]:50s} docx={size_kb}KB")


if __name__ == "__main__":
    main()

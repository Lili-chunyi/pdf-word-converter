# -*- coding: utf-8 -*-
"""生成测试样例文件：
- sample_text.pdf    文本型 PDF（中文段落、彩色文字、表格、图片、超链接，3 页）
- sample_scanned.pdf 扫描件模拟（由 sample_text.pdf 整页渲染成图片重组，2 页）
- sample.docx        Word 样例（标题、中文段落、表格）
输出到 tests/out/ 目录。
"""
import os
import sys

import fitz
from docx import Document
from docx.shared import Pt

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "out")
os.makedirs(OUT, exist_ok=True)


def make_text_pdf(path, pages=3):
    doc = fitz.open()
    for i in range(pages):
        page = doc.new_page(width=595, height=842)  # A4
        # 标题
        page.insert_text((72, 90), f"项目验收报告（第 {i + 1} 页）",
                         fontname="china-ss", fontsize=20, color=(0.1, 0.1, 0.6))
        # 段落
        text = ("本报告用于验证文档格式双向转换工具的核心能力。" * 3)
        rect = fitz.Rect(72, 120, 523, 220)
        page.insert_textbox(rect, text, fontname="china-ss",
                            fontsize=11, color=(0.15, 0.15, 0.15), align=0)
        # 彩色强调文字
        page.insert_text((72, 250), "重点事项：验收标准必须满足全部硬性指标。",
                         fontname="china-ss", fontsize=12, color=(0.8, 0.1, 0.1))
        # 表格（3x4，含中文）
        top = 290
        widths = [120, 110, 100, 80]
        headers = ["验收项目", "指标要求", "实测结果", "结论"]
        rows = [
            ["文本还原率", ">= 99%", "99.7%", "通过"],
            ["图片完整率", "100%", "100%", "通过"],
            ["页序一致性", "100%", "100%", "通过"],
            ["扫描件处理", "不丢页", "不丢页", "通过"],
        ]
        row_h = 26
        x = 72
        xs = [x]
        for w in widths:
            xs.append(xs[-1] + w)
        for r, row in enumerate([headers] + rows):
            y = top + r * row_h
            for c in range(len(widths)):
                cell = fitz.Rect(xs[c], y, xs[c + 1], y + row_h)
                page.draw_rect(cell, color=(0.4, 0.4, 0.4), width=0.7)
                label = row[c]
                page.insert_text((xs[c] + 6, y + 17), label,
                                 fontname="china-ss", fontsize=10,
                                 color=(0, 0, 0.5) if r == 0 else (0.1, 0.1, 0.1))
        # 图片（渐变色块模拟插图）
        img_rect = fitz.Rect(72, 460, 292, 590)
        pix = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, 220, 130), False)
        for yy in range(130):
            for xx in range(220):
                pix.set_pixel(xx, yy, (xx / 220, yy / 130, 0.5))
        page.insert_image(img_rect, pixmap=pix)
        page.insert_text((72, 620), "插图 1：趋势示意图", fontname="china-ss",
                         fontsize=9, color=(0.35, 0.35, 0.35))
        # 超链接
        link_rect = fitz.Rect(72, 660, 260, 676)
        page.insert_text((72, 672), "开源转换引擎：LibreOffice 官网",
                         fontname="china-ss", fontsize=10, color=(0.05, 0.3, 0.85))
        page.insert_link({"kind": fitz.LINK_URI, "from": link_rect,
                          "uri": "https://www.libreoffice.org/"})
        # 页脚页码
        page.insert_text((280, 810), f"- {i + 1} -", fontname="china-ss",
                         fontsize=10, color=(0.4, 0.4, 0.4))
    doc.save(path)
    doc.close()
    print("已生成:", path)


def make_scanned_pdf(src_pdf, path):
    """把文本 PDF 的每一页渲染成图片再重组，模拟扫描件。"""
    src = fitz.open(src_pdf)
    out = fitz.open()
    for page in src:
        pix = page.get_pixmap(dpi=150, alpha=False)
        new = out.new_page(width=page.rect.width, height=page.rect.height)
        new.insert_image(new.rect, pixmap=pix)
    out.save(path)
    out.close()
    src.close()
    print("已生成:", path)


def make_docx(path):
    d = Document()
    d.add_heading("产品验收总结", level=1)
    d.add_paragraph("本文档用于测试 Word 转 PDF 功能，包含中文、英文、表格与多级标题。")
    d.add_heading("一、验收结论", level=2)
    d.add_paragraph("全部核心指标通过验收：转换稳定、排版保持良好、无丢页现象。")
    t = d.add_table(rows=3, cols=3)
    t.style = "Table Grid"
    data = [["指标", "目标", "结果"], ["排版一致性", "良好", "良好"], ["转换耗时", "< 60s", "23s"]]
    for r, row in enumerate(data):
        for c, val in enumerate(row):
            t.cell(r, c).text = val
    d.add_heading("二、后续计划", level=2)
    d.add_paragraph("持续优化复杂表格与多栏版式的还原精度。")
    for p in d.paragraphs:
        for run in p.runs:
            run.font.size = Pt(11)
    d.save(path)
    print("已生成:", path)


if __name__ == "__main__":
    make_text_pdf(os.path.join(OUT, "sample_text.pdf"))
    make_scanned_pdf(os.path.join(OUT, "sample_text.pdf"),
                     os.path.join(OUT, "sample_scanned.pdf"))
    make_docx(os.path.join(OUT, "sample.docx"))
    print("样例生成完毕 ->", OUT)
    sys.exit(0)

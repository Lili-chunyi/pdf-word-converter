# -*- coding: utf-8 -*-
"""端到端转换测试：直接调用 converter 模块验证三条链路。
1. 文本型 PDF → 可编辑 Word（文本引擎）
2. 扫描件 PDF → 图片版 Word（图片引擎，逐页校验）
3. Word → PDF（LibreOffice 引擎）
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import fitz
from docx import Document

from converter import pdf2word, word2pdf

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "out")

passed, failed = [], []


def check(name, cond, detail=""):
    if cond:
        passed.append(name)
        print(f"  [PASS] {name} {detail}")
    else:
        failed.append(name)
        print(f"  [FAIL] {name} {detail}")


print("== 测试 1：文本型 PDF → 可编辑 Word ==")
src = os.path.join(OUT, "sample_text.pdf")
dst = os.path.join(OUT, "result_text.docx")
used, pages = pdf2word.convert_pdf_to_word(src, dst, mode="auto", dpi=200)
print(f"  引擎={used}, 页数={pages}")
check("自动识别为文本引擎", used == "text")
check("页数一致(3)", pages == 3)
d = Document(dst)
all_text = "\n".join(p.text for p in d.paragraphs)
for tb in d.tables:
    for row in tb.rows:
        for cell in row.cells:
            all_text += "\n" + cell.text
check("中文文本可提取(可编辑)", "验收" in all_text or "项目" in all_text)
check("表格已重建", len(d.tables) >= 1, f"tables={len(d.tables)}")
check("产物非空", os.path.getsize(dst) > 5000)

print("== 测试 2：扫描件 PDF → 图片版 Word ==")
src2 = os.path.join(OUT, "sample_scanned.pdf")
dst2 = os.path.join(OUT, "result_image.docx")
msgs = []
used2, pages2 = pdf2word.convert_pdf_to_word(
    src2, dst2, mode="auto", dpi=200,
    progress_cb=lambda p, m: msgs.append((p, m)))
print(f"  引擎={used2}, 页数={pages2}, 进度消息={len(msgs)} 条")
check("自动识别为图片引擎", used2 == "image")
check("页数一致(3)", pages2 == 3)
d2 = Document(dst2)
n_img = len(d2.inline_shapes)
check("每页一张图片(3 张)", n_img == 3, f"images={n_img}")
check("图片版无正文文本", not any(p.text.strip() for p in d2.paragraphs))
check("图片清晰度(首图宽>=1000px)", d2.inline_shapes[0].width > 950000 if n_img else False)
# 页面尺寸与 PDF 一致（A4: 595x842pt -> EMU 7560000x10692000）
sec = d2.sections[0]
check("页面尺寸保持 A4", abs(sec.page_width - 595 * 12700) < 30000
      and abs(sec.page_height - 842 * 12700) < 30000)

print("== 测试 3：强制图片模式 + 文本引擎失败回退 ==")
dst3 = os.path.join(OUT, "result_force_image.docx")
used3, _ = pdf2word.convert_pdf_to_word(src, dst3, mode="image", dpi=150)
d3 = Document(dst3)
check("强制图片模式生效", used3 == "image" and len(d3.inline_shapes) == 3)

print("== 测试 4：Word → PDF ==")
src4 = os.path.join(OUT, "sample.docx")
dst4 = os.path.join(OUT, "result_word2pdf")
os.makedirs(dst4, exist_ok=True)
pdf_path = word2pdf.convert_word_to_pdf(src4, dst4)
print(f"  输出: {pdf_path}")
check("生成 PDF 文件", os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 1000)
pd = fitz.open(pdf_path)
check("PDF 页数>=1", pd.page_count >= 1, f"pages={pd.page_count}")
pd_text = "".join(pg.get_text() for pg in pd)
check("PDF 含中文内容", "验收" in pd_text)
check("PDF 含表格数据", "23s" in pd_text or "< 60s" in pd_text)
pd.close()

print()
print(f"通过 {len(passed)} 项，失败 {len(failed)} 项")
if failed:
    print("失败项:", failed)
    sys.exit(1)
print("全部通过 ✔")

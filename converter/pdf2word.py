# -*- coding: utf-8 -*-
"""PDF → Word 转换模块

双引擎设计：
- 文本引擎（pdf2docx）：面向文本型 PDF，输出可编辑 Word，尽量保留
  文字、段落、字体、字号、颜色、表格、图片、超链接等。
- 图片引擎（PyMuPDF + python-docx）：面向扫描件/图片型 PDF，
  每页渲染为一张高清图片，按原顺序、原方向、原比例逐页嵌入 Word。

自动识别：统计每页可提取文本量，判定文档类型；文本引擎失败时
自动回退图片引擎，保证不丢页、不出空白文档。
"""
import io
import logging
import os

import fitz  # PyMuPDF
from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.shared import Emu, Pt

# 平均每页可提取文本字符数低于该阈值 → 判定为扫描件/图片型
SCANNED_AVG_CHARS = 60
# 单页文本字符数低于该值视为“无文本页”
EMPTY_PAGE_CHARS = 20
# 渲染图片最长边像素上限（防止超大页面耗尽内存）
MAX_RENDER_PIXELS = 8000

PT_TO_EMU = 12700  # 1 pt = 12700 EMU


class ConversionError(Exception):
    """转换失败，message 面向最终用户展示。"""


def analyze_pdf(pdf_path):
    """打开并分析 PDF，返回 (doc, [每页统计])。统计含文本字符数与图片数。"""
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        raise ConversionError(f"无法打开 PDF 文件，文件可能已损坏：{e}")
    if doc.needs_pass:
        doc.close()
        raise ConversionError("该 PDF 已加密，请先解除密码保护后再转换。")
    if doc.page_count == 0:
        doc.close()
        raise ConversionError("该 PDF 没有任何页面，无法转换。")
    pages = []
    for page in doc:
        text = page.get_text("text").strip()
        pages.append({
            "chars": len(text),
            "images": len(page.get_images(full=True)),
        })
    return doc, pages


def detect_scanned(pages):
    """判定是否为扫描件/图片型 PDF。"""
    n = len(pages)
    avg = sum(p["chars"] for p in pages) / n
    if avg < SCANNED_AVG_CHARS:
        return True
    # 混合型：绝大多数页面无文本（图片占版面主体）时也按扫描件处理，
    # 保证扫描页不会被转换成空白。
    empty_ratio = sum(1 for p in pages if p["chars"] < EMPTY_PAGE_CHARS) / n
    return empty_ratio > 0.8


def convert_pdf_to_word(pdf_path, docx_path, mode="auto", dpi=200, progress_cb=None):
    """转换入口。

    mode: "auto" 自动识别 | "text" 强制可编辑文本 | "image" 强制图片版
    dpi:  图片版渲染清晰度
    返回 (实际使用的引擎, 页数)
    """
    def report(p, m):
        if progress_cb:
            progress_cb(p, m)

    doc, pages = analyze_pdf(pdf_path)
    page_count = doc.page_count
    doc.close()

    if mode == "auto":
        chosen = "image" if detect_scanned(pages) else "text"
        label = "扫描件/图片型" if chosen == "image" else "文本型"
        report(8, f"已识别为{label} PDF，共 {page_count} 页，开始转换…")
    else:
        chosen = mode
        report(8, f"共 {page_count} 页，开始转换…")

    if chosen == "text":
        try:
            report(15, "正在解析版式并重建 Word 内容（复杂页面耗时较长）…")
            _convert_text_mode(pdf_path, docx_path)
            _verify_text_docx(docx_path)
        except Exception as e:
            # 文本引擎失败 → 回退图片引擎，保证不丢页
            report(35, f"文本引擎处理失败（{e}），已自动改用图片版引擎…")
            chosen = "image"
            _convert_image_mode(pdf_path, docx_path, dpi, report)
    else:
        _convert_image_mode(pdf_path, docx_path, dpi, report)

    if chosen == "image":
        _verify_image_docx(docx_path, page_count)

    return chosen, page_count


# ---------------------------------------------------------------------------
# 引擎一：文本型 PDF → 可编辑 Word（pdf2docx）
# ---------------------------------------------------------------------------
def _convert_text_mode(pdf_path, docx_path):
    logging.getLogger("pdf2docx").setLevel(logging.ERROR)
    from pdf2docx import Converter
    cv = None
    try:
        cv = Converter(pdf_path)
        cv.convert(docx_path, start=0, end=None)
    finally:
        if cv is not None:
            try:
                cv.close()
            except Exception:
                pass
    if not os.path.exists(docx_path) or os.path.getsize(docx_path) == 0:
        raise ConversionError("文本引擎未生成有效文档。")


def _verify_text_docx(docx_path):
    """校验文本引擎产物非空：至少含有文字、图片或表格，否则视为失败。"""
    try:
        d = Document(docx_path)
    except Exception as e:
        raise ConversionError(f"生成的 Word 无法读取：{e}")
    has_text = any(p.text.strip() for p in d.paragraphs)
    has_img = len(d.inline_shapes) > 0
    has_table = len(d.tables) > 0
    if not (has_text or has_img or has_table):
        raise ConversionError("文本引擎未提取到任何内容。")


# ---------------------------------------------------------------------------
# 引擎二：扫描件/图片型 PDF → 图片版 Word（逐页渲染嵌入）
# ---------------------------------------------------------------------------
def _pix_to_bytes(pix, quality=92):
    """将渲染位图编码为 JPEG（失败时回退 PNG）。"""
    try:
        return pix.tobytes("jpeg", jpg_quality=quality)
    except Exception:
        return pix.tobytes("png")


def _render_zoom(rect, dpi):
    """计算渲染倍率，超大幅面自动降低以保证内存可控。"""
    zoom = dpi / 72.0
    scale = max(rect.width, rect.height) * zoom
    if scale > MAX_RENDER_PIXELS:
        zoom = MAX_RENDER_PIXELS / max(rect.width, rect.height)
    return zoom


def _convert_image_mode(pdf_path, docx_path, dpi=200, progress_cb=None):
    doc = fitz.open(pdf_path)
    out = Document()
    n = doc.page_count
    try:
        for i, page in enumerate(doc):
            rect = page.rect  # PyMuPDF 的 rect 已应用页面旋转，方向与肉眼所见一致
            zoom = _render_zoom(rect, dpi)
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
            img_bytes = _pix_to_bytes(pix)
            pix = None

            if i > 0:
                out.add_section(WD_SECTION.NEW_PAGE)
            sec = out.sections[-1]
            sec.page_width = Emu(int(rect.width * PT_TO_EMU))
            sec.page_height = Emu(int(rect.height * PT_TO_EMU))
            if rect.width > rect.height:
                sec.orientation = WD_ORIENT.LANDSCAPE
            else:
                sec.orientation = WD_ORIENT.PORTRAIT
            sec.left_margin = Emu(0)
            sec.right_margin = Emu(0)
            sec.top_margin = Emu(0)
            sec.bottom_margin = Emu(0)
            sec.header_distance = Emu(0)
            sec.footer_distance = Emu(0)

            # 图片四周各留 0.5pt 白边，避免尺寸取整误差把图片挤到下一页产生空白页
            img_w = Emu(int(rect.width * PT_TO_EMU) - PT_TO_EMU)
            img_h = Emu(int(rect.height * PT_TO_EMU) - PT_TO_EMU)
            p = out.add_paragraph()
            pf = p.paragraph_format
            pf.space_before = Pt(0)
            pf.space_after = Pt(0)
            pf.line_spacing = 1.0
            run = p.add_run()
            run.add_picture(io.BytesIO(img_bytes), width=img_w, height=img_h)

            if progress_cb:
                progress_cb(10 + int(85 * (i + 1) / n),
                            f"正在嵌入第 {i + 1}/{n} 页…")
        out.save(docx_path)
    finally:
        doc.close()


def _verify_image_docx(docx_path, page_count):
    """硬性校验：图片版 Word 中图片数量必须与 PDF 页数一致（不丢页、无空白）。"""
    d = Document(docx_path)
    imgs = len(d.inline_shapes)
    if imgs != page_count:
        raise ConversionError(
            f"结果校验失败：PDF 共 {page_count} 页，Word 中仅含 {imgs} 张页面图片。")

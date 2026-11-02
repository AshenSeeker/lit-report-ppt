# -*- coding: utf-8 -*-
"""lit-report-ppt 公用构件。

所有坐标单位 is inch，所有颜色是 'RRGGBB'。
核心是一件事：把 content.json 里的 `Cr^{3+}` / `Cs_{2}NaInCl_{6}` 标记
翻译成真正的上下标 run —— 不能靠 Unicode 上标字符，那玩意儿换个字体就散。
"""
from __future__ import annotations

import re
from typing import Iterable

from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

# ---------------------------------------------------------------- 设计常量
# 全部来自 samples 实测，改之前看 references/style-tokens.md
CANVAS_W = 13.333
CANVAS_H = 7.5

FONT_LATIN = "Times New Roman"
FONT_CJK = "楷体"
FONT_NUM = "等线"
FONT_COVER_EN = "Verdana"

H0, H1, H2 = 66, 22, 18
B1, B2 = 16, 14
C1, C2 = 11, 10

BLUE = "2A7DC0"
INK = "132E4A"
MATH = "4472C4"
RED = "C00000"
ORANGE = "ED7D31"
BODY = "1A1A1A"
MUTED = "6B6B6B"
WHITE = "FFFFFF"
RULE = "D6DEE6"

LAYOUT_BLANK = 6
LAYOUT_TITLE_CONTENT = 1

# 安全边距
MARGIN_L = 0.40
MARGIN_R = 0.40
MARGIN_T = 0.30
MARGIN_B = 0.35

# ------------------------------------------------------------------ 主图页
# 样本里两种排法是**并列**的，不是二选一的风格倾向：
#   A 底栏短图注 —— P9/P12/P17，图满宽、图注一两行压页面底边。
#   B 右栏长图注 —— P8/P16/P21/P26，图缩到 8.4–9.1 宽，右侧 2.1–2.3 宽竖窄栏。
# 早期我把两种合并成「全右栏」，导致所有图都被压窄 —— 这是排版走形的真正原因。
# 现在按图注长度自动选：短的走 A（图最大），长的走 B。
FIG_A_IMG = (0.61, 0.15, 11.71, 6.20)    # 图 x y w h
FIG_A_CAP = (0.72, 6.44, 11.80, 0.92)    # 图注
FIG_B_IMG = (1.35, 0.16, 8.45, 7.16)
FIG_B_CAP = (9.95, 0.14, 2.75, 7.14)

# 图注容量：不按字数，按**排版宽度**算行数。
# 曾按字数比（样本 267 字符走底栏 → 中文 300 字也走底栏）是错的：
# 一个汉字占 1em，一个西文字母只占约 0.5em，267 字符英文 ≈ 130 个汉字。
# 所以用 est_text_width 估总宽度，再除每行可用宽度得到行数。
# 断行有损耗（西文不拆词、标点避头尾），实测乘 1.20 才贴近。
CAP_WRAP_LOSS = 1.20


# ---------------------------------------------------------------- 色
def rgb(h: str) -> RGBColor:
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


# ---------------------------------------------------------------- 上下标解析
_TOKEN = re.compile(r"(\^\{[^}]*\}|_\{[^}]*\})")


def parse_markup(text: str) -> list[tuple[str, str | None]]:
    """把 'Cs_{2}NaInCl_{6}：Cr^{3+}' 拆成 [(片断, 'super'|'sub'|None), ...]。

    只认 `^{...}` 和 `_{...}` 两种标记；没闭合的花括号当普通文本。
    """
    out: list[tuple[str, str | None]] = []
    pos = 0
    for m in _TOKEN.finditer(text):
        if m.start() > pos:
            out.append((text[pos:m.start()], None))
        tok = m.group(0)
        kind = "super" if tok.startswith("^") else "sub"
        out.append((tok[2:-1], kind))
        pos = m.end()
    if pos < len(text):
        out.append((text[pos:], None))
    return out or [(text, None)]


def plain(text: str) -> str:
    """去掉上下标标记，用于字数统计。"""
    return "".join(seg for seg, _ in parse_markup(text))


# ---------------------------------------------------------------- run 与段落
_CJK = re.compile(r"[\u3000-\u303f\u3400-\u4dbf\u4e00-\u9fff\uff00-\uffef]")


def _split_scripts(s: str) -> Iterable[tuple[str, bool]]:
    """按中英文切段，返回 (文本, 是否中文)。中文连跑，英文含空格连跑。"""
    buf, cur = [], None
    for ch in s:
        if ch == "\n":
            buf.append(ch)
            continue
        is_cjk = bool(_CJK.match(ch))
        if cur is None or is_cjk == cur:
            buf.append(ch)
            cur = is_cjk
        else:
            yield "".join(buf), cur
            buf, cur = [ch], is_cjk
    if buf:
        yield "".join(buf), bool(cur)


def add_runs(paragraph, text: str, size: int, *, bold=False,
             color: str = BODY, font_latin: str = FONT_LATIN,
             font_cjk: str = FONT_CJK) -> None:
    """把带上下标标记的文本写进段落，自动中英分 run。"""
    for seg, kind in parse_markup(text):
        if kind is not None:
            r = paragraph.add_run()
            r.text = seg
            f = r.font
            f.size = Pt(round(size * 0.72, 1))
            f.bold = bold
            f.name = font_latin
            f.color.rgb = rgb(color)
            # 上标 30000 / 下标 -25000，与原稿一致
            r._r.get_or_add_rPr().set("baseline", "30000" if kind == "super" else "-25000")
            continue
        for chunk, is_cjk in _split_scripts(seg):
            r = paragraph.add_run()
            r.text = chunk
            f = r.font
            f.size = Pt(size)
            f.bold = bold
            f.name = font_cjk if is_cjk else font_latin
            f.color.rgb = rgb(color)
            if is_cjk:
                # 东亚字体单独指定，否则 PowerPoint 会拿 latin 字体渲染汉字
                rPr = r._r.get_or_add_rPr()
                ea = rPr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}ea")
                if ea is None:
                    ea = rPr.makeelement(
                        "{http://schemas.openxmlformats.org/drawingml/2006/main}ea", {})
                    rPr.append(ea)
                ea.set("typeface", font_cjk)


def write_box(tf, lines, size: int, *, bold=False, color: str = BODY,
              align=PP_ALIGN.LEFT, space_after: int = 4,
              line_spacing: float = 1.25, first_line_only=False,
              font_latin: str = FONT_LATIN, font_cjk: str = FONT_CJK) -> None:
    """lines 可以是 str 或 list[str]。

    字符串里的 `\n` 会拆成独立段落 —— 直接把换行塞进 a:t 里，LibreOffice 和
    PowerPoint 各自处理方式不同，实测会变成「首行左、次行右」的错位。
    """
    if isinstance(lines, str):
        lines = lines.split("\n")
    else:
        expanded = []
        for ln in lines:
            expanded.extend(str(ln).split("\n"))
        lines = expanded
    tf.word_wrap = True
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        p.space_after = Pt(space_after)
        if first_line_only and i > 0:
            break
        add_runs(p, ln, size, bold=bold, color=color,
                 font_latin=font_latin, font_cjk=font_cjk)


# ---------------------------------------------------------------- 形状
def textbox(slide, x, y, w, h, *, anchor=MSO_ANCHOR.TOP,
            margins=(0.04, 0.02, 0.04, 0.02)):
    """返回 (shape, text_frame)。margins = (l, t, r, b) inch。"""
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left, tf.margin_top, tf.margin_right, tf.margin_bottom = (
        Inches(m) for m in margins)
    return box, tf


def rect(slide, x, y, w, h, *, fill: str | None = None,
         line: str | None = None, line_w: float = 1.0,
         dash: str | None = None, radius: float | None = None):
    """纯色块/描边框。dash 用 'dash' / 'sysDot' 等 MSO_LINE_DASH_STYLE 名。"""
    from pptx.enum.shapes import MSO_SHAPE
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    sh = slide.shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))
    if radius:
        sh.adjustments[0] = radius
    if fill:
        sh.fill.solid()
        sh.fill.fore_color.rgb = rgb(fill)
    else:
        sh.fill.background()
    if line:
        sh.line.color.rgb = rgb(line)
        sh.line.width = Pt(line_w)
    else:
        sh.line.fill.background()
    if dash:
        from pptx.enum.dml import MSO_LINE_DASH_STYLE
        sh.line.dash_style = getattr(MSO_LINE_DASH_STYLE, dash.upper(), None)
    sh.shadow.inherit = False
    return sh


def hline(slide, x, y, w, color: str = BLUE, weight: float = 1.0):
    from pptx.enum.shapes import MSO_CONNECTOR
    ln = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, Inches(x), Inches(y), Inches(x + w), Inches(y))
    ln.line.color.rgb = rgb(color)
    ln.line.width = Pt(weight)
    return ln


def set_bg(slide, color: str | None):
    """None 表示沿用版式背景（母版水波纹）。"""
    if color is None:
        return
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = rgb(color)


# ---------------------------------------------------------------- 表格
def cell_border(cell, color: str = RULE, weight: float = 0.75,
                edges=("L", "R", "T", "B")) -> None:
    """给单元格画边框。python-pptx 没有这个 API，只能往 tcPr 里塞 a:ln。

    CT_TableCellProperties 的子元素顺序是固定的：
    lnL, lnR, lnT, lnB, lnTlToBr, lnBlToTr, cell3D, <fill>, headers, extLst。
    塞错位置 PowerPoint 会报文件损坏，所以统一插在第一个「填充类/尾部」元素之前。
    """
    from lxml import etree
    from pptx.oxml.ns import qn
    tcPr = cell._tc.get_or_add_tcPr()
    anchors = ("a:noFill", "a:solidFill", "a:gradFill", "a:blipFill",
               "a:pattFill", "a:grpFill", "a:headers", "a:extLst")
    anchor = None
    for a in anchors:
        anchor = tcPr.find(qn(a))
        if anchor is not None:
            break
    for tag in edges:                       # 按 L,R,T,B 顺序插，位置自然正确
        name = qn("a:ln" + tag)
        for old in tcPr.findall(name):
            tcPr.remove(old)
        ln = etree.Element(name)
        ln.set("w", str(int(weight * 12700)))
        ln.set("cap", "flat")
        ln.set("cmpd", "sng")
        ln.set("algn", "ctr")
        fill = etree.SubElement(ln, qn("a:solidFill"))
        clr = etree.SubElement(fill, qn("a:srgbClr"))
        clr.set("val", color)
        if anchor is not None:
            anchor.addprevious(ln)
        else:
            tcPr.append(ln)


def est_text_width(text: str, size: int, *, bold: bool = False) -> float:
    """粗估一行文字宽度（inch）。中文按 1.0em、西文按 0.52em，够用来分列宽。"""
    em = size / 72.0
    w = 0.0
    for ch in plain(text):
        if _CJK.search(ch):
            w += em
        elif ch in "iljtIf.,:;'|!":
            w += em * 0.30
        elif ch in "MWmw@%":
            w += em * 0.80
        else:
            w += em * 0.52
    return w * (1.06 if bold else 1.0)


def clear_slide(slide):
    """删掉版式带来的所有占位符，留空画布自己排。"""
    for sh in list(slide.shapes):
        sh._element.getparent().remove(sh._element)


# ---------------------------------------------------------------- 图注排法判定
# 11pt 楷体：行高约 0.186in，段后 0
CAP_LINE_H = 0.186


def caption_lines(text: str, box_w: float, size: int = 11) -> int:
    """估图注塞进 box_w 宽要占几行。"""
    total = est_text_width(text, size) * CAP_WRAP_LOSS
    per_line = max(box_w - 0.08, 0.5)      # 扣左右内边距
    import math
    return max(1, int(math.ceil(total / per_line)))


def line_h(size: int = 11) -> float:
    """字号对应的行高（含行距余量）。"""
    return CAP_LINE_H * (size / 11.0)


def caption_fits(text: str, box_w: float, box_h: float, size: int = 11) -> bool:
    return caption_lines(text, box_w, size) * line_h(size) <= box_h


# 断行的优先级：一句话讲完 > 一个分句讲完 > 一个顿号 > 硬切
_CUT_MARKS = "。！？；.!?;"
_SOFT_MARKS = "，、,：:）) "


def split_caption(text: str, box_w: float, box_h: float, size: int = 11):
    """把图注切成「塞得下的一段」+「剩下的」，切点尽量落在句子边界。

    返回 (head, tail)；塞得下时 tail 为空。
    size 必须和真正写字的字号一致 —— 按 11pt 估、用 16pt 写，
    估出来放得下、渲出来就出界（踩过）。
    """
    if not text:
        return "", ""
    # 留一行余量，估宽本身有误差
    usable = box_h - line_h(size)
    if caption_fits(text, box_w, usable, size):
        return text, ""

    lo, hi = 0, len(text)
    while lo < hi:
        mid = (lo + hi + 1) // 2
        if caption_fits(text[:mid], box_w, usable, size):
            lo = mid
        else:
            hi = mid - 1
    cut = lo
    if cut <= 0:
        return "", text

    # 往回找一个像样的断点，最多退 25% 的长度
    floor = max(1, int(cut * 0.75))
    for marks in (_CUT_MARKS, _SOFT_MARKS):
        for i in range(cut, floor, -1):
            if text[i - 1] in marks:
                cut = i
                break
        if cut < lo:
            break

    head, tail = text[:cut].rstrip(), text[cut:].lstrip()
    return head, tail


# ---------------------------------------------------------------- 图片等比
def fit_contain(path: str, x, y, w, h) -> tuple[float, float, float, float]:
    """等比缩放居中塞进 (x, y, w, h)，返回实际 (x, y, w, h)。"""
    from PIL import Image
    with Image.open(path) as im:
        iw, ih = im.size
    ar = iw / ih
    if ar >= w / h:
        nw, nh = w, w / ar
    else:
        nh, nw = h, h * ar
    return (x + (w - nw) / 2, y + (h - nh) / 2, nw, nh)


def place_image(slide, path, x, y, w, h):
    nx, ny, nw, nh = fit_contain(path, x, y, w, h)
    return slide.shapes.add_picture(path, Inches(nx), Inches(ny),
                                    Inches(nw), Inches(nh))


# ---------------------------------------------------------------- 字体审核
def audit_fonts(slide, *, allowed_cjk=(FONT_CJK, FONT_NUM, FONT_COVER_EN),
                allowed_latin=(FONT_LATIN, FONT_COVER_EN, FONT_NUM)):
    """返回该页所有不在白名单里的 (形状名, 字体名) —— 主要用于 check_deck。"""
    bad = []
    for sh in slide.shapes:
        if not sh.has_text_frame:
            continue
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                name = r.font.name or ""
                if not name:
                    continue
                pool = allowed_cjk if _CJK.search(r.text or "") else allowed_latin
                if name not in pool and name not in (FONT_LATIN, FONT_CJK):
                    bad.append((sh.name, name))
    return bad

# -*- coding: utf-8 -*-
"""把 content.json 编译成文献汇报 PPT。

用法：
    python build_deck.py content.json -o out.pptx
    python build_deck.py content.json -o out.pptx --template my_template.pptx

设计原则：本脚本只负责「排」，不负责「写」。所有文字都在 content.json 里，
想改文案改 json 就行，不用碰代码。
"""
from __future__ import annotations

import argparse
import json
import os
import sys

from pptx import Presentation
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import lr_common as L  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL = os.path.dirname(HERE)
DEFAULT_TEMPLATE = os.path.join(SKILL, "assets", "template.pptx")


# ------------------------------------------------------------------ 背景
def _bg(slide, color=None):
    L.set_bg(slide, color)


# ------------------------------------------------------------------ 顶部装饰
def _top_bar(slide, logo=True):
    """封面/致谢用的顶栏。样本实测：白底矩形不是蓝条，蓝的是下面那条 1.16 的分隔线。"""
    L.rect(slide, 0, 0.06, 13.26, 0.91, fill=L.WHITE)
    L.hline(slide, 0, 1.16, 13.33, L.BLUE, 1.0)
    if logo:
        p = os.path.join(SKILL, "assets", "logo.png")
        if os.path.exists(p):
            slide.shapes.add_picture(p, Inches(0.15), Inches(0.12),
                                     Inches(1.00), Inches(1.00))
    # 校名：样本是一行 18pt Verdana 粗 + 一行 14pt 斜体，同框左对齐、居中对齐块。
    box, tf = L.textbox(slide, 1.21, 0.33, 7.48, 0.64,
                        anchor=MSO_ANCHOR.MIDDLE,
                        margins=(0.04, 0.0, 0.04, 0.0))
    tf.word_wrap = True
    p1 = tf.paragraphs[0]
    p1.line_spacing = 1.0
    p1.space_after = Pt(0)
    L.add_runs(p1, "Nanyang Normal University", 18, bold=True, color=L.BLUE,
               font_latin=L.FONT_COVER_EN, font_cjk=L.FONT_COVER_EN)
    p2 = tf.add_paragraph()
    p2.line_spacing = 1.0
    p2.space_after = Pt(0)
    L.add_runs(p2, "College of Physics and Electronic Engineering", 14,
               color=L.BLUE, font_latin=L.FONT_COVER_EN, font_cjk=L.FONT_COVER_EN)
    for r in p2.runs:
        r.font.italic = True


# ------------------------------------------------------------------ 封面
def _tag_text(slide, text, y=None, x=None, h=None):
    """左上角的章节标签文字。**没有框**，就是竖排的一行正文字。

    原来这里是「蓝色虚线框 + 13pt 粗 主蓝竖排字」。用户 2026-11-02 说
    「框去一下，不需要了」，随后补一句「框删了之后里面的字不要删，
    同时给他从小蓝字变为和正文同字体」，再反问「原来的框不是竖着的吗」——
    意思很清楚：框去掉，字留着，还竖着，只是换成正文字（楷体 11pt 正文色）。
    """
    if not text:
        return
    x = L.TAG_X if x is None else x
    y = L.TAG_Y if y is None else y
    h = L.TAG_H if h is None else h
    box, tf = L.textbox(slide, x, y, L.TAG_W, h, anchor=MSO_ANCHOR.TOP,
                        margins=(0.0, 0.0, 0.0, 0.0))
    # 竖排：东亚竖排从上往下读
    tf._txBody.bodyPr.set("vert", "eaVert")
    L.write_box(tf, text, L.C1, color=L.BODY, align=PP_ALIGN.CENTER,
                line_spacing=1.10, space_after=0)


def build_cover(prs, meta):
    slide = prs.slides.add_slide(prs.slide_layouts[L.LAYOUT_TITLE_CONTENT])
    L.clear_slide(slide)
    _bg(slide, None)
    _top_bar(slide)

    # 主标题：样本 (4.68, 3.01) 3.91×1.21，66pt 华文楷体粗 132E4A 居中
    main = meta.get("cover_text", "文献汇报")
    box, tf = L.textbox(slide, 2.60, 2.55, 8.13, 1.21, anchor=MSO_ANCHOR.BOTTOM)
    L.write_box(tf, main, L.H0, bold=True, color=L.INK,
                align=PP_ALIGN.CENTER, font_cjk=L.FONT_CJK)

    # 副标题：样本没有，但中文题名有 40 字，66pt 一行放不下，只能另起。
    sub = meta.get("subtitle")
    if sub:
        box, tf = L.textbox(slide, 1.10, 3.86, 11.13, 0.90, anchor=MSO_ANCHOR.TOP)
        L.write_box(tf, sub, L.H2, color=L.MUTED, align=PP_ALIGN.CENTER,
                    line_spacing=1.32)

    # 汇报人行：样本 (3.87, 5.21) 5.91 宽 —— 样本自己把「日」字挤到下一行。
    # 位置不变，只把宽度放到 7.2 并居中对齐，视觉上与原稿一致，不再换行。
    # 姓名/时间没给就写 XX，绝不替我编。
    box, tf = L.textbox(slide, 3.07, 5.21, 7.20, 0.44, anchor=MSO_ANCHOR.MIDDLE)
    L.write_box(tf, who_line(meta), 20, bold=True, color=L.BODY,
                align=PP_ALIGN.CENTER)
    return slide


# ------------------------------------------------------------------ 文章首页
def build_paper_home(prs, spec, base_dir):
    """第二页：文章首页截图 + 作者信息。

    样本 P2 只有一张整页截图、一个字都不写；制作规范却要求配作者信息，
    所以这里是「样本构图 + 规范内容」的结合：左边占 6 成截首页，右边作者卡。
    这一页的英文（截图里的题名、作者）**保持原样不译** —— 译了就对不上原文了。
    """
    slide = prs.slides.add_slide(prs.slide_layouts[L.LAYOUT_BLANK])
    _bg(slide, None)

    img = spec.get("image")
    iw, ih = 7.15, 6.30
    if img:
        path = _resolve(img, base_dir)
        if os.path.exists(path):
            L.place_image(slide, path, 0.72, 0.62, iw, ih)
        else:
            _missing_box(slide, 0.72, 0.62, iw, ih, img)

    # 右侧作者卡
    cx, cw = 8.28, 4.42
    heading = spec.get("heading", "文献来源")
    box, tf = L.textbox(slide, cx, 0.60, cw, 0.42, anchor=MSO_ANCHOR.MIDDLE)
    L.write_box(tf, heading, L.H1, bold=True, color=L.INK)
    L.hline(slide, cx, 1.12, cw, L.RULE, 0.75)

    y = 1.30
    for a in spec.get("authors", []):
        box, tf = L.textbox(slide, cx, y, cw, 0.26)
        L.write_box(tf, a.get("role", ""), L.C2, bold=True, color=L.BLUE,
                    space_after=0, line_spacing=1.05)
        y += 0.24
        box, tf = L.textbox(slide, cx, y, cw, 0.32)
        L.write_box(tf, a.get("name", ""), L.B2, bold=True, color=L.INK,
                    space_after=0, line_spacing=1.05)
        y += 0.34
        for line in filter(None, [a.get("affil"), a.get("note")]):
            box, tf = L.textbox(slide, cx, y, cw, 0.28)
            L.write_box(tf, line, L.C1, color=L.MUTED, space_after=0, line_spacing=1.12)
            y += 0.24
        y += 0.14

    hl = spec.get("highlights") or []
    if hl:
        yy = y + 0.10
        L.hline(slide, cx, yy, cw, L.RULE, 0.75)
        yy += 0.14
        for h in hl:
            box, tf = L.textbox(slide, cx, yy, cw, 0.30)
            L.write_box(tf, "· " + h, L.C1, color=L.BODY, space_after=0, line_spacing=1.15)
            yy += 0.30
    return slide


# ------------------------------------------------------------------ 摘要精读
def build_summary(prs, spec):
    slide = prs.slides.add_slide(prs.slide_layouts[L.LAYOUT_BLANK])
    _bg(slide, None)
    _tag_text(slide, spec.get("tag"))

    box, tf = L.textbox(slide, 0.95, 0.26, 11.90, 0.48)
    L.write_box(tf, spec.get("heading", "摘要说了什么"), L.H1, bold=True, color=L.INK)
    L.hline(slide, 0.95, 0.80, 11.90, L.RULE, 0.75)

    items = spec.get("items", [])[:6]
    if not items:
        return slide
    top, bottom = 1.02, 6.85
    gap = 0.16
    n = len(items)
    # 每条按内容给高度，短条目不要撑满
    unit = (bottom - top - gap * (n - 1)) / n
    y = top
    for it in items:
        h = unit
        L.rect(slide, 1.00, y + 0.06, 0.06, h - 0.12, fill=L.BLUE)
        box, tf = L.textbox(slide, 1.22, y + 0.02, 1.35, h - 0.04, anchor=MSO_ANCHOR.MIDDLE)
        L.write_box(tf, it.get("no", ""), L.H2, bold=True, color=L.BLUE,
                    align=PP_ALIGN.CENTER)
        box, tf = L.textbox(slide, 2.72, y + 0.02, 9.95, h - 0.04, anchor=MSO_ANCHOR.MIDDLE)
        t = it.get("title", "")
        d = it.get("text", "")
        # 标题粗体独立成段，正文用同一文本框的第二段
        p = tf.paragraphs[0]
        p.line_spacing = 1.18
        p.space_after = Pt(2)
        L.add_runs(p, t, L.B2, bold=True, color=L.INK)
        if d:
            p2 = tf.add_paragraph()
            p2.line_spacing = 1.20
            p2.space_after = Pt(0)
            L.add_runs(p2, d, L.B1, color=L.BODY)
        y += h + gap
    return slide


# ------------------------------------------------------------------ 表格
def build_table(prs, spec):
    slide = prs.slides.add_slide(prs.slide_layouts[L.LAYOUT_BLANK])
    _bg(slide, None)
    # 样本 P5 是「一张光板表居中浮着」，没有章节标签、没有标题行。
    # 我早期把它拉满整页（11.90 宽 + 行高铺满）反而更难看，这里照样本。

    cols = spec.get("columns", [])
    rows = spec.get("rows", [])
    if not cols:
        return slide

    n_r, n_c = len(rows) + 1, len(cols)
    # 样本表宽 10.29 居中 (x=1.65)，顶边 y=1.85
    width = spec.get("width") or 10.29
    left = (L.CANVAS_W - width) / 2
    top = spec.get("top") or 1.62
    row_h = spec.get("row_height") or 0.62

    # 列宽按内容估算，不再等分 —— 等分会把「3.8 %」这种窄列撑成半个页面，
    # 而「近红外发光最强，全文主样品」被挤成两行。
    pad = 0.24
    ideal = []
    for j in range(n_c):
        cells = [cols[j]] + [(row[j] if j < len(row) else "") for row in rows]
        wmax = max(L.est_text_width(str(c), L.B1) for c in cells)
        hmax = L.est_text_width(str(cols[j]), L.B2, bold=True)
        ideal.append(max(wmax, hmax) + pad)
    # 第 1 列（行首标签）多给一点，读起来有落点
    if n_c > 1:
        ideal[0] += 0.30
    total_ideal = sum(ideal)
    col_w = [w * width / total_ideal for w in ideal]

    shape = slide.shapes.add_table(n_r, n_c, Inches(left), Inches(top),
                                   Inches(width), Inches(row_h * n_r))
    table = shape.table
    table.first_row = True
    for i in range(n_r):
        table.rows[i].height = Inches(row_h)
    for j in range(n_c):
        table.columns[j].width = Inches(col_w[j])

    for j, c in enumerate(cols):
        cell = table.cell(0, j)
        cell.fill.solid()
        cell.fill.fore_color.rgb = L.rgb("EAF1F8")
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.margin_left = cell.margin_right = Inches(0.10)
        L.cell_border(cell, "C8D8E8", 0.75)
        tf = cell.text_frame
        tf.word_wrap = True
        L.write_box(tf, str(c), L.B2, bold=True, color=L.INK,
                    align=PP_ALIGN.CENTER, space_after=0, line_spacing=1.10)

    for i, row in enumerate(rows, start=1):
        for j in range(n_c):
            val = row[j] if j < len(row) else ""
            cell = table.cell(i, j)
            cell.fill.solid()
            cell.fill.fore_color.rgb = L.rgb(L.WHITE if i % 2 else "FAFBFC")
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = cell.margin_right = Inches(0.10)
            L.cell_border(cell, L.RULE, 0.75)
            tf = cell.text_frame
            tf.word_wrap = True
            numeric = _looks_numeric(val)
            # 列比内容宽很多时，右对齐会在数字和表头之间留一道空；
            # 这种窄内容宽列一律居中，列宽和内容接近时才用右对齐对齐位数。
            if numeric and j > 0:
                content_w = L.est_text_width(str(val), L.B1)
                align = (PP_ALIGN.CENTER
                         if col_w[j] > content_w * 1.8 else PP_ALIGN.RIGHT)
            else:
                align = PP_ALIGN.LEFT
            L.write_box(tf, str(val) if val is not None else "", L.B1,
                        color=L.BODY, space_after=0, line_spacing=1.10,
                        align=align)

    cap = spec.get("caption")
    if cap:
        box, tf = L.textbox(slide, left, top + row_h * n_r + 0.18, width, 0.70)
        L.write_box(tf, cap, L.C1, color=L.MUTED, line_spacing=1.22)
    return slide


def _looks_numeric(v) -> bool:
    if v is None:
        return False
    s = str(v).strip().replace("%", "").replace("~", "").replace("±", "")
    try:
        float(s)
        return True
    except ValueError:
        return False


# ------------------------------------------------------------------ 文本页
def build_text(prs, spec):
    slide = prs.slides.add_slide(prs.slide_layouts[L.LAYOUT_BLANK])
    _bg(slide, None)
    _tag_text(slide, spec.get("tag"))

    box, tf = L.textbox(slide, 0.95, 0.26, 11.90, 0.48)
    L.write_box(tf, spec.get("heading", "合成方法"), L.H1, bold=True, color=L.INK)
    L.hline(slide, 0.95, 0.80, 11.90, L.RULE, 0.75)

    blocks = spec.get("blocks", [])
    if not blocks:
        return slide
    top, bottom = 1.05, 6.90
    unit = (bottom - top) / len(blocks)
    y = top
    for b in blocks:
        # 垂直淡色分隔线，把块与块分开
        L.rect(slide, 0.95, y + 0.08, 0.05, unit - 0.28, fill=L.BLUE)
        box, tf = L.textbox(slide, 1.22, y + 0.02, 11.60, unit - 0.18)
        p = tf.paragraphs[0]
        p.line_spacing = 1.15
        p.space_after = Pt(3)
        L.add_runs(p, b.get("title", ""), L.H2, bold=True, color=L.INK)
        for line in (b.get("body") or []):
            p2 = tf.add_paragraph()
            p2.line_spacing = 1.30
            p2.space_after = Pt(2)
            L.add_runs(p2, line, L.B1, color=L.BODY)
        y += unit
    return slide


# ------------------------------------------------------------------ 主图页
def _pick_layout(cap_text):
    """按图注实际占的行数选排法，优先级：底栏整段放得下 > 右栏整段放得下 > 右栏+续页。

    这是样本本来的做法，不是风格偏好 —— P9/P12/P17 是短注走底栏，
    P8/P16/P21/P26 是长注走右栏。

    ⚠️ 判定必须和「真的会切成几段」用同一套尺子。
    曾经写成「放不进底栏 → 就地切页」，结果图 5 的 287 字差一行挤不进底栏，
    就被劈成两页，续页上只剩「实验测得 k2 = k2a + k2b。」一句 —— 纯属自找的。
    所以这里直接问 split_caption：底栏切完还剩不剩，剩了就换右栏再问一遍。
    """
    if not cap_text:
        return "A"
    _, tail = L.split_caption(cap_text, L.FIG_A_CAP[2], L.FIG_A_CAP[3], size=L.C1)
    if not tail:
        return "A"
    _, tail = L.split_caption(cap_text, L.FIG_B_CAP[2], L.FIG_B_CAP[3], size=L.C1)
    if not tail:
        return "B"
    return "B"          # 右栏也放不下 → build_figure 会自动续页


def build_figure(prs, spec, base_dir, page_no):
    slide = prs.slides.add_slide(prs.slide_layouts[L.LAYOUT_BLANK])
    _bg(slide, None)

    cap_text = spec.get("caption", "")
    mode = _pick_layout(cap_text)

    if mode == "A":
        ix, iy, iw, ih = L.FIG_A_IMG
        cx, cy, cw, ch = L.FIG_A_CAP
        # A 排法图从 x=0.61 起，左边塞不下竖排标签，样本上这种页也没有标签
    else:
        ix, iy, iw, ih = L.FIG_B_IMG
        cx, cy, cw, ch = L.FIG_B_CAP
        # B 排法图从 1.35 起，左边 0.30 起有一整条空档，标签就竖在这儿
        _tag_text(slide, spec.get("tag"))

    img = spec.get("image")
    path = _resolve(img, base_dir) if img else ""
    if path and os.path.exists(path):
        L.place_image(slide, path, ix, iy, iw, ih)
    else:
        _missing_box(slide, ix, iy, iw, ih, img or "(未提供图片)")

    # 图注放不下就续页，不缩字号也不硬塞 —— 宁可多开一页。
    head, tail = L.split_caption(cap_text, cw, ch) if cap_text else ("", "")
    if head:
        box, tf = L.textbox(slide, cx, cy, cw, ch, anchor=MSO_ANCHOR.TOP)
        L.write_box(tf, head, L.C1, color=L.BODY,
                    align=PP_ALIGN.JUSTIFY, line_spacing=1.22, space_after=0)

    taken = page_no
    while tail:
        taken += 1
        tail = _caption_overflow(prs, spec, tail, taken, path)
    return slide


def _caption_overflow(prs, spec, text, page_no, img_path):
    """图注续页：仍然「图 + 注」一体，图缩到上半页，剩下的图注铺在下半页。

    不能只把图注挪过去 —— 那样讲这一页时图在旁边根本看不到，
    图注里说的 a/b/c 子图就全悬空了。所以续页把同一张图带上，
    只是占得小一点（图注在这页是整幅宽，行数少，用不着那么大图）。
    """
    slide = prs.slides.add_slide(prs.slide_layouts[L.LAYOUT_BLANK])
    _bg(slide, None)

    fig_no = spec.get("label") or spec.get("heading") or "图注"
    box, tf = L.textbox(slide, 0.95, 0.30, 11.30, 0.34, anchor=MSO_ANCHOR.TOP)
    L.add_runs(tf.paragraphs[0], "%s（续）" % fig_no, L.B2, bold=True, color=L.MUTED)

    # 图：上半页居中，等比 contain
    if img_path and os.path.exists(img_path):
        L.place_image(slide, img_path, 1.60, 0.70, 10.10, 3.55)
    else:
        _missing_box(slide, 1.60, 0.70, 10.10, 3.55, fig_no)

    # 注：下半页整幅宽，字号跟主页面上的图注一致（都是 C1），
    # 不然一句话在两页之间突然变大，读起来像换了个东西
    x, y, w, h = 0.95, 4.48, 11.30, 2.70
    head, tail = L.split_caption(text, w, h, size=L.C1)
    box, tf2 = L.textbox(slide, x, y, w, h, anchor=MSO_ANCHOR.TOP)
    L.write_box(tf2, head, L.C1, color=L.BODY,
                align=PP_ALIGN.JUSTIFY, line_spacing=1.30, space_after=4)
    return tail


# ------------------------------------------------------------------ 致谢
def build_thanks(prs, meta, spec):
    slide = prs.slides.add_slide(prs.slide_layouts[L.LAYOUT_TITLE_CONTENT])
    L.clear_slide(slide)
    _bg(slide, None)
    _top_bar(slide)

    text = spec.get("text", "感谢各位老师批评指正，谢谢大家！")
    box, tf = L.textbox(slide, 1.10, 2.45, 11.13, 1.08, anchor=MSO_ANCHOR.MIDDLE)
    L.write_box(tf, text, 40, bold=True, color=L.INK, align=PP_ALIGN.CENTER)

    img = spec.get("image")
    if img:
        path = _resolve(img, os.path.dirname(os.path.abspath(__file__)))
        if os.path.exists(path):
            L.place_image(slide, path, 1.65, 3.75, 10.03, 2.30)

    box, tf = L.textbox(slide, 3.07, 6.35, 7.20, 0.44, anchor=MSO_ANCHOR.MIDDLE)
    L.write_box(tf, who_line(meta), 18, bold=True, color=L.MUTED,
                align=PP_ALIGN.CENTER)
    return slide


# ------------------------------------------------------------------ 工具
PLACEHOLDER = "XX"


def who_line(meta):
    """封面/致谢页的汇报人行。

    没填就写 XX —— 宁可留个显眼的占位，也别替我编一个人名或日期。
    填了（content.json 的 meta，或命令行 --presenter/--date）就用填的。
    """
    who = (meta.get("presenter") or "").strip() or PLACEHOLDER
    when = (meta.get("date") or "").strip() or PLACEHOLDER
    return "汇报人：{}      时间：{}".format(who, when)


def _resolve(p, base_dir):
    return p if os.path.isabs(p) else os.path.normpath(os.path.join(base_dir, p))


def _missing_box(slide, x, y, w, h, label):
    sh = L.rect(slide, x, y, w, h, fill="F4F6F8", line=L.RULE, line_w=1.0, dash="DASH")
    tf = sh.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    L.add_runs(p, "缺图：{}".format(label), L.B2, color=L.MUTED)


# ------------------------------------------------------------------ 主流程
BUILDERS = {
    "summary": build_summary,
    "table": build_table,
    "text": build_text,
}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("content")
    ap.add_argument("-o", "--out", required=True)
    ap.add_argument("--template", default=DEFAULT_TEMPLATE)
    # 用户临时口头说了姓名/时间，不想改 json 就直接传这两个
    ap.add_argument("--presenter", help="汇报人姓名；不给就用 content.json 的 meta，再不给写 XX")
    ap.add_argument("--date", help="汇报时间；同上")
    args = ap.parse_args()

    with open(args.content, encoding="utf-8") as f:
        spec = json.load(f)
    base_dir = os.path.dirname(os.path.abspath(args.content))

    prs = Presentation(args.template)
    # 模板里已删光 slide，但保证尺寸正确
    prs.slide_width = Inches(L.CANVAS_W)
    prs.slide_height = Inches(L.CANVAS_H)

    meta = dict(spec.get("meta", {}))
    if args.presenter:
        meta["presenter"] = args.presenter
    if args.date:
        meta["date"] = args.date
    build_cover(prs, meta)
    if spec.get("paper_home"):
        build_paper_home(prs, spec["paper_home"], base_dir)

    page = 2
    for s in spec.get("slides", []):
        page += 1
        t = s.get("type", "figure")
        if t == "figure":
            build_figure(prs, s, base_dir, page)
        elif t == "thanks":
            continue
        else:
            fn = BUILDERS.get(t)
            if fn is None:
                print("[skip] 未知页型 {}: {}".format(t, s.get("heading", "")),
                      file=sys.stderr)
                continue
            fn(prs, s)

    build_thanks(prs, meta, spec.get("thanks", {}) or {})

    out_dir = os.path.dirname(os.path.abspath(args.out))
    os.makedirs(out_dir, exist_ok=True)
    prs.save(args.out)
    print("saved {} slides -> {}".format(len(prs.slides.__iter__.__self__._sldIdLst),
                                         args.out))


if __name__ == "__main__":
    main()

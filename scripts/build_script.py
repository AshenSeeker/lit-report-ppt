#!/usr/bin/env python3
"""把 content.json 的 script 段排成一份能照着念的 Word 讲稿。

设计要点（三件事，别的都不做）：
  1. **页序取自 content.json 本身**，不手抄页码 —— 改了 PPT 讲稿自动跟上，
     不会出现「讲稿说第 8 页、PPT 已经变成第 9 页」。
  2. **不写进 PPT 备注栏**（用户 2026-11-02 决定）。备注栏在放映时看不见，
     真念的时候还得另开一个窗口，不如单独一份 Word 放在旁边。
  3. **每页给目标秒数**，末尾按 200 字/分钟核总时长。讲稿不是越全越好，
     是按时间配出来的。

这个脚本只管排版。讲什么话由人（或模型）写进 content.json 的 script 段。

用法：
  python build_script.py content.json -o out/讲稿.docx
"""
import argparse
import json
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Cm

# 中文正文档用宋体、标题用黑体，屏幕和打印都稳。
# 华文楷体本机没有，讲稿里不用它 —— 讲稿要能直接念，别赌字体。
FONT_CJK = "宋体"
FONT_HEAD = "黑体"
FONT_LATIN = "Times New Roman"

INK = RGBColor(0x1A, 0x1A, 0x1A)
MUTED = RGBColor(0x6B, 0x6B, 0x6B)
BLUE = RGBColor(0x2A, 0x7D, 0xC0)
RED = RGBColor(0xC0, 0x00, 0x00)

CPS = 250.0          # 中文口播速度：字/分钟（学术汇报正常语速 250 上下，200 是慢讲到念稿的速度）
_CN = "一二三四五六七八九十"


# ------------------------------------------------------------------ 上下角标
def parse_markup(text):
    """拆 ^{...} 上标 / _{...} 下标。返回 [(片段, 'sup'|'sub'|None)]。"""
    out, i, buf = [], 0, []
    while i < len(text):
        ch = text[i]
        if ch in "^_" and i + 1 < len(text) and text[i + 1] == "{":
            j = text.find("}", i + 2)
            if j > 0:
                if buf:
                    out.append(("".join(buf), None))
                    buf = []
                out.append((text[i + 2:j], "sup" if ch == "^" else "sub"))
                i = j + 1
                continue
        buf.append(ch)
        i += 1
    if buf:
        out.append(("".join(buf), None))
    return out


def plain(text):
    return "".join(s for s, _ in parse_markup(text))


def add_runs(p, text, size, *, bold=False, color=INK, italic=False):
    for seg, kind in parse_markup(text):
        if not seg:
            continue
        r = p.add_run(seg)
        r.font.size = Pt(size * 0.72 if kind else size)
        r.font.bold = bold
        r.font.italic = italic
        r.font.color.rgb = color
        r.font.name = FONT_LATIN
        r._element.rPr.rFonts.set(qn("w:eastAsia"), FONT_CJK)
        if kind == "sup":
            r.font.superscript = True
        elif kind == "sub":
            r.font.subscript = True
    return p


def para(doc, text="", size=11, *, bold=False, color=INK, italic=False,
         space_before=0, space_after=4, align=None, first_indent=True):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = 1.35
    if align is not None:
        p.alignment = align
    if first_indent and not bold:
        pf.first_line_indent = Cm(0.74)     # 中文两字缩进
    if text:
        add_runs(p, text, size, bold=bold, color=color, italic=italic)
    return p


def heading(doc, text, size=15, *, color=INK, space_before=14):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(6)
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.font.bold = True
    r.font.color.rgb = color
    r.font.name = FONT_LATIN
    r._element.rPr.rFonts.set(qn("w:eastAsia"), FONT_HEAD)
    return p


def hline(doc, color="D6DEE6"):
    """一条细横线分隔。Word 没有真 API，借段落底边框。"""
    from docx.oxml import OxmlElement
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(6)
    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)
    pbdr.append(bottom)
    pPr.append(pbdr)
    return p


# ------------------------------------------------------------------ 页序推导
def page_plan(spec):
    """按 build_deck.py 的规则推算真实页序，讲稿页码永远和 PPT 对得上。

    返回 [(页码, 页型, 内容 dict)]
    """
    plan = [(1, "cover", spec.get("meta", {}))]
    n = 1
    if spec.get("paper_home"):
        n += 1
        plan.append((n, "paper_home", spec["paper_home"]))
    for s in spec.get("slides", []):
        if s.get("type") == "thanks":
            continue
        n += 1
        plan.append((n, s.get("type", "figure"), s))
        # 主图页图注太长会续页，页数要跟着算
        if s.get("type") == "figure" and s.get("caption"):
            extra = _overflow_pages(s["caption"])
            for _ in range(extra):
                n += 1
                plan.append((n, "figure_cont", s))
    n += 1
    plan.append((n, "thanks", spec.get("thanks", {}) or {}))
    return plan


def _overflow_pages(caption):
    """复算图注续了几页。用和 build_deck 同一套估算函数。"""
    try:
        import lr_common as L
    except ImportError:
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        import lr_common as L
    tail = caption
    n = 0
    # 主图页那一屏能装多少
    _, tail = L.split_caption(tail, L.FIG_B_CAP[2], L.FIG_B_CAP[3], size=L.C1)
    while tail and n < 6:
        n += 1
        _, tail = L.split_caption(tail, 11.30, 2.70, size=L.C1)
    return n


TYPE_NAMES = {
    "cover": "封面",
    "paper_home": "文献来源",
    "summary": "摘要",
    "table": "表格",
    "text": "正文",
    "figure": "主图",
    "figure_cont": "主图续页",
    "thanks": "致谢",
}


def page_title(typ, data):
    if typ == "cover":
        return data.get("cover_text", "文献汇报")
    if typ == "thanks":
        return ""
    if typ == "figure_cont":
        return "%s（续）" % (data.get("label") or "图注")
    title = data.get("heading") or data.get("tag") or ""
    # 页型名和标题撞了就不重复印（P2 曾经印成「文献来源 · 文献来源」）
    if title == TYPE_NAMES.get(typ, typ):
        return ""
    return title


# ------------------------------------------------------------------ 正文
def build(doc, spec):
    script = spec.get("script", {}) or {}
    meta = spec.get("meta", {})
    plan = page_plan(spec)

    # ---- 抬头
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("文献汇报讲稿")
    r.font.size = Pt(22)
    r.font.bold = True
    r.font.color.rgb = BLUE
    r.font.name = FONT_LATIN
    r._element.rPr.rFonts.set(qn("w:eastAsia"), FONT_HEAD)

    if meta.get("subtitle"):
        for line in meta["subtitle"].split("\n"):
            para(doc, line, 12, align=WD_ALIGN_PARAGRAPH.CENTER,
                 color=MUTED, first_indent=False, space_after=2)

    tail = "汇报人：%s" % ((meta.get("presenter") or "").strip() or "XX")
    tail += "　　时间：%s" % ((meta.get("date") or "").strip() or "XX")
    para(doc, tail, 10.5, align=WD_ALIGN_PARAGRAPH.CENTER,
         color=MUTED, first_indent=False, space_before=4)

    # ---- 讲前十分钟（有内容才出这一节，没写整段跳过）
    thesis = script.get("thesis")
    concepts = script.get("concepts") or []
    qa = script.get("qa") or []
    sec_no = 0
    if thesis or concepts or qa:
        sec_no += 1
        heading(doc, "%s、上台前十分钟看这一页" % _CN[sec_no - 1], 14, color=BLUE,
                space_before=16)
        hline(doc)

    if thesis:
        para(doc, "三句话讲清这篇论文", 11.5, bold=True, space_before=4, space_after=2)
        for i, line in enumerate(_as_list(thesis), 1):
            para(doc, "%d. %s" % (i, line), 11, space_after=3)

    if concepts:
        para(doc, "必须讲对的几个概念", 11.5, bold=True, space_before=10, space_after=2)
        for c in concepts:
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.line_spacing = 1.35
            p.paragraph_format.first_line_indent = Cm(0.74)
            add_runs(p, c.get("term", ""), 11, bold=True, color=INK)
            add_runs(p, "　" + c.get("plain", ""), 11, color=INK)

    qa = script.get("qa") or []
    if qa:
        para(doc, "可能被问到的问题", 11.5, bold=True, space_before=10, space_after=2)
        for i, q in enumerate(qa, 1):
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.35
            p.paragraph_format.first_line_indent = Cm(0.74)
            add_runs(p, "问：", 11, bold=True, color=RED)
            add_runs(p, q.get("q", ""), 11, color=INK)
            p2 = doc.add_paragraph()
            p2.paragraph_format.space_after = Pt(6)
            p2.paragraph_format.line_spacing = 1.35
            p2.paragraph_format.left_indent = Cm(0.74)
            add_runs(p2, "答：" + q.get("a", ""), 11, color=INK)

    # ---- 逐页：narration_only 模式下正文就是正文，不印「开口第一句」「指图」这些提示词
    narration_only = bool(script.get("narration_only"))
    sec_no += 1
    heading(doc, "%s、逐页讲稿" % _CN[sec_no - 1], 14, color=BLUE, space_before=16)
    hline(doc)

    # 讲稿条目不必和 PPT 逐页一一对应：一段话可以覆盖连续几页。
    # 写了 p 就按 p 覆盖（支持 3 / "3" / "2-3" / [2,3]），没写 p 就顺延到上一段之后。
    plan_by_no = dict((no, (typ, data)) for no, typ, data in plan)
    entries = []
    cursor = 1
    for sp in (script.get("pages") or []):
        nos = [n for n in _entry_pages(sp, cursor) if n in plan_by_no]
        if not nos:
            print("[warn] 这段讲稿标的页 %r 在 PPT 里不存在（共 %d 页），已顺延到 P%d"
                  % (sp.get("p"), len(plan), cursor if cursor in plan_by_no else 1),
                  file=sys.stderr)
            nos = [cursor if cursor in plan_by_no else min(plan_by_no)]
        entries.append((nos, sp))
        cursor = max(nos) + 1

    total_secs = 0
    timing = []
    for nos, sp in entries:
        first, last = nos[0], nos[-1]
        label = "P%d" % first if first == last else "P%d–P%d" % (first, last)
        typ = plan_by_no[first][0]
        data = plan_by_no[first][1]
        names = []
        for n in nos:
            nm = TYPE_NAMES.get(plan_by_no[n][0], plan_by_no[n][0])
            if nm not in names:
                names.append(nm)
        secs = sp.get("secs") or 0
        total_secs += secs
        timing.append((label, " / ".join(names), page_title(typ, data), secs))

        # 页头：P3 · 摘要 — 目标 120 秒
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.space_before = Pt(12)
        pf.space_after = Pt(3)
        head = "%s　%s" % (label, " / ".join(names))
        title = page_title(typ, data)
        if title:
            head += " · " + title
        r = p.add_run(head)
        r.font.size = Pt(12)
        r.font.bold = True
        r.font.color.rgb = BLUE
        r.font.name = FONT_LATIN
        r._element.rPr.rFonts.set(qn("w:eastAsia"), FONT_HEAD)
        if secs:
            r2 = p.add_run("　［目标 %d 秒］" % secs)
            r2.font.size = Pt(9.5)
            r2.font.color.rgb = MUTED
            r2.font.name = FONT_LATIN
            r2._element.rPr.rFonts.set(qn("w:eastAsia"), FONT_CJK)

        if not sp:
            para(doc, "（这页还没写讲稿）", 10.5, color=MUTED, space_after=4)
            continue

        opener = sp.get("opener")
        if opener:
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.line_spacing = 1.35
            if narration_only:
                add_runs(p, opener, 11.5, color=INK)
            else:
                add_runs(p, "开口第一句：", 10, bold=True, color=MUTED)
                add_runs(p, opener, 12, bold=True, color=INK)

        for line in _as_list(sp.get("body")):
            para(doc, line, 11.5, space_after=5)

        point = sp.get("point")
        if point and not narration_only:
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.line_spacing = 1.35
            p.paragraph_format.left_indent = Cm(0.4)
            add_runs(p, "指图：", 10, bold=True, color=MUTED)
            add_runs(p, point, 10.5, color=INK, italic=True)

        bridge = sp.get("bridge")
        if bridge and not narration_only:
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.line_spacing = 1.35
            add_runs(p, "过渡：", 10, bold=True, color=MUTED)
            add_runs(p, bridge, 10.5, color=MUTED, italic=True)

    # ---- 时间核
    heading(doc, "%s、时间表" % _CN[sec_no], 14, color=BLUE, space_before=18)
    hline(doc)

    words = 0
    for _nos, sp in entries:
        for line in _as_list(sp.get("body")):
            words += len(plain(line))
        words += len(plain(sp.get("opener") or ""))

    tbl = doc.add_table(rows=1, cols=4)
    tbl.style = "Table Grid"
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = ["页", "类型", "内容", "秒"]
    for i, h in enumerate(hdr):
        cell = tbl.rows[0].cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_runs(p, h, 10, bold=True, color=INK)
    for label, typ, title, secs in timing:
        if not secs:
            continue
        row = tbl.add_row()
        vals = [label, typ, title, str(secs)]
        for i, v in enumerate(vals):
            cell = row.cells[i]
            cell.text = ""
            p = cell.paragraphs[0]
            if i in (0, 3):
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_runs(p, v, 10, color=INK)

    est = words / CPS * 60 if words else 0
    para(doc, "", 6, space_after=0, first_indent=False)
    line = "合计目标 %d 秒（约 %d 分 %d 秒）" % (
        total_secs, total_secs // 60, total_secs % 60)
    para(doc, line, 11, bold=True, space_after=3, first_indent=False)

    est_line = "讲稿正文 %d 字，按 %d 字/分钟约需 %d 秒" % (int(words), int(CPS), int(est))
    if total_secs and abs(est - total_secs) > max(30, total_secs * 0.25):
        est_line += "　⚠️ 和目标的 %d 秒差得多，要么删字要么改目标" % total_secs
    para(doc, est_line, 10.5, color=MUTED, space_after=6, first_indent=False)

    closing = script.get("closing")
    if closing:
        para(doc, "收尾", 11.5, bold=True, space_before=10, space_after=2)
        for line in _as_list(closing):
            para(doc, line, 11, space_after=3)


def _as_list(v):
    if not v:
        return []
    if isinstance(v, str):
        return [v]
    return list(v)


def _entry_pages(sp, cursor):
    """这段讲稿覆盖哪几页 PPT。

    写了 p 就按 p 走：3 / "3" / "2-3" / [2,3]。
    没写就顺延到上一段之后（cursor）。返回升序去重的页码列表。
    """
    raw = sp.get("p") if isinstance(sp, dict) else None
    if raw in (None, "", []):
        return [cursor]
    if isinstance(raw, int):
        return [raw]
    if isinstance(raw, (list, tuple)):
        out = []
        for v in raw:
            out.extend(_entry_pages({"p": v}, cursor))
        return sorted(set(out))
    text = str(raw).replace("–", "-").replace("—", "-").replace(" ", "")
    out = []
    for chunk in text.split(","):
        if not chunk:
            continue
        if "-" in chunk:
            a, _, b = chunk.partition("-")
            try:
                lo, hi = int(a.lstrip("Pp")), int(b.lstrip("Pp"))
            except ValueError:
                continue
            if lo > hi:
                lo, hi = hi, lo
            out.extend(range(lo, hi + 1))
        else:
            try:
                out.append(int(chunk.lstrip("Pp")))
            except ValueError:
                continue
    return sorted(set(out)) or [cursor]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("content")
    ap.add_argument("-o", "--out", required=True)
    ap.add_argument("--presenter", help="汇报人姓名；不给就用 content.json 的 meta，再不给写 XX")
    ap.add_argument("--date", help="汇报时间；同上")
    args = ap.parse_args()

    with open(args.content, encoding="utf-8") as f:
        spec = json.load(f)
    if args.presenter or args.date:
        spec["meta"] = dict(spec.get("meta", {}))
        if args.presenter:
            spec["meta"]["presenter"] = args.presenter
        if args.date:
            spec["meta"]["date"] = args.date

    doc = Document()
    # A4 页边距，行宽窄一点念着不串行
    sec = doc.sections[0]
    sec.top_margin = Cm(2.2)
    sec.bottom_margin = Cm(2.2)
    sec.left_margin = Cm(2.6)
    sec.right_margin = Cm(2.6)
    style = doc.styles["Normal"]
    style.font.name = FONT_LATIN
    style.font.size = Pt(11)
    style.element.rPr.rFonts.set(qn("w:eastAsia"), FONT_CJK)

    build(doc, spec)

    out_dir = os.path.dirname(os.path.abspath(args.out))
    os.makedirs(out_dir, exist_ok=True)
    doc.save(args.out)
    print("saved -> %s" % args.out)


if __name__ == "__main__":
    main()

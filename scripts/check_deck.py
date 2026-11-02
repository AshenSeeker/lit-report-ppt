#!/usr/bin/env python3
"""生成后的自检：把能机械查的错全查一遍，别靠人眼一页页翻。

查六类：
  1. 出界 —— 任何形状超出 13.333×7.5 画布（图注被切最常见）
  2. 重叠 —— 图片与图注框相交（样本 P8 那个毛病）
  3. 空注 —— 主图页没有图注，或图注没有子图编号
  4. 字体 —— 用了本机装不上的字体，换机器就变样
  5. 字号 —— 出现规范外的字号
  6. 缺图 —— 图片路径解析不到

用法：
  python check_deck.py out/deck.pptx
  python check_deck.py out/deck.pptx --content content.json   # 连内容源一起校验
"""
import argparse
import json
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from pptx import Presentation
from pptx.util import Emu

CANVAS_W, CANVAS_H = 13.333, 7.5
EPS = 0.02

# 规范内的字号（见 references/style-tokens.md）
ALLOWED_SIZES = {66, 60, 40, 22, 20, 18, 16, 14, 13, 11, 10}

errors = []
warns = []
info = []


def _inch(v):
    return round(Emu(v).inches, 3) if v is not None else 0.0


def _iter_shapes(shapes):
    for sh in shapes:
        yield sh
        if sh.shape_type == 6:            # GROUP
            for sub in _iter_shapes(sh.shapes):
                yield sub


def check_bounds(idx, shapes):
    for sh in _iter_shapes(shapes):
        if sh.left is None or sh.top is None:
            continue
        x, y = _inch(sh.left), _inch(sh.top)
        w, h = _inch(sh.width), _inch(sh.height)
        if x < -EPS or y < -EPS:
            errors.append("P%d 形状越出左上边界 x=%.2f y=%.2f" % (idx, x, y))
        if x + w > CANVAS_W + EPS:
            errors.append("P%d 形状越出右边界 x+w=%.2f > %.3f" % (idx, x + w, CANVAS_W))
        if y + h > CANVAS_H + EPS:
            errors.append("P%d 形状越出下边界 y+h=%.2f > %.3f（多半被切掉了）"
                          % (idx, y + h, CANVAS_H))


def check_figure_page(idx, shapes, sizes_found):
    pics = [s for s in _iter_shapes(shapes) if s.shape_type == 13]
    texts = [s for s in _iter_shapes(shapes) if s.has_text_frame
             and s.text_frame.text.strip()]
    # 跳过封面/致谢（有顶栏校名）
    if any("Nanyang Normal University" in t.text_frame.text for t in texts):
        return
    if not pics:
        return
    caps = sorted(texts, key=lambda t: len(t.text_frame.text), reverse=True)
    cap = caps[0].text_frame.text.strip()
    # 跳过文章首页那种「图 + 一堆卡片」的页：它的最长文字是文献信息，
    # 不是图注。判据是文字块多、且最长那段不以 图/表 开头。
    if len(texts) >= 5 and not cap.startswith(("图", "表")):
        return
    if not caps:
        errors.append("P%d 有图但一个字都没有 —— 主图页必须有图注" % idx)
        return
    if len(cap) < 40:
        warns.append("P%d 图注只有 %d 字，是不是漏了子图说明" % (idx, len(cap)))
    # 子图交代：要么有 (a)，要么用左/右/插图把结构讲清了。
    # 图 1 那种「左侧……右侧……」的纯示意图没有字母编号，是正常的。
    has_letter = any(m in cap for m in ("(a)", "（a）", "(A)"))
    has_spatial = any(m in cap for m in ("左侧", "右侧", "左为", "右为", "插图", "上图", "下图"))
    if not (has_letter or has_spatial):
        warns.append("P%d 图注没交代子图结构（既没有 (a) 也没有左/右/插图）" % idx)
    pic = pics[0]
    if _overlap(pic, caps[0]):
        errors.append("P%d 图注框和图片叠在一起了" % idx)


def _overlap(a, b):
    ax, ay, aw, ah = _inch(a.left), _inch(a.top), _inch(a.width), _inch(a.height)
    bx, by, bw, bh = _inch(b.left), _inch(b.top), _inch(b.width), _inch(b.height)
    ox = min(ax + aw, bx + bw) - max(ax, bx)
    oy = min(ay + ah, by + bh) - max(ay, by)
    if ox > 0.05 and oy > 0.05:
        # 都在容差内的小压边不算
        return ox > 0.15 and oy > 0.15
    return False


def check_fonts(prs):
    used = set()
    sizes_found = set()
    A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    for slide in prs.slides:
        for sh in _iter_shapes(slide.shapes):
            if not sh.has_text_frame:
                continue
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    if r.font.name:
                        used.add(r.font.name)
                    # 上下标的字号是按比例缩的（16pt × 0.72 ≈ 12pt），
                    # 它们不进字号规范检查，不然永远报 8/12 两个假问题
                    rPr = r._r.find(A + "rPr")
                    baseline = rPr.get("baseline") if rPr is not None else None
                    if r.font.size and baseline in (None, "0"):
                        sizes_found.add(round(r.font.size.pt))
                    if rPr is not None:
                        ea = rPr.find(A + "ea")
                        if ea is not None and ea.get("typeface"):
                            used.add(ea.get("typeface"))
    return used, sizes_found


def check_font_available(names):
    """查字体在本机装没装 —— 装不上的字体会被静默替换，换台机器就变样。"""
    windir = os.environ.get("WINDIR", r"C:\Windows")
    fonts_dir = os.path.join(windir, "Fonts")
    if not os.path.isdir(fonts_dir):
        return
    # 中文字体的显示名和文件名对不上，这里只做粗查 + 提示
    known_alias = {
        "楷体": ("simkai", "KaiTi"),
        "等线": ("deng", "DengXian"),
        "宋体": ("simsun",),
        "黑体": ("simhei",),
        "微软雅黑": ("msyh",),
    }
    files = [f.lower() for f in os.listdir(fonts_dir)]
    missing = []
    for name in names:
        if name.lower().startswith(("times", "arial", "verdana", "calibri", "cambria")):
            key = name.lower().split()[0]
            if not any(key in f for f in files):
                missing.append(name)
            continue
        keys = known_alias.get(name, ())
        if keys and not any(k in f for f in files for k in [keys[0]]):
            missing.append(name)
    if missing:
        warns.append("这些字体本机可能没有，会被静默替换：%s" % "、".join(missing))


def check_sizes(sizes_found):
    bad = sorted(s for s in sizes_found if s not in ALLOWED_SIZES)
    if bad:
        warns.append("出现规范外字号 %s（规范：%s）"
                     % (bad, sorted(ALLOWED_SIZES, reverse=True)))


def check_content(content_path, base_dir):
    with open(content_path, encoding="utf-8") as f:
        spec = json.load(f)
    figures = [s for s in spec.get("slides", []) if s.get("type") == "figure"]
    for i, s in enumerate(figures, 1):
        img = s.get("image")
        if not img:
            errors.append("图 %d 没写 image 字段" % i)
        else:
            p = img if os.path.isabs(img) else os.path.join(base_dir, img)
            if not os.path.exists(p):
                errors.append("图 %d 的图片找不到：%s" % (i, p))
        if not s.get("caption"):
            errors.append("图 %d 没有图注 —— 组内通病，必须有" % i)
    summary = [s for s in spec.get("slides", []) if s.get("type") == "summary"]
    for s in summary:
        n = len(s.get("items", []))
        if n > 6:
            warns.append("摘要页有 %d 条，建议拆两页（超过 6 条读不过来）" % n)
    return spec


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--content", help="同时校验内容源 content.json")
    args = ap.parse_args()

    prs = Presentation(args.pptx)
    info.append("共 %d 页" % len(prs.slides._sldIdLst))

    # 画布尺寸
    if abs(_inch(prs.slide_width) - CANVAS_W) > 0.01:
        errors.append("画布宽 %.2f，应为 %.3f" % (_inch(prs.slide_width), CANVAS_W))

    for idx, slide in enumerate(prs.slides, 1):
        check_bounds(idx, slide.shapes)
        check_figure_page(idx, slide.shapes, None)

    used, sizes = check_fonts(prs)
    info.append("用到字体：%s" % "、".join(sorted(used)))
    info.append("用到字号：%s" % sorted(sizes, reverse=True))
    check_font_available(used)
    check_sizes(sizes)

    if args.content:
        base = os.path.dirname(os.path.abspath(args.content))
        check_content(args.content, base)

    # ---- 输出
    for m in info:
        print("  " + m)
    for m in warns:
        print("  [warn] " + m)
    for m in errors:
        print("  [FAIL] " + m)

    if errors:
        print("\n%d 个必须修的问题，%d 个建议" % (len(errors), len(warns)))
        return 1
    print("\n通过，%d 个建议" % len(warns))
    return 0


if __name__ == "__main__":
    sys.exit(main())

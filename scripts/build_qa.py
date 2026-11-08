#!/usr/bin/env python3
"""把 qa.json 排成一份双色答辩预测问答 HTML（默认第三产物，必出）。

走 lit-report-ppt 这个 skill 就出三份产物（PPT + Word 讲稿 + 本 HTML），
本脚本是第三份，默认要交，不用用户单独点名。除非用户明确说「不要问答」。

这个脚本只管排版，问什么、答什么由人（或模型）写进 qa.json。
六条硬要求（和 PPT / 讲稿一样，是内容判断，脚本不替你写）：
  1. 十题左右，每题都有答案（问题栏与答案栏一一对应）。
  2. 问题蓝、答案绿 —— 不同颜色，上下角标写 ^{} / _{} 转成真 sup/sub。
  3. 只给「问题 + 直接答案」两栏，要言简意赅；不要考察内容 / 考察点 /
     解析 / 应答思路等任何附加内容。
  4. 答案 ≤30 字（只数汉字；化学式、符号、字母不计）—— 超了只警告，不阻断。
  5. 提问方向覆盖：为什么选这个来做 / 国内外哪里做得多 /
     可以换成别的做吗、进而勾出它们的性质（替换后性质怎么变）。
  6. 排版要整齐：表格、对齐、上下角标用 ^{} / _{}。

CSS 是照着一份真实成品逐字抄下来的，别凭感觉改颜色或间距 ——
改一处就会和预览图对不上。样式常量在下面 STYLE 里。

用法：
  python build_qa.py qa.json -o out/答辩预测问题.html
"""
import argparse
import html
import json
import os
import re
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# ---------------------------------------------------------------- 样式（照抄成品）
STYLE = """  :root{
    --q:#1d4ed8; --a:#15803d; --ink:#1f2937; --mut:#64748b;
    --line:#e2e8f0; --head:#f1f5f9; --alt:#f8fafc;
  }
  *{box-sizing:border-box}
  body{
    margin:0; padding:34px 30px 46px; background:#fff; color:var(--ink);
    font-family:"Microsoft YaHei","PingFang SC","Segoe UI",Arial,sans-serif;
    font-size:15px; line-height:1.65;
  }
  .wrap{max-width:1020px; margin:0 auto}
  h1{font-size:22px; margin:0 0 8px; letter-spacing:.4px}
  .sub{margin:0; font-size:13px; color:var(--mut)}
  .meta{margin:0 0 18px; font-size:12.5px; color:var(--mut)}
  .rule{height:2px; background:var(--line); margin:0 0 14px}
  .legend{margin:0 0 14px; font-size:12.5px; color:var(--mut)}
  .legend b{font-weight:600}
  .legend .q{color:var(--q)} .legend .a{color:var(--a)}
  table{width:100%; border-collapse:collapse; table-layout:fixed}
  col.no{width:52px} col.tag{width:76px} col.q{width:43%} col.a{width:auto}
  th{
    background:var(--head); text-align:left; font-size:12.5px; font-weight:600;
    color:#475569; padding:9px 12px; border:1px solid var(--line);
  }
  td{padding:11px 12px; border:1px solid var(--line); vertical-align:top}
  tbody tr:nth-child(even){background:var(--alt)}
  td.no{text-align:center; color:var(--mut); font-variant-numeric:tabular-nums}
  td.tag{text-align:center; color:#64748b; font-size:12.5px}
  td.q{color:var(--q); font-weight:600}
  td.a{color:var(--a)}"""

HEAD = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>{title}</title>
<style>
{style}
</style>
</head>
<body>
<div class="wrap">
  <h1>{title}</h1>
  <p class="sub">{subtitle}</p>
  <p class="meta">{meta}</p>
  <div class="rule"></div>
  <p class="legend"><b class="q">蓝色＝问题</b>　　<b class="a">绿色＝答案</b>　（{legend}）</p>

  <table>
    <colgroup><col class="no"><col class="tag"><col class="q"><col class="a"></colgroup>
    <thead>
      <tr><th>序号</th><th>类别</th><th>问题</th><th>答案</th></tr>
    </thead>
    <tbody>
"""

TAIL = """    </tbody>
  </table>
</div>
</body>
</html>
"""

# ---------------------------------------------------------------- 上下角标
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


def rich(text):
    """把 ^{} / _{} 转成 HTML，并对普通片段做转义。"""
    parts = []
    for seg, kind in parse_markup(text):
        if not seg:
            continue
        seg = html.escape(seg, quote=False)
        if kind:
            parts.append("<%s>%s</%s>" % (kind, seg, kind))
        else:
            parts.append(seg)
    return "".join(parts)


def cn_count(text):
    """数「字」：只算汉字，符号、字母、化学式都不计（用户口径：≤30 字，字母可以不限）。"""
    return sum(1 for c in text if "\u4e00" <= c <= "\u9fff")


# ---------------------------------------------------------------- 渲染
def render(spec):
    rows = spec.get("qa", [])
    meta = spec.get("citation", "").strip()
    if meta:
        meta += " ｜ 共 %d 题" % len(rows)
    else:
        meta = "共 %d 题" % len(rows)

    out = [HEAD.format(
        title=html.escape(spec.get("title", "文献汇报 · 答辩预测问题与答案"), quote=False),
        style=STYLE,
        subtitle=html.escape(spec.get("subtitle", ""), quote=False),
        meta=html.escape(meta, quote=False),
        legend=html.escape(spec.get("legend", "答案均 ≤30 字，符号与字母不计"), quote=False),
    )]

    over = []
    for i, row in enumerate(rows, 1):
        no = str(row.get("no", "%02d" % i))
        tag = str(row.get("tag", ""))
        q = str(row.get("q", ""))
        a = str(row.get("a", ""))
        n = cn_count(a)
        if n > 30:
            over.append("第 %s 题答案 %d 字（>30）" % (no, n))
        out.append(
            "      <tr><td class=\"no\">%s</td><td class=\"tag\">%s</td>\n"
            "        <td class=\"q\">%s</td>\n"
            "        <td class=\"a\">%s</td></tr>\n" % (
                html.escape(no, quote=False),
                html.escape(tag, quote=False),
                rich(q), rich(a),
            )
        )
        if i < len(rows):
            out.append("\n")

    out.append(TAIL)
    return "".join(out), over


def main():
    ap = argparse.ArgumentParser(description="qa.json -> 双色答辩预测问答 HTML")
    ap.add_argument("qa", help="问答数据 JSON（复制 assets/qa.template.json 改）")
    ap.add_argument("-o", "--out", required=True, help="输出 HTML 路径")
    args = ap.parse_args()

    with open(args.qa, encoding="utf-8") as f:
        spec = json.load(f)

    text, over = render(spec)

    out_dir = os.path.dirname(os.path.abspath(args.out))
    os.makedirs(out_dir, exist_ok=True)
    with open(args.out, "w", encoding="utf-8", newline="\n") as f:
        f.write(text)

    print("saved -> %s (%d 题)" % (args.out, len(spec.get("qa", []))))
    for w in over:
        print("  [警告] " + w)


if __name__ == "__main__":
    main()

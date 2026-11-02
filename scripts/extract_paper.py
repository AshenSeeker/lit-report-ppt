# -*- coding: utf-8 -*-
"""抽取论文 PDF 的文本、元信息与候选图注，给写 content.json 和讲稿用。

用法：
    python extract_paper.py paper.pdf -o paper.json
    python extract_paper.py paper.pdf -o paper.txt --text-only

输出 paper.json：
    {
      "meta": {"title","journal","year","doi","pages","source","authors_raw"},
      "pages": [{"no":1,"text":"..."}],
      "captions": ["Figure 1. ...", "Figure 2. a) ..."],   # 正文里抓到的图注
      "abstract": "...",                                    # 单独拎出来的摘要
      "full_text": "...",
      "chars": 48157
    }

引擎优先级：PyMuPDF > pypdf。
    PyMuPDF 的文本流正常带空格；pypdf 抽这种双栏排版的期刊会丢空格、
    还会把 fi/fl 连字拆坏（'ﬁrst' → '\\ufb01rst'）。只在没装 pymupdf 时用 pypdf 兜底。

⚠️ 只做抽取，不做判断。图注抓取是正则，漏抓错抓都是常态 ——
   抽出来的是**线索**，最终写进 content.json 的图注必须对着 PDF 逐条核过。
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys

# Windows 控制台默认 gbk，打印中文会炸。强制 UTF-8。
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:  # noqa: BLE001
        pass

CAPTION_RE = re.compile(
    r"^(?:Fig(?:ure)?\.?\s*\d+[a-z]?|Scheme\s*\d+|Table\s*\d+|图表\s*\d+)\s*[\.\:\s]",
    re.IGNORECASE)
DOI_RE = re.compile(r"\b(10\.\d{4,9}/[-._;()/:A-Za-z0-9]+)\b")
ABSTRACT_RE = re.compile(r"(?:Abstract|摘要)\s*[:：]\s*(.+?)(?=\n\s*\n|\bKeywords?\b|\b关键词\b|\Z)",
                         re.IGNORECASE | re.DOTALL)

JOURNALS = (
    r"(Angew\.?\s*Chem\.?[^\n]{0,70})",
    r"(J\.?\s*Am\.?\s*Chem\.?\s*Soc\.?[^\n]{0,40})",
    r"(Adv(?:anced)?\.?\s*Mater\.?[^\n]{0,40})",
    r"(Nature\s+[A-Z][a-z]+[^\n]{0,40})",
    r"(Nat\.?\s*[A-Z][a-z]+[^\n]{0,40})",
    r"((?:ACS|RSC|Adv\.)\s*[A-Z][A-Za-z]+\.[^\n]{0,40})",
    r"(Chem\.?\s*Eng\.?\s*J\.?[^\n]{0,40})",
    r"(Nano\s*Lett\.[^\n]{0,40})",
)


# ------------------------------------------------------------------ 引擎
def _pages_pymupdf(path):
    import pymupdf
    doc = pymupdf.open(path)
    pages, spans = [], []
    for i, pg in enumerate(doc, start=1):
        pages.append({"no": i, "text": pg.get_text("text")})
        if i == 1:
            d = pg.get_text("dict")
            for b in d["blocks"]:
                for ln in b.get("lines", []):
                    for s in ln["spans"]:
                        if s["text"].strip():
                            spans.append((round(s["size"], 1), s["text"].strip(),
                                          round(s["bbox"][1], 1)))
    return pages, spans


def _pages_pypdf(path):
    from pypdf import PdfReader
    reader = PdfReader(path, strict=False)
    pages = []
    for i, pg in enumerate(reader.pages, start=1):
        try:
            t = pg.extract_text() or ""
        except Exception as e:  # noqa: BLE001
            print("[warn] 第 {} 页抽取失败: {}".format(i, e), file=sys.stderr)
            t = ""
        pages.append({"no": i, "text": t})
    return pages, []


def load_pages(path):
    try:
        return _pages_pymupdf(path), "pymupdf"
    except ImportError:
        print("[info] 没装 pymupdf，退回 pypdf（文本可能丢空格）", file=sys.stderr)
        return _pages_pypdf(path), "pypdf"


# ------------------------------------------------------------------ 清洗
def clean(s: str) -> str:
    s = s.replace("\u00ad", "")          # 软连字符
    s = s.replace("\ufb01", "fi").replace("\ufb02", "fl")   # 连字
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s.strip()


def join_lines(raw: str) -> str:
    """PDF 抽出来是按视觉行断的，把被硬断开的句子接回去。"""
    out, buf = [], []
    for ln in raw.splitlines():
        s = ln.strip()
        if not s:
            if buf:
                out.append(" ".join(buf))
                buf = []
            out.append("")
            continue
        buf.append(s)
    if buf:
        out.append(" ".join(buf))
    text = "\n".join(out)
    # 中文行尾不该留空格
    text = re.sub(r"(?<=[\u4e00-\u9fff]) (?=[\u4e00-\u9fff])", "", text)
    return text


# ------------------------------------------------------------------ 标题
def guess_title(spans, page1_text):
    if spans:
        top = max(s for s, _, _ in spans)
        # 取最大字号里 y 最小的一段（标题总在最上面），按 y 聚行
        cands = [(y, t) for s, t, y in spans if s >= top - 0.2]
        cands.sort()
        # 同一 y 附近的拼起来
        lines, cur_y, cur = [], None, []
        for y, t in cands:
            if cur_y is None or abs(y - cur_y) <= 2.5:
                cur.append(t)
                cur_y = y if cur_y is None else cur_y
            else:
                lines.append(" ".join(cur))
                cur, cur_y = [t], y
        if cur:
            lines.append(" ".join(cur))
        title = " ".join(lines).strip()
        if 10 < len(title) < 400:
            return title

    # 兜底：首页里挑最像标题的一行
    for ln in page1_text.splitlines():
        s = ln.strip()
        if 25 < len(s) < 250 and not s.lower().startswith(("doi", "how to cite")):
            return s
    return ""


# ------------------------------------------------------------------ 图注
def grab_captions(page_texts: list[str]) -> list[str]:
    """逐行扫描**未合并**的页文本 —— 必须先看行首，所以不能传 join 过的全文。"""
    caps, cur = [], None
    for ptxt in page_texts:
        for ln in ptxt.splitlines():
            s = ln.strip()
            if CAPTION_RE.match(s):
                if cur:
                    caps.append(cur)
                cur = s
            elif cur is not None:
                if not s:
                    caps.append(cur)
                    cur = None
                elif len(cur) < 1400:
                    cur = cur + " " + s
                else:
                    caps.append(cur)
                    cur = None
    if cur:
        caps.append(cur)
    # 去掉明显是正文引用的（"Figure 1 shows that..."），再按图号去重取最长
    out, best = [], {}
    for c in caps:
        head = c[:60].lower()
        if re.search(r"\b(shows?|shown|seen|reveals?|indicates?|confirms?|suggests?)\b", head):
            continue
        if len(c) < 120:
            continue
        c = re.sub(r"\s+", " ", c).strip()
        m = re.match(r"(?i)(fig(?:ure)?|scheme|table)\s*\.?\s*(\d+)", c)
        key = "{} {}".format(m.group(1).lower()[:3], m.group(2)) if m else c[:24]
        # 同一图号里挑最长的那条（真正的图注总是比正文引用长得多）
        if key not in best or len(c) > len(best[key]):
            best[key] = c
    out = sorted(best.values(), key=lambda x: (x.lower().startswith("table"), len(x) * -1))
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("-o", "--out", required=True)
    ap.add_argument("--text-only", action="store_true")
    ap.add_argument("--no-fulltext", action="store_true")
    args = ap.parse_args()

    (pages, spans), engine = load_pages(args.pdf)
    pages = [{"no": p["no"], "text": clean(p["text"])} for p in pages]
    full = join_lines("\n".join(p["text"] for p in pages))
    first = pages[0]["text"] if pages else ""

    title = guess_title(spans, first)

    journal, year, doi = "", "", ""
    for pat in JOURNALS:
        m = re.search(pat, first)
        if m:
            journal = re.sub(r"\s+", " ", m.group(1)).strip()
            break
    m = DOI_RE.search(first) or DOI_RE.search(full[:4000])
    if m:
        doi = m.group(1).rstrip(".,;)")
    m = re.search(r"\b(19[89]\d|20[0-4]\d)\b", first)
    if m:
        year = m.group(0)

    m = ABSTRACT_RE.search(full)
    abstract = re.sub(r"\s+", " ", m.group(1)).strip() if m else ""

    # 作者行：标题下方、以逗号分隔、含 * 的那一行
    authors_raw = ""
    for ln in first.splitlines():
        s = ln.strip()
        if ("*" in s or "," in s) and 15 < len(s) < 300 and not s.lower().startswith(
                ("doi", "how to cite", "abstract")):
            if re.search(r"[A-Z][a-z]+\s+[A-Z][a-z]+", s):
                authors_raw = s
                break

    data = {
        "meta": {
            "title": title,
            "journal": journal,
            "year": year,
            "doi": doi,
            "pages": len(pages),
            "source": os.path.basename(args.pdf),
            "engine": engine,
            "authors_raw": authors_raw,
        },
        "pages": pages,
        "captions": grab_captions([p["text"] for p in pages]),
        "abstract": abstract,
        "full_text": full,
        "chars": len(full),
    }

    if args.text_only:
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(full)
        print("saved text {} chars -> {}".format(len(full), args.out))
        return

    if args.no_fulltext:
        data.pop("full_text", None)
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print("[{}] {} pages / {} captions / {} chars -> {}".format(
        engine, len(pages), len(data["captions"]), data["chars"], args.out))
    print("title   : {}".format(title[:110]))
    print("journal : {} {}".format(journal[:90], year))
    print("doi     : {}".format(doi))


if __name__ == "__main__":
    main()

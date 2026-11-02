# -*- coding: utf-8 -*-
"""从论文 PDF 里抠图 —— 首页截图、正文插图、局部放大。

用法：
    # 1) 首页整页截图（规范②要的第二页内容）
    python crop_figures.py paper.pdf --page1 2 -o figs/

    # 2) 列出每页的图片块，看看哪里有大图
    python crop_figures.py paper.pdf --list

    # 3) 按页 + 区域裁一张（坐标用 0~1 的相对值，好写）
    python crop_figures.py paper.pdf --crop 3,0.05,0.42,0.95,0.86 -o figs/fig1.png

    # 4) 把某页里所有够大的图片块单独抠出来
    python crop_figures.py paper.pdf --blocks 3 -o figs/

    # 5) 一页拆成上下两半（一张图有太多子图，拆页用）
    python crop_figures.py paper.pdf --split 4 -o figs/

坐标说明（--crop）：
    page,x0,y0,x1,y1   都是 0~1 相对值，相对整页宽高
    (0,0) 是左上角，(1,1) 是右下角

为什么要 0~1 相对坐标：PDF 页面尺寸千差万别（A4 / Letter / 双栏），
写死毫米数换个期刊就全错。相对坐标跟着页面走。
"""
from __future__ import annotations

import argparse
import os
import sys

# Windows 控制台默认 gbk，打印中文会炸。强制 UTF-8。
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:  # noqa: BLE001
        pass

try:
    import pymupdf as fitz
except ImportError:  # 旧版本叫 fitz
    try:
        import fitz  # type: ignore
    except ImportError:
        sys.exit("需要 PyMuPDF：pip install pymupdf")

MIN_BLOCK_IN = 1.2      # 图片块短边小于这个英寸就当装饰/logo 丢掉
DPI_COVER = 200         # 首页截图
DPI_FIG = 300           # 正文插图


def _page(pdf, idx_1based):
    if not 1 <= idx_1based <= len(pdf):
        sys.exit("页码越界：{} （共 {} 页）".format(idx_1based, len(pdf)))
    return pdf[idx_1based - 1]


def _save_pix(pix, path):
    os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)
    pix.save(path)
    print("saved {}  {}x{}".format(path, pix.width, pix.height))


def page_snapshot(pdf, page_no, out_dir, dpi=DPI_COVER):
    pg = _page(pdf, page_no)
    pix = pg.get_pixmap(dpi=dpi)
    path = os.path.join(out_dir, "page{:02d}.png".format(page_no))
    _save_pix(pix, path)
    return path


def crop_region(pdf, page_no, x0, y0, x1, y1, out_path, dpi=DPI_FIG):
    pg = _page(pdf, page_no)
    r = pg.rect
    clip = fitz.Rect(r.x0 + x0 * r.width, r.y0 + y0 * r.height,
                     r.x0 + x1 * r.width, r.y0 + y1 * r.height)
    if clip.is_empty:
        sys.exit("裁切区域无效：{}".format((x0, y0, x1, y1)))
    pix = pg.get_pixmap(dpi=dpi, clip=clip)
    _save_pix(pix, out_path)
    return out_path


def list_blocks(pdf):
    for i, pg in enumerate(pdf, start=1):
        r = pg.rect
        imgs = pg.get_images(full=True)
        big = []
        for b in pg.get_image_info(xrefs=True):
            bb = fitz.Rect(b["bbox"])
            w_in, h_in = bb.width / 72, bb.height / 72
            if min(w_in, h_in) >= MIN_BLOCK_IN:
                big.append((bb, w_in, h_in))
        # 版面也扫一遍文本块里的图注，方便配对
        captions = [ln for ln in pg.get_text("text").splitlines()
                    if ln.strip().lower().startswith(("fig", "scheme", "table"))]
        if big or captions:
            print("p{:<3} 图片块 {}  图注行 {}".format(i, len(big), len(captions)))
            for bb, w_in, h_in in big:
                print("       x{:.2f}-{:.2f} y{:.2f}-{:.2f}  {:.2f}×{:.2f}in".format(
                    bb.x0 / r.width, bb.x1 / r.width,
                    bb.y0 / r.height, bb.y1 / r.height, w_in, h_in))
            for c in captions[:3]:
                print("       cap: " + c.strip()[:90])


def extract_blocks(pdf, page_no, out_dir, dpi=DPI_FIG):
    """把该页够大的图片块逐个抠出来。"""
    pg = _page(pdf, page_no)
    out = []
    for k, b in enumerate(pg.get_image_info(xrefs=True), start=1):
        bb = fitz.Rect(b["bbox"])
        if min(bb.width, bb.height) / 72 < MIN_BLOCK_IN:
            continue
        pix = pg.get_pixmap(dpi=dpi, clip=bb)
        path = os.path.join(out_dir, "p{:02d}_img{:02d}.png".format(page_no, k))
        _save_pix(pix, path)
        out.append(path)
    if not out:
        print("p{} 没有够大的图片块（可能整页是一张大的矢量图）".format(page_no))
    return out


def split_page(pdf, page_no, out_dir, n=2, overlap=0.04, dpi=DPI_FIG):
    """纵向均分，带少量重叠防切字。"""
    pg = _page(pdf, page_no)
    out = []
    step = 1.0 / n
    for k in range(n):
        y0 = max(0.0, k * step - (overlap if k else 0))
        y1 = min(1.0, (k + 1) * step + (overlap if k < n - 1 else 0))
        path = os.path.join(out_dir, "p{:02d}_part{}.png".format(page_no, k + 1))
        crop_region(pdf, page_no, 0.03, y0, 0.97, y1, path, dpi=dpi)
        out.append(path)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("-o", "--out", default="figs")
    ap.add_argument("--page1", type=int, help="导出整页截图（给第二页用）")
    ap.add_argument("--list", action="store_true", help="列出图片块与图注")
    ap.add_argument("--crop", help="page,x0,y0,x1,y1 → -o 作为文件路径")
    ap.add_argument("--blocks", type=int, help="抠出该页所有大图块")
    ap.add_argument("--split", type=int, help="该页纵向拆成 N 张")
    ap.add_argument("--parts", type=int, default=2, help="配合 --split")
    ap.add_argument("--dpi", type=int, default=DPI_FIG)
    args = ap.parse_args()

    pdf = fitz.open(args.pdf)

    if args.list:
        list_blocks(pdf)
        return

    if args.page1:
        page_snapshot(pdf, args.page1, args.out, dpi=max(args.dpi, DPI_COVER))
        return

    if args.crop:
        try:
            p, x0, y0, x1, y1 = args.crop.split(",")
            page_no, vals = int(p), (float(x0), float(y0), float(x1), float(y1))
        except ValueError:
            sys.exit("--crop 格式：page,x0,y0,x1,y1（0~1 相对坐标）")
        out_path = args.out
        if not out_path.lower().endswith((".png", ".jpg", ".jpeg")):
            os.makedirs(out_path, exist_ok=True)
            out_path = os.path.join(out_path, "p{:02d}_crop.png".format(page_no))
        crop_region(pdf, page_no, *vals, out_path, dpi=args.dpi)
        return

    if args.blocks:
        extract_blocks(pdf, args.blocks, args.out, dpi=args.dpi)
        return

    if args.split:
        split_page(pdf, args.split, args.out, n=args.parts, dpi=args.dpi)
        return

    ap.print_help()


if __name__ == "__main__":
    main()

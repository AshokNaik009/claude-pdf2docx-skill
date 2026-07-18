#!/usr/bin/env python3
"""Extract text spans, tables, and vector drawings from a PDF into JSON.

Usage:
    python extract_layout.py input.pdf [--page N] [--out layout.json]

    --page is 1-indexed; default is all pages.

Span fields: text, font, size, bold, italic, color, x, y, x1, y1,
    block (PyMuPDF block id — group by it on multi-column pages),
    in_table (present+true when the span lies inside a table bbox;
    render such spans ONLY via the table, never as body text).
Table fields: bbox, rows (null cell = merged continuation or empty —
    infer colspan/rowspan from nulls), nested_in (index of the table
    this one sits inside, e.g. a grid within a quadrant layout).

Requires: pymupdf, pdfplumber
    pip install pymupdf pdfplumber --break-system-packages
"""
import argparse
import json
import sys

import fitz  # pymupdf
import pdfplumber

BOLD_FLAG = 2 ** 4
ITALIC_FLAG = 2 ** 1


def int_to_hex(color):
    """PyMuPDF span colors are packed ints (0xRRGGBB)."""
    try:
        return "#{:06x}".format(int(color))
    except (TypeError, ValueError):
        return "#000000"


def rgb_to_hex(color):
    """PyMuPDF drawing colors are float tuples in 0..1."""
    if not color:
        return None
    r, g, b = (max(0, min(255, round(v * 255))) for v in color[:3])
    return "#{:02X}{:02X}{:02X}".format(r, g, b)


def extract_spans(page):
    spans = []
    for block_id, block in enumerate(page.get_text("dict")["blocks"]):
        if block.get("type") != 0:  # 0 = text block
            continue
        for line in block["lines"]:
            for s in line["spans"]:
                text = s["text"].strip()
                if not text:
                    continue
                spans.append({
                    "text": text,
                    "font": s["font"],
                    "size": round(s["size"], 1),
                    "bold": "Bold" in s["font"] or bool(s["flags"] & BOLD_FLAG),
                    "italic": bool(s["flags"] & ITALIC_FLAG),
                    "color": int_to_hex(s["color"]),
                    "x": round(s["bbox"][0]),
                    "y": round(s["bbox"][1]),
                    "x1": round(s["bbox"][2]),
                    "y1": round(s["bbox"][3]),
                    # PyMuPDF block id: spans sharing a block belong together.
                    # On multi-column pages, group by block — the flat (y, x)
                    # sort interleaves side-by-side columns.
                    "block": block_id,
                })
    spans.sort(key=lambda r: (r["y"], r["x"]))
    return spans


def extract_tables(plumber_page):
    """null cells = merged (continuation of a rowspan/colspan) or empty."""
    tables = []
    for table in plumber_page.find_tables():
        tables.append({
            "bbox": [round(v) for v in table.bbox],
            "rows": table.extract(),
        })
    # A table whose bbox sits inside another's is nested (e.g. a financial
    # grid inside a quadrant-layout cell).
    for i, t in enumerate(tables):
        for j, outer in enumerate(tables):
            if i == j:
                continue
            ob, tb = outer["bbox"], t["bbox"]
            if (ob[0] <= tb[0] + 2 and ob[1] <= tb[1] + 2
                    and ob[2] >= tb[2] - 2 and ob[3] >= tb[3] - 2):
                t["nested_in"] = j
                break
    return tables


def tag_spans_in_tables(spans, tables):
    """Mark spans whose center falls inside a table bbox, so the renderer
    never emits them twice (once as body text, once via the table grid)."""
    n = 0
    for s in spans:
        cx = (s["x"] + s["x1"]) / 2
        cy = (s["y"] + s["y1"]) / 2
        for t in tables:
            b = t["bbox"]
            if b[0] <= cx <= b[2] and b[1] <= cy <= b[3]:
                s["in_table"] = True
                n += 1
                break
    return n


def extract_drawings(page):
    """Horizontal rules and filled rects — these carry design intent
    (colored dividers, shaded header bands, etc.)."""
    drawings = []
    for d in page.get_drawings():
        stroke = rgb_to_hex(d.get("color"))
        fill = rgb_to_hex(d.get("fill"))
        color = stroke or fill
        if color is None or color == "#FFFFFF":  # skip background fills
            continue
        rect = d["rect"]
        is_thin = (rect.height <= 3) or (rect.width <= 3)
        drawings.append({
            "type": "line" if is_thin else "rect",
            "color": color,
            "width": round(d.get("width") or 0, 1),
            "rect": [round(rect.x0), round(rect.y0), round(rect.x1), round(rect.y1)],
        })
    return drawings


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("pdf")
    ap.add_argument("--page", type=int, default=None, help="1-indexed single page")
    ap.add_argument("--out", default="layout.json")
    args = ap.parse_args()

    doc = fitz.open(args.pdf)
    plumber = pdfplumber.open(args.pdf)

    if args.page is not None:
        if not (1 <= args.page <= len(doc)):
            sys.exit("page {} out of range (1-{})".format(args.page, len(doc)))
        page_indexes = [args.page - 1]
    else:
        page_indexes = range(len(doc))

    pages = []
    for i in page_indexes:
        page = doc[i]
        spans = extract_spans(page)
        tables = extract_tables(plumber.pages[i])
        drawings = extract_drawings(page)
        in_table = tag_spans_in_tables(spans, tables)
        nested = sum(1 for t in tables if "nested_in" in t)
        pages.append({
            "page_number": i + 1,
            "width": round(page.rect.width, 1),
            "height": round(page.rect.height, 1),
            "spans": spans,
            "tables": tables,
            "drawings": drawings,
            "has_text": len(spans) > 0,
        })
        print("page {}: {} spans ({} in tables), {} table{}{}, {} drawing{}{}".format(
            i + 1, len(spans), in_table,
            len(tables), "" if len(tables) == 1 else "s",
            " ({} nested)".format(nested) if nested else "",
            len(drawings), "" if len(drawings) == 1 else "s",
            "" if spans else "  [NO TEXT - scanned page? needs OCR]",
        ))

    with open(args.out, "w") as f:
        json.dump({"source": args.pdf, "pages": pages}, f, indent=1)
    print("wrote {}".format(args.out))

    plumber.close()
    doc.close()


if __name__ == "__main__":
    main()

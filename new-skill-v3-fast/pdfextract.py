#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pdfextract.py — deterministic PDF page extraction for the pdf-to-docx pipeline.

Python 3.9.6 compatible. Uses PyMuPDF (fitz) — the SAME MuPDF engine as the
Node `mupdf` package the pipeline already uses (src/pdf.js), so its structured
output matches `pageStructured()` block-for-block. pdfplumber is used only as an
optional cross-check for table geometry (never required).

WHY THIS EXISTS
  The vision pipeline sends each complex page as a PNG to `claude -p` (expensive
  vision tokens). This script extracts the page's text/geometry deterministically
  so that:
    - CLEAN pages skip the LLM entirely — the JSON here feeds the pipeline's
      existing deterministic builder (src/extract.js buildPageSections), 0 tokens.
    - COMPLEX pages send a compact TEXT description (--text-prompt) to claude -p
      instead of the image — an order of magnitude fewer tokens per page.

OUTPUT SHAPE (matches mupdf.js toStructuredText(...).asJSON())
  { "blocks": [
      { "type": "text",
        "bbox": {"x","y","w","h"},
        "lines": [ { "wmode":0, "bbox":{"x","y","w","h","flags"},
                     "font": {"name","family","weight","style","size"},
                     "x","y", "text" } ] },
      { "type": "image", "bbox": {"x","y","w","h"} } ] }

USAGE
  python3 pdfextract.py file.pdf                      # human summary of all pages
  python3 pdfextract.py file.pdf --page 1 --emit both # {size, structured} for page 1 (JSON)  <- Node seam
  python3 pdfextract.py file.pdf --page 1 --emit structured
  python3 pdfextract.py file.pdf --page 1 --text-prompt   # token-cheap text for claude -p
  python3 pdfextract.py file.pdf --all --outdir work/<name>/pages   # page_0001.json ...
"""
import sys, os, json, argparse

try:
    import fitz  # PyMuPDF
except ImportError:
    sys.stderr.write("pdfextract.py requires PyMuPDF: pip install pymupdf\n")
    sys.exit(2)

# fitz span flag bits (see PyMuPDF docs: TextPage span "flags").
FLAG_SUPERSCRIPT = 1 << 0
FLAG_ITALIC      = 1 << 1
FLAG_SERIF       = 1 << 2
FLAG_MONO        = 1 << 3
FLAG_BOLD        = 1 << 4

# Match mupdf.js: "preserve-whitespace" plus keep ligatures, and surface image
# blocks so they can be cropped downstream exactly as the Node path does.
TEXT_FLAGS = (fitz.TEXT_PRESERVE_WHITESPACE
              | fitz.TEXT_PRESERVE_LIGATURES
              | fitz.TEXT_PRESERVE_IMAGES)


def _bbox(t):
    """fitz (x0,y0,x1,y1) tuple -> {x,y,w,h} like mupdf.js. Rounded to 0.01pt."""
    x0, y0, x1, y1 = t
    return {"x": round(x0, 2), "y": round(y0, 2),
            "w": round(x1 - x0, 2), "h": round(y1 - y0, 2)}


def _line_font(spans):
    """One font per line, like mupdf.js — which reports the size of the line's
    real-text span. Pick the span carrying the most NON-WHITESPACE characters:
    trailing/leading whitespace is often emitted as its own span at a different
    size (e.g. a heading's trailing space at 11pt), and choosing by raw size or
    first-span would misattribute the whole line's chars to that stray size and
    corrupt body-size / heading detection downstream. Ties fall to the larger
    size, then the earlier span. Bold/italic also sniff the font name, since some
    PDFs omit the style flag bits."""
    def score(s):
        return (len((s.get("text") or "").strip()), s.get("size", 0.0))
    dom = max(spans, key=score)
    flags = int(dom.get("flags", 0))
    name = dom.get("font", "") or ""
    lname = name.lower()
    bold = bool(flags & FLAG_BOLD) or "bold" in lname or "black" in lname or "heavy" in lname
    italic = bool(flags & FLAG_ITALIC) or "italic" in lname or "oblique" in lname
    serif = bool(flags & FLAG_SERIF)
    return {
        "name": name,
        "family": "serif" if serif else "sans-serif",
        "weight": "bold" if bold else "normal",
        "style": "italic" if italic else "normal",
        "size": round(float(dom.get("size", 0.0)), 2),
    }


def _line(line):
    spans = line.get("spans", [])
    if not spans:
        return None
    text = "".join(s.get("text", "") for s in spans)
    ox, oy = spans[0].get("origin", (line["bbox"][0], line["bbox"][3]))
    bb = _bbox(line["bbox"])
    bb["flags"] = 0  # mupdf.js stamps a flags field on the line bbox
    return {
        "wmode": int(line.get("wmode", 0)),
        "bbox": bb,
        "font": _line_font(spans),
        "x": round(ox, 2),
        "y": round(oy, 2),
        "text": text,
    }


def page_structured(page):
    """{ blocks:[...] } in mupdf.js shape for one fitz page."""
    raw = page.get_text("dict", flags=TEXT_FLAGS)
    blocks = []
    for b in raw.get("blocks", []):
        if b.get("type", 0) == 1:  # image block
            blocks.append({"type": "image", "bbox": _bbox(b["bbox"])})
            continue
        lines = []
        for ln in b.get("lines", []):
            out = _line(ln)
            if out is not None:
                lines.append(out)
        blocks.append({"type": "text", "bbox": _bbox(b["bbox"]), "lines": lines})
    return {"blocks": blocks}


def page_size(page):
    r = page.rect
    return {"width": round(r.width, 2), "height": round(r.height, 2)}


def _hex(c):
    """fitz color tuple (0..1 floats) -> #RRGGBB."""
    if not c:
        return None
    try:
        return "#%02X%02X%02X" % tuple(int(round(v * 255)) for v in c)
    except (TypeError, ValueError):
        return None


def page_drawings(page):
    """The VECTOR layer mupdf/pdfplumber both expose but the text layer drops:
    filled rects (section/header/cell backgrounds) with their EXACT hex color and
    the text they sit behind, a summary of border strokes, AND the outer content
    frame (the largest stroked rect that wraps most of the page). This is how the
    quadrant color bars, table borders and the page frame are recovered
    deterministically — no image, no model guessing. These facts are INJECTED by
    the renderer (src/docx.js renderLayoutJson), never left for the model to
    decide, because the model reliably drops the frame when asked to infer it.
    Returns {fills:[...], borders:[...], frame:{...}|None}."""
    from collections import Counter
    words = page.get_text("words")  # (x0,y0,x1,y1,word,...)
    pw, ph = page.rect.width, page.rect.height
    fills, border_counts = [], Counter()
    frame = None            # largest near-full-page stroked rect
    frame_area = 0.0
    for d in page.get_drawings():
        r = d["rect"]
        rw, rh = r.x1 - r.x0, r.y1 - r.y0
        if d.get("fill"):
            # Skip a full-page background rect (usually white) — it's not a fill
            # anyone styles a cell with.
            if rw >= pw - 2 and rh >= ph - 2:
                pass
            else:
                covered = " ".join(w[4] for w in words
                                   if r.x0 - 1 <= w[0] and w[2] <= r.x1 + 1
                                   and r.y0 - 1 <= w[1] and w[3] <= r.y1 + 1).strip()
                fills.append({"fill": _hex(d["fill"]),
                              "bbox": {"x": round(r.x0, 2), "y": round(r.y0, 2),
                                       "w": round(rw, 2), "h": round(rh, 2)},
                              "text": covered[:60]})
        if d.get("color"):
            width = round(d.get("width") or 0, 2)
            border_counts[(_hex(d["color"]), width)] += 1
            # An outer frame is a stroked rect spanning most of the page (>=60%
            # of both dimensions) but NOT the full-bleed page edge. Keep the
            # biggest such rect — that's the content border the model keeps losing.
            if 0.60 * pw <= rw < pw - 1 and 0.60 * ph <= rh < ph - 1 and rw * rh > frame_area:
                frame_area = rw * rh
                frame = {"color": _hex(d["color"]), "width": width,
                         "bbox": {"x": round(r.x0, 2), "y": round(r.y0, 2),
                                  "w": round(rw, 2), "h": round(rh, 2)}}
    borders = [{"color": c, "width": w, "count": n} for (c, w), n in border_counts.most_common()]
    return {"fills": fills, "borders": borders, "frame": frame}


# ------------------------------------------------------- token-cheap text prompt
def text_prompt(structured, size, drawings=None):
    """A compact, reading-order text rendering of the page for `claude -p` to turn
    into docx.js code — the deterministic stand-in for the page PNG. Groups short
    aligned rows as pipe-separated table rows, keeps font size so the model can
    tell headings from body, and (when `drawings` is supplied) hands the model the
    EXACT fill colors + border spec from the vector layer so section bars, header
    shading and table gridlines are reproduced, not guessed. A few hundred tokens
    instead of a raster."""
    blocks = [b for b in structured["blocks"] if b.get("type") == "text"]
    # reading order: top-to-bottom, then left-to-right (same tolerance as src/extract.js)
    blocks.sort(key=lambda b: (round(b["bbox"]["y"] / 4), b["bbox"]["x"]))
    body = _body_size(structured)
    out = ["PAGE %.0fx%.0f pt (body text ~%.1fpt). Reading order:" % (size["width"], size["height"], body)]
    for b in blocks:
        for ln in b.get("lines", []):
            t = ln["text"].rstrip()
            if not t.strip():
                continue
            sz = ln["font"]["size"]
            tag = ""
            if sz >= body * 1.15:
                tag = " [HEADING ~%.0fpt]" % sz
            elif ln["font"]["weight"] == "bold":
                tag = " [bold]"
            out.append("  x=%-5.0f y=%-5.0f | %s%s" % (ln["x"], ln["y"], t, tag))
    if drawings and drawings.get("fills"):
        out.append("")
        out.append("FILLS — exact background colors from the PDF (apply as the cell/section shading fill,")
        out.append("with contrasting text; these are NOT guesses):")
        for f in drawings["fills"]:
            bb = f["bbox"]
            out.append("  %s  at x=%-4.0f y=%-4.0f w=%-4.0f h=%-3.0f  behind: \"%s\""
                       % (f["fill"], bb["x"], bb["y"], bb["w"], bb["h"], f["text"]))
    if drawings and drawings.get("borders"):
        spec = ", ".join("%s %.1fpt (x%d)" % (b["color"], b["width"], b["count"]) for b in drawings["borders"])
        out.append("BORDERS present in the PDF (draw table gridlines to match): " + spec)
    return "\n".join(out)


def _body_size(structured):
    counts = {}
    for b in structured["blocks"]:
        for ln in b.get("lines", []):
            s = ln["font"]["size"]
            counts[s] = counts.get(s, 0) + len(ln["text"])
    if not counts:
        return 11.0
    return max(counts.items(), key=lambda kv: kv[1])[0]


# ------------------------------------------------------------------------- main
def open_pdf(path):
    if not os.path.exists(path):
        sys.stderr.write("PDF not found: %s\n" % path)
        sys.exit(2)
    return fitz.open(path)


def main():
    ap = argparse.ArgumentParser(description="Deterministic PDF page extractor (PyMuPDF, mupdf.js-shaped).")
    ap.add_argument("pdf")
    ap.add_argument("--page", type=int, help="1-indexed page to emit")
    ap.add_argument("--emit", choices=["structured", "size", "both"], default="both",
                    help="what to print for --page (default: both)")
    ap.add_argument("--text-prompt", action="store_true",
                    help="print the compact text description for claude -p instead of JSON")
    ap.add_argument("--all", action="store_true", help="write every page as page_NNNN.json")
    ap.add_argument("--outdir", help="output directory for --all")
    ap.add_argument("--dump-all", action="store_true",
                    help="print a JSON array of every page's {page,size,structured} to stdout "
                         "(one process for the whole document — used by the Node --py-extract seam)")
    args = ap.parse_args()

    doc = open_pdf(args.pdf)

    if args.dump_all:
        pages = []
        for i in range(doc.page_count):
            page = doc.load_page(i)
            size = page_size(page)
            structured = page_structured(page)
            drawings = page_drawings(page)
            pages.append({"page": i + 1, "size": size, "structured": structured,
                          "drawings": drawings,
                          "text_prompt": text_prompt(structured, size, drawings)})
        json.dump(pages, sys.stdout, ensure_ascii=False)
        return

    if args.all:
        outdir = args.outdir or "."
        os.makedirs(outdir, exist_ok=True)
        for i in range(doc.page_count):
            page = doc.load_page(i)
            rec = {"page": i + 1, "size": page_size(page), "structured": page_structured(page)}
            with open(os.path.join(outdir, "page_%04d.json" % (i + 1)), "w") as f:
                json.dump(rec, f, ensure_ascii=False)
        print("wrote %d page file(s) to %s" % (doc.page_count, outdir))
        return

    if args.page:
        if args.page < 1 or args.page > doc.page_count:
            sys.stderr.write("page %d out of range (1..%d)\n" % (args.page, doc.page_count))
            sys.exit(2)
        page = doc.load_page(args.page - 1)
        structured = page_structured(page)
        size = page_size(page)
        if args.text_prompt:
            print(text_prompt(structured, size, page_drawings(page)))
        elif args.emit == "structured":
            json.dump(structured, sys.stdout, ensure_ascii=False)
        elif args.emit == "size":
            json.dump(size, sys.stdout, ensure_ascii=False)
        else:
            json.dump({"size": size, "structured": structured}, sys.stdout, ensure_ascii=False)
        return

    # No page selected: human summary of the whole document.
    print("=== %s : %d page(s) ===" % (os.path.basename(args.pdf), doc.page_count))
    for i in range(doc.page_count):
        page = doc.load_page(i)
        st = page_structured(page)
        sz = page_size(page)
        tblocks = [b for b in st["blocks"] if b["type"] == "text"]
        iblocks = [b for b in st["blocks"] if b["type"] == "image"]
        chars = sum(len(ln["text"]) for b in tblocks for ln in b["lines"])
        print("page %2d : %.0fx%.0f pt | %d text block(s), %d image(s), %d chars, body ~%.1fpt"
              % (i + 1, sz["width"], sz["height"], len(tblocks), len(iblocks), chars, _body_size(st)))


if __name__ == "__main__":
    main()

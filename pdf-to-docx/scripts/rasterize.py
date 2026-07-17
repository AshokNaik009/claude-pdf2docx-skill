#!/usr/bin/env python3
"""Rasterize PDF pages to JPGs for use as the visual ground-truth reference.

Usage:
    python rasterize.py input.pdf [--dpi 100] [--outdir pages/] [--prefix page]

Uses pdftoppm (Poppler, preinstalled in the Claude container). Falls back to
PyMuPDF rendering if pdftoppm is unavailable.

Note: pdftoppm zero-pads page numbers to the width of the page count
(page-1.jpg for a 3-page PDF, page-01.jpg ... page-12.jpg for a 12-page PDF).
"""
import argparse
import glob
import os
import shutil
import subprocess
import sys


def rasterize_pdftoppm(pdf, outdir, prefix, dpi):
    subprocess.run(
        ["pdftoppm", "-jpeg", "-r", str(dpi), pdf, os.path.join(outdir, prefix)],
        check=True,
    )


def rasterize_pymupdf(pdf, outdir, prefix, dpi):
    import fitz  # pymupdf

    doc = fitz.open(pdf)
    pad = len(str(len(doc)))  # match pdftoppm's zero-padding convention
    for i, page in enumerate(doc, start=1):
        path = os.path.join(outdir, "{}-{:0{}d}.jpg".format(prefix, i, pad))
        page.get_pixmap(dpi=dpi).save(path)
    doc.close()


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("pdf")
    ap.add_argument("--dpi", type=int, default=100)
    ap.add_argument("--outdir", default="pages")
    ap.add_argument("--prefix", default="page")
    args = ap.parse_args()

    os.makedirs(args.outdir, exist_ok=True)

    if shutil.which("pdftoppm"):
        rasterize_pdftoppm(args.pdf, args.outdir, args.prefix, args.dpi)
    else:
        print("pdftoppm not found, falling back to PyMuPDF", file=sys.stderr)
        rasterize_pymupdf(args.pdf, args.outdir, args.prefix, args.dpi)

    images = sorted(glob.glob(os.path.join(args.outdir, args.prefix + "*.jpg")))
    if not images:
        sys.exit("rasterize.py: no images were produced")
    print("Generated page images:")
    for img in images:
        print(img)


if __name__ == "__main__":
    main()

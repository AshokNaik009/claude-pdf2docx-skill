#!/usr/bin/env python
"""Assessment harness: run Docling on a PDF and dump every structured output
we might feed into docx.js, so we can judge it against our current page.html.

Usage: docling_extract.py <input.pdf> <outdir>
Writes into <outdir>/:
  doc.md          - export_to_markdown()  (reading-order text + tables + image refs)
  doc.html        - export_to_html()      (richer structure, best for pandoc)
  doc.json        - full DoclingDocument JSON (layout, bboxes, element types)
  images/         - extracted picture assets
  SUMMARY.txt     - element-type counts + per-page table/figure tallies
"""
import sys, json, time
from pathlib import Path
from collections import Counter

def main():
    pdf, outdir = Path(sys.argv[1]), Path(sys.argv[2])
    outdir.mkdir(parents=True, exist_ok=True)
    (outdir / "images").mkdir(exist_ok=True)

    from docling.document_converter import DocumentConverter, PdfFormatOption
    from docling.datamodel.base_models import InputFormat
    from docling.datamodel.pipeline_options import PdfPipelineOptions

    # Ask Docling to keep page images + picture crops so we can see what it isolates.
    opts = PdfPipelineOptions()
    opts.do_ocr = False  # trust the PDF's own text layer (9x faster, identical quality here)
    opts.generate_page_images = True
    opts.generate_picture_images = True
    opts.images_scale = 2.0
    conv = DocumentConverter(
        format_options={InputFormat.PDF: PdfFormatOption(pipeline_options=opts)}
    )

    t0 = time.time()
    result = conv.convert(str(pdf))
    dt = time.time() - t0
    doc = result.document

    (outdir / "doc.md").write_text(doc.export_to_markdown(), encoding="utf-8")
    (outdir / "doc.html").write_text(doc.export_to_html(), encoding="utf-8")
    (outdir / "doc.json").write_text(
        json.dumps(doc.export_to_dict(), indent=2, default=str), encoding="utf-8"
    )

    # Save picture assets.
    n_img = 0
    for i, pic in enumerate(doc.pictures):
        img = pic.get_image(doc)
        if img is not None:
            img.save(outdir / "images" / f"pic_{i+1}.png")
            n_img += 1

    # Structure summary.
    kinds = Counter(item.label.value if hasattr(item.label, "value") else str(item.label)
                    for item, _ in doc.iterate_items())
    lines = [
        f"pdf: {pdf.name}",
        f"convert_seconds: {dt:.1f}",
        f"pages: {len(doc.pages)}",
        f"tables: {len(doc.tables)}",
        f"pictures: {len(doc.pictures)}  (saved {n_img})",
        f"texts: {len(doc.texts)}",
        "",
        "element_type_counts:",
    ]
    for k, v in kinds.most_common():
        lines.append(f"  {k}: {v}")
    (outdir / "SUMMARY.txt").write_text("\n".join(lines) + "\n", encoding="utf-8")
    print("\n".join(lines))

if __name__ == "__main__":
    main()

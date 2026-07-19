import sys, time
from pathlib import Path
from docling.document_converter import DocumentConverter, PdfFormatOption
from docling.datamodel.base_models import InputFormat
from docling.datamodel.pipeline_options import PdfPipelineOptions
pdf, outdir = Path(sys.argv[1]), Path(sys.argv[2]); outdir.mkdir(parents=True, exist_ok=True)
opts = PdfPipelineOptions(); opts.do_ocr = False   # trust the PDF's own text layer
conv = DocumentConverter(format_options={InputFormat.PDF: PdfFormatOption(pipeline_options=opts)})
t0=time.time(); doc = conv.convert(str(pdf)).document; dt=time.time()-t0
(outdir/"doc.md").write_text(doc.export_to_markdown(), encoding="utf-8")
print(f"noocr convert_seconds={dt:.1f} pages={len(doc.pages)} tables={len(doc.tables)}")

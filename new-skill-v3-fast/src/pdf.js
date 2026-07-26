// Thin wrapper over mupdf.js (pure-WASM, no native build): per-page text,
// structured layout (blocks/lines/fonts/bboxes), full-page render, and region
// cropping. Coordinates from mupdf are in PDF points (72/inch).
import { readFileSync } from "node:fs";
import { execFileSync } from "node:child_process";
import * as mupdfNS from "mupdf";
import { encodePngRGB } from "./png.js";

const mu = mupdfNS.Document ? mupdfNS : (mupdfNS.default ?? mupdfNS);
if (!mu || !mu.Document) {
  throw new Error('Could not load "mupdf". Run `npm install` in this directory first.');
}

export function openPdf(path) {
  const data = readFileSync(path);
  return mu.Document.openDocument(new Uint8Array(data), "application/pdf");
}

export function pageCount(doc) {
  return doc.countPages();
}

/** Plain text layer of page i (0-indexed). Empty for image-only pages. */
export function pageText(doc, i) {
  return doc.loadPage(i).toStructuredText().asText();
}

/** Parsed structured layout: { blocks:[{type,bbox,lines:[{font,x,y,text,bbox}]}] } */
export function pageStructured(doc, i) {
  const json = doc.loadPage(i).toStructuredText("preserve-whitespace").asJSON();
  return JSON.parse(json);
}

/**
 * Deterministic Python-side extraction (--py-extract). Shells out to
 * pdfextract.py (PyMuPDF — the same MuPDF engine, so its structured output
 * matches pageStructured() block-for-block, verified against the Node path).
 * This is the token-free replacement for the page image on the generation side:
 * clean pages feed buildPageSections directly, complex pages get a compact TEXT
 * prompt (pdfextract.py --text-prompt) instead of a raster.
 * @returns {{ structured: object, width: number, height: number }}
 */
const _pyExtractCache = new Map(); // pdfPath -> [{page,size,structured}, ...]

export function pageStructuredPy(pdfPath, i, cfg) {
  let pages = _pyExtractCache.get(pdfPath);
  if (!pages) {
    // One Python process for the WHOLE document (fitz import + PDF open paid
    // once, not per page), then cache — concurrent page workers all read it.
    const bin = cfg.pyExtractBin || "python3";
    const script = cfg.pyExtractScript || "pdfextract.py";
    const out = execFileSync(bin, [script, pdfPath, "--dump-all"], { maxBuffer: 256 * 1024 * 1024 });
    pages = JSON.parse(out.toString());
    _pyExtractCache.set(pdfPath, pages);
  }
  const rec = pages[i];
  return { structured: rec.structured, width: rec.size.width, height: rec.size.height,
           textPrompt: rec.text_prompt, drawings: rec.drawings };
}

/** Page size in points: { width, height }. */
export function pageSize(doc, i) {
  const b = doc.loadPage(i).getBounds(); // [x0,y0,x1,y1]
  return { width: b[2] - b[0], height: b[3] - b[1] };
}

/**
 * Render a page once into a reusable raster object. `fullPng` is the whole
 * page; `crop(bbox)` slices any sub-region straight from the pixel buffer.
 */
export function renderPage(doc, i, dpi) {
  const page = doc.loadPage(i);
  const scale = dpi / 72;
  const pix = page.toPixmap(mu.Matrix.scale(scale, scale), mu.ColorSpace.DeviceRGB, false);
  const width = pix.getWidth();
  const height = pix.getHeight();
  const ncomp = pix.getNumberOfComponents();
  const stride = pix.getStride();
  const fullPng = Buffer.from(pix.asPNG());

  // getPixels() returns a live view into mupdf's WASM linear memory. Any
  // later mupdf call that grows/reallocates that memory (asPNG() above does)
  // detaches the old view — its .length silently becomes 0 and every read is
  // `undefined`, producing an all-black crop with no error. Fetching pixels
  // fresh on every crop() call (never caching across other mupdf calls) is
  // what actually avoids this, since we can't assume no other mupdf work
  // happens between renderPage() returning and crop() being called.
  function crop(bboxPts) {
    const pixels = pix.getPixels();
    const x0 = Math.max(0, Math.floor(bboxPts.x * scale));
    const y0 = Math.max(0, Math.floor(bboxPts.y * scale));
    const x1 = Math.min(width, Math.ceil((bboxPts.x + bboxPts.w) * scale));
    const y1 = Math.min(height, Math.ceil((bboxPts.y + bboxPts.h) * scale));
    const cw = Math.max(1, x1 - x0);
    const ch = Math.max(1, y1 - y0);
    const out = Buffer.alloc(cw * ch * 3);
    for (let y = 0; y < ch; y++) {
      for (let x = 0; x < cw; x++) {
        const si = (y0 + y) * stride + (x0 + x) * ncomp;
        const di = (y * cw + x) * 3;
        out[di] = pixels[si];
        out[di + 1] = pixels[si + 1];
        out[di + 2] = pixels[si + 2];
      }
    }
    return { png: encodePngRGB(cw, ch, out), width: cw, height: ch };
  }

  return { width, height, scale, fullPng, crop };
}

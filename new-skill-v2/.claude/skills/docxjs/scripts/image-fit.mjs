/**
 * Image sizing for docx.js page pipelines.
 *
 * WHY: `ImageRun` requires explicit pixel width AND height. A model looking at a page image can
 * only guess them, and a wrong guess is invisible — a 600x50 squash packs successfully and
 * produces a distorted figure that nobody catches in a 20-page batch.
 *
 * This module reads each asset's intrinsic size from its file header (no dependencies) and scales
 * it to the available width with the aspect ratio preserved. The generated code then never
 * specifies dimensions at all: `ctx.image("fig-2.png")` is enough.
 *
 * Usage in the host:
 *   import { makeImageCtx } from "./image-fit.mjs";
 *   const ctx = makeImageCtx(docx, { imageDir, pageWidthTwips: 12240, marginTwips: 1440 });
 *   const { sections, styles } = buildPage(docx, ctx);
 */
import { readFileSync, existsSync } from "node:fs";
import { join, extname } from "node:path";

const TWIPS_PER_INCH = 1440;
const PX_PER_INCH = 96; // Word treats docx.js "pixels" as 96dpi
export const twipsToPx = (t) => Math.round((t / TWIPS_PER_INCH) * PX_PER_INCH);

const EXT_TYPE = { ".png": "png", ".jpg": "jpg", ".jpeg": "jpg", ".gif": "gif", ".bmp": "bmp" };

/**
 * Read intrinsic pixel dimensions from a PNG/JPEG/GIF/BMP buffer.
 * @returns {{width:number,height:number}|null}
 */
export function imageSize(buf) {
  if (!buf || buf.length < 24) return null;
  // PNG: 8-byte signature, then IHDR with width/height as big-endian uint32
  if (buf[0] === 0x89 && buf[1] === 0x50 && buf[2] === 0x4e && buf[3] === 0x47) {
    return { width: buf.readUInt32BE(16), height: buf.readUInt32BE(20) };
  }
  // GIF: "GIF87a"/"GIF89a", little-endian uint16 at 6 and 8
  if (buf[0] === 0x47 && buf[1] === 0x49 && buf[2] === 0x46) {
    return { width: buf.readUInt16LE(6), height: buf.readUInt16LE(8) };
  }
  // BMP: "BM", int32 LE at 18 and 22 (height may be negative for top-down)
  if (buf[0] === 0x42 && buf[1] === 0x4d) {
    return { width: buf.readInt32LE(18), height: Math.abs(buf.readInt32LE(22)) };
  }
  // JPEG: walk the segment markers to a SOFn frame header
  if (buf[0] === 0xff && buf[1] === 0xd8) {
    let off = 2;
    while (off < buf.length - 9) {
      if (buf[off] !== 0xff) { off++; continue; }
      const marker = buf[off + 1];
      // SOF0..SOF15 except DHT(c4), JPGA(c8), DAC(cc) carry the frame size
      if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
        return { height: buf.readUInt16BE(off + 5), width: buf.readUInt16BE(off + 7) };
      }
      off += 2 + buf.readUInt16BE(off + 2);
    }
  }
  return null;
}

/**
 * Scale intrinsic dimensions to fit a max width (and optional max height), preserving ratio.
 */
export function fitTo(intrinsic, maxWidthPx, maxHeightPx) {
  if (!intrinsic || !intrinsic.width || !intrinsic.height) {
    return { width: maxWidthPx, height: Math.round(maxWidthPx * 0.62) }; // sane fallback
  }
  let w = intrinsic.width;
  let h = intrinsic.height;
  if (w > maxWidthPx) { h = Math.round(h * (maxWidthPx / w)); w = maxWidthPx; }
  if (maxHeightPx && h > maxHeightPx) { w = Math.round(w * (maxHeightPx / h)); h = maxHeightPx; }
  return { width: w, height: h };
}

/**
 * Build the `ctx` object handed to buildPage(docx, ctx).
 *
 * ctx.image(file, opts?) -> docx.ImageRun sized to fit.
 *   opts.fit    "page" (default) | "column" | "half"
 *   opts.widthPx / opts.maxHeightPx  explicit overrides when the caller really knows better
 *
 * @param {object} docx  the docx module
 * @param {object} o
 * @param {string} o.imageDir
 * @param {number} [o.pageWidthTwips=12240]   Letter; A4 = 11906
 * @param {number} [o.marginTwips=1440]
 * @param {number} [o.columnCount=1]
 * @param {number} [o.columnSpaceTwips=560]
 * @param {number} [o.maxHeightPx=620]        keeps a tall figure from pushing off the page
 * @param {(msg:string)=>void} [o.onWarn]
 */
export function makeImageCtx(docx, o = {}) {
  const {
    imageDir = ".",
    pageWidthTwips = 12240,
    marginTwips = 1440,
    columnCount = 1,
    columnSpaceTwips = 560,
    maxHeightPx = 620,
    onWarn = () => {},
  } = o;

  const contentTwips = pageWidthTwips - 2 * marginTwips;
  const colTwips = columnCount > 1
    ? Math.floor((contentTwips - columnSpaceTwips * (columnCount - 1)) / columnCount)
    : contentTwips;

  const widths = {
    page: twipsToPx(contentTwips),
    column: twipsToPx(colTwips),
    half: Math.round(twipsToPx(contentTwips) / 2),
  };

  return {
    widths,
    image(file, opts = {}) {
      const path = join(imageDir, file);
      const type = EXT_TYPE[extname(file).toLowerCase()] || "png";
      if (!existsSync(path)) {
        onWarn(`image not found: ${file}`);
        return new docx.TextRun({ text: `[ missing image: ${file} ]`, italics: true, color: "999999" });
      }
      const data = readFileSync(path);
      const maxW = opts.widthPx || widths[opts.fit] || widths.page;
      const { width, height } = fitTo(imageSize(data), maxW, opts.maxHeightPx || maxHeightPx);
      return new docx.ImageRun({ type, data, transformation: { width, height } });
    },
  };
}

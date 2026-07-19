// MARKDOWN OUTPUT STRATEGY (--markdown).
// Assemble the whole ordered document as ONE pure-GFM Markdown string, then let
// the markdown-docx library turn it into the final .docx.
//
// Why a sanitizer: markdown-docx is a CommonMark+GFM converter with NO raw-HTML
// rendering — a stray <div>/<sup>/&nbsp; would land in the document as literal
// visible text. sanitizeGfm() downgrades the handful of inline tags a vision
// model tends to slip in (<b>/<i>/<sup>/<br>/&nbsp;) into their GFM/plain-text
// equivalents and drops block wrappers, so the output stays clean even when the
// generated Markdown isn't perfectly pure.
import { readFileSync } from "node:fs";
import { join, isAbsolute, basename } from "node:path";
import markdownDocx, { Packer } from "markdown-docx";
import { blocksToMarkdown } from "./markdown.js";

/** Downgrade the common HTML tags a model slips in to GFM / plain text. */
export function sanitizeGfm(md) {
  return md
    .replace(/&nbsp;/gi, " ")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/?(?:b|strong)>/gi, "**")
    .replace(/<\/?(?:i|em)>/gi, "*")
    .replace(/<sup>([^<]*)<\/sup>/gi, "$1")
    .replace(/<sub>([^<]*)<\/sub>/gi, "$1")
    .replace(/<\/?(?:div|p|span|u)[^>]*>/gi, "\n")
    .replace(/<\/?[a-z][^>]*>/gi, "") // any remaining stray tag
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

const MAX_IMG_W = 624; // 6.5in content width at 96dpi — keeps images within margins

/**
 * Inline each relative image as a base64 data URI: ![alt](data:...;base64,... "WxH").
 * The intrinsic pixel size becomes the title (downscaled to the content width) so
 * markdown-docx sizes it correctly. Absolute/remote/already-inlined srcs are left as-is.
 */
function inlineImages(md, assetsAbsDir) {
  return md.replace(/!\[([^\]]*)\]\(([^)\s]+)((?:\s+"[^"]*")?)\)/g, (m, alt, src, title) => {
    if (/^(?:https?:|data:|file:)/i.test(src) || isAbsolute(src)) return m;
    const abs = join(assetsAbsDir, basename(src));
    try {
      const data = readFileSync(abs);
      const meta = imageMeta(data);
      let { width, height } = meta;
      if (width > MAX_IMG_W) { height = Math.round((height * MAX_IMG_W) / width); width = MAX_IMG_W; }
      const mime = meta.type === "jpg" ? "image/jpeg" : `image/${meta.type}`;
      const uri = `data:${mime};base64,${data.toString("base64")}`;
      const sizeTitle = title && title.trim() ? title : ` "${width}x${height}"`;
      return `![${alt}](${uri}${sizeTitle})`;
    } catch {
      return `*[missing image: ${basename(src)}]*`; // asset not on disk — graceful note
    }
  });
}

/** Best-effort GFM for a single page (md-mode file, text-page blocks, or fallback). */
export function pageToMarkdown(workDir, page) {
  const assetsAbsDir = page.assetsDir ? join(workDir, page.assetsDir) : workDir;

  let md;
  if (page.markdownFile) {
    md = sanitizeGfm(readFileSync(join(workDir, page.markdownFile), "utf8"));
  } else {
    const blocks = page.sections ? page.sections.flatMap((s) => s.blocks || []) : page.blocks || [];
    md = blocks.length ? blocksToMarkdown(blocks) : (page.fallbackText || []).join("\n\n");
  }
  return inlineImages(md, assetsAbsDir);
}

/** Assemble the ordered document model into one GFM string. */
export function buildMarkdownDocument(workDir, model) {
  return model.pages.map((pg) => pageToMarkdown(workDir, pg)).join("\n\n");
}

// ---- Image dimensions (markdown-docx requires width/height) -----------------

function pngMeta(buf) {
  return { type: "png", width: buf.readUInt32BE(16), height: buf.readUInt32BE(20) };
}

/** Minimal JPEG SOF scan for intrinsic dimensions. */
function jpegMeta(buf) {
  let o = 2;
  while (o < buf.length) {
    if (buf[o] !== 0xff) { o++; continue; }
    const marker = buf[o + 1];
    // SOF0..SOF3, SOF5..SOF7, SOF9..SOF11, SOF13..SOF15 carry frame dimensions.
    if ((marker >= 0xc0 && marker <= 0xc3) || (marker >= 0xc5 && marker <= 0xc7) ||
        (marker >= 0xc9 && marker <= 0xcb) || (marker >= 0xcd && marker <= 0xcf)) {
      return { type: "jpg", height: buf.readUInt16BE(o + 5), width: buf.readUInt16BE(o + 7) };
    }
    o += 2 + buf.readUInt16BE(o + 2);
  }
  return { type: "jpg", width: 600, height: 400 };
}

function imageMeta(buf) {
  if (buf.length > 24 && buf.readUInt32BE(0) === 0x89504e47) return pngMeta(buf);
  if (buf.length > 4 && buf[0] === 0xff && buf[1] === 0xd8) return jpegMeta(buf);
  return { type: "png", width: 600, height: 400 };
}

/** Image adapter: decode the base64 data URIs inlineImages() produced. */
async function dataUriImageAdapter(token) {
  const m = /^data:([^;]+);base64,(.+)$/s.exec(token.href || "");
  if (!m) return null; // non-data src (shouldn't occur) → graceful placeholder
  try {
    const data = Buffer.from(m[2], "base64");
    return { data, ...imageMeta(data) };
  } catch {
    return null;
  }
}

/** Produce the final .docx Buffer from the assembled Markdown. */
export async function renderMarkdownDocx(markdown, opts = {}) {
  const doc = await markdownDocx(markdown, { gfm: true, imageAdapter: dataUriImageAdapter, ...opts });
  return Packer.toBuffer(doc);
}

// PANDOC HTML OUTPUT STRATEGY (--pandoc).
// Assemble the ordered document as ONE semantic-HTML string and let the pandoc
// CLI turn it into the final .docx. Pandoc PARSES HTML (unlike markdown-docx,
// which renders raw tags as literal text): headings, <strong>/<em>, lists,
// <a>, and <table> become real Word constructs, and <img> assets are embedded
// as real Word media parts. Pandoc ignores CSS, so layout that must survive
// (multiple columns, sidebars) is expressed as <table> cells, not flex/float.
import { readFileSync, writeFileSync, existsSync, mkdtempSync } from "node:fs";
import { join, isAbsolute, basename } from "node:path";
import { tmpdir } from "node:os";
import { execFile, execFileSync } from "node:child_process";
import { blocksToMarkdown } from "./markdown.js";

/** True if a usable `pandoc` binary is on PATH. */
export function pandocAvailable() {
  try {
    execFileSync("pandoc", ["--version"], { stdio: "ignore" });
    return true;
  } catch {
    return false;
  }
}

/** Escape a string for safe inclusion in HTML text/attribute context. */
function esc(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

/** Rewrite relative <img src> / <img data-src> to absolute on-disk paths pandoc can read & embed. */
function absolutizeImgs(html, assetsAbsDir) {
  return html.replace(/(<img\b[^>]*?\bsrc\s*=\s*)(["'])(.*?)\2/gi, (m, pre, q, src) => {
    if (/^(?:https?:|data:|file:)/i.test(src) || isAbsolute(src)) return m;
    return `${pre}${q}${join(assetsAbsDir, basename(src))}${q}`;
  });
}

/** Minimal GFM-blocks → HTML, for deterministic text pages that never hit the model. */
function blocksToHtml(blocks) {
  const out = [];
  for (const b of blocks || []) {
    switch (b.type) {
      case "heading":
        out.push(`<h${Math.min(6, b.level || 1)}>${esc(b.text)}</h${Math.min(6, b.level || 1)}>`);
        break;
      case "paragraph":
        out.push(`<p>${esc(b.text)}</p>`);
        break;
      case "list": {
        const tag = b.ordered ? "ol" : "ul";
        out.push(`<${tag}>` + b.items.map((it) => `<li>${esc(it)}</li>`).join("") + `</${tag}>`);
        break;
      }
      case "table":
        if (b.rows?.length) {
          const row = (cells, cell) => "<tr>" + cells.map((c) => `<${cell}>${esc(c)}</${cell}>`).join("") + "</tr>";
          out.push("<table>" + row(b.rows[0], "th") + b.rows.slice(1).map((r) => row(r, "td")).join("") + "</table>");
        }
        break;
      case "image":
        out.push(`<p><img src="${esc(basename(b.path))}" alt="${esc(b.alt || "")}"></p>`);
        break;
    }
  }
  return out.join("\n");
}

/** Best-effort HTML for a single page (model-generated file, text-page blocks, or fallback). */
export function pageToHtml(workDir, page) {
  const assetsAbsDir = page.assetsDir ? join(workDir, page.assetsDir) : workDir;

  let html;
  if (page.htmlFile) {
    html = readFileSync(join(workDir, page.htmlFile), "utf8");
  } else {
    const blocks = page.sections ? page.sections.flatMap((s) => s.blocks || []) : page.blocks || [];
    html = blocks.length
      ? blocksToHtml(blocks)
      : (page.fallbackText || []).map((t) => `<p>${esc(t)}</p>`).join("\n");
  }
  return absolutizeImgs(html, assetsAbsDir);
}

// A distinctive paragraph we drop between source pages. Pandoc's HTML reader
// turns it into a Para of one Str, which we swap for a real page-break block in
// the AST (below). We can't put the page break directly in the HTML because the
// HTML reader — the one that actually parses <table>/<img> — passes raw OpenXML
// through as inert text; only the docx WRITER honors an openxml RawBlock.
const PB_MARKER = "@@PDF2DOCX_PAGEBREAK@@";
const PB_RAWBLOCK = { t: "RawBlock", c: ["openxml", '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'] };

/** Assemble the ordered document model into one HTML string (pages split by markers). */
export function buildHtmlDocument(workDir, model) {
  return model.pages.map((pg) => pageToHtml(workDir, pg)).join(`\n<p>${PB_MARKER}</p>\n`);
}

/** Run pandoc, feeding `input` on stdin; resolve with captured stdout Buffer. */
function runPandoc(args, input) {
  return new Promise((resolve, reject) => {
    const child = execFile("pandoc", args, { encoding: "buffer", maxBuffer: 128 * 1024 * 1024 }, (err, stdout, stderr) => {
      if (err) {
        reject(new Error(`pandoc failed: ${err.message}${stderr ? ` | ${String(stderr).slice(0, 400)}` : ""}`));
        return;
      }
      resolve(stdout);
    });
    child.stdin.end(input);
  });
}

/** Plain text of a Para's inline content (enough to recognize the page-break marker). */
function inlineText(inlines) {
  return (inlines || []).map((n) => (n.t === "Str" ? n.c : n.t === "Space" ? " " : "")).join("");
}

/**
 * Produce the final .docx Buffer from assembled HTML via pandoc, in two stages:
 *   HTML → pandoc AST (JSON): parses <table>/<img>/<strong>/lists into native Word constructs
 *   AST → docx:               with each page-break marker replaced by a real OpenXML page break
 */
export async function renderPandocDocx(html, opts = {}) {
  const jsonBuf = await runPandoc(["-f", "html", "-t", "json"], html);
  const ast = JSON.parse(jsonBuf.toString("utf8"));
  ast.blocks = (ast.blocks || []).map((b) =>
    b.t === "Para" && inlineText(b.c).trim() === PB_MARKER ? PB_RAWBLOCK : b
  );

  const dir = mkdtempSync(join(tmpdir(), "pdf2docx-pandoc-"));
  const outPath = join(dir, "out.docx");
  const args = ["-f", "json", "-t", "docx", "-o", outPath];
  if (opts.referenceDoc && existsSync(opts.referenceDoc)) args.push("--reference-doc", opts.referenceDoc);
  if (opts.luaFilter && existsSync(opts.luaFilter)) args.push("--lua-filter", opts.luaFilter);
  await runPandoc(args, JSON.stringify(ast));
  return readFileSync(outPath);
}

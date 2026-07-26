// Headless Claude worker. Per the architecture, Claude LOOKS AT THE PAGE IMAGE
// and GENERATES the docx-library JavaScript that reconstructs the page. We do
// not parse Markdown — Claude is the code generator. The merge step executes
// the generated module and stitches pages into the final .docx.
import { execFile } from "node:child_process";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";
// Host-side extraction ships WITH the skill (its "host-side hardening", codegen.md §6):
// it strips fences AND prose the model may emit around the module — the failure that
// prompt rules alone can't eliminate. Reuse it instead of a weaker local sanitizer.
import { extractModule } from "../.claude/skills/docxjs/scripts/extract-code.mjs";

const HERE = dirname(fileURLToPath(import.meta.url));
const SKILL_DIR = resolve(HERE, "..", ".claude", "skills", "docxjs");

function buildPrompt(imageAbsPath, assetFiles, priorAttempt) {
  const assets = assetFiles && assetFiles.length
    ? assetFiles.map((f) => `"${f}"`).join(", ")
    : "(none)";
  const retryNote = priorAttempt ? [
    "",
    "RETRY — your previous attempt failed verification. Fix these problems; do not repeat them:",
    ...priorAttempt.errors.map((e) => `  - ${e}`),
    "",
    "Your previous module (for reference — rewrite it, don't just patch around the errors):",
    "```js",
    priorAttempt.code,
    "```",
  ] : [];
  return [
    "You reconstruct a single document page as editable Word content by writing JavaScript with the docx.js library.",
    "The docxjs skill below already contains every authoring rule you need — read it and follow it rather than guessing.",
    "",
    "STEP 1 — Read the docxjs skill with the Read tool, in this order:",
    `  ${SKILL_DIR}/references/codegen.md     (READ FIRST — you are generating code a host will execute. This is the skill's "Mode B": the buildPage return contract, the double-quote string-escaping that prevents SyntaxErrors while transcribing page text, output framing, the API mistakes that silently corrupt output, and a pre-emit self-check.)`,
    `  ${SKILL_DIR}/SKILL.md                   (building blocks + the reusable helpers — runsFromText auto-italicizes "et al.", bodyPara, heading, caption.)`,
    `  ${SKILL_DIR}/references/layout.md        (ONLY if the page mixes column counts, e.g. a full-width table/figure inside two-column text — sections and CONTINUOUS breaks.)`,
    `  ${SKILL_DIR}/references/tables.md        (ONLY if the page has a table — borders, spans, cell margins, numeric/decimal-aligned columns, totals rows.)`,
    `  ${SKILL_DIR}/references/images.md        (ONLY if the page has a figure/diagram — placement and the missing-asset placeholder rule.)`,
    "",
    `STEP 2 — Read the page image to reconstruct: ${imageAbsPath}`,
    "",
    "STEP 3 — Emit an ES module for THIS host. Its exact contract (the skill calls this \"Mode B\"):",
    "  export function buildPage(docx, ctx) { ...; return { sections, styles, header, footer, pageNumberStart }; }",
    "  - `sections` is an ARRAY of Word section objects, each { properties, children }.",
    "  - `styles` (optional) is a document-wide default; set it to the page's body font/size when you can identify it,",
    "    e.g. { default: { document: { run: { font: \"Palatino Linotype\", size: 22 } } } }.",
    "  - Put page size/margins on the FIRST section; give later sections type: docx.SectionType.CONTINUOUS so they don't start a new page.",
    "  - TABLE FIDELITY (you can SEE the styling in the image — reproduce what is actually there, don't invent it):",
    "    * Header row with a fill: set the cell `shading:{fill}` to the color you see, and set its TextRun color to the",
    "      contrasting text color you see (white/light text on a dark fill, dark text on a light fill). Never leave light text",
    "      on no fill — it becomes invisible in Word.",
    "    * Row-header column: if the LEFT column is labels/categories styled like a header (fill or bold) with data to its",
    "      right, style those left cells like the header — same fill/bold — not like plain body cells.",
    "    * Borders: if the table shows visible gridlines, draw them; if it is a clean borderless/editorial table, set borders",
    "      to NONE (see tables.md) — don't add a grid the original doesn't have.",
    "  - SHADED SECTION BAR: a heading shown as a full-width colored bar (e.g. a dark band with light text like \"THE COMPANY\")",
    "    is a one-cell full-width table (or a paragraph with `shading`) using the fill + contrasting text color you see — not a plain heading.",
    "  - RUNNING HEADER/FOOTER: text that recurs at the very top or bottom margin of every page (a journal masthead, author/title",
    "    running head, a page number, \"Page 3 of 20\", a confidentiality line) is FURNITURE, NOT body content. NEVER place it in",
    "    `children` — it would repeat mid-flow. Instead return it as `header` and/or `footer`:",
    "      header: new docx.Header({ children: [ new docx.Paragraph({ children: [ new docx.TextRun(\"…masthead…\") ] }) ] }),",
    "      footer: new docx.Footer({ children: [ new docx.Paragraph({ children: [ new docx.TextRun({ children: [docx.PageNumber.CURRENT] }) ] }) ] }),",
    "    Use docx.PageNumber.CURRENT / docx.PageNumber.TOTAL_PAGES as the LIVE page number — never hard-code the digit you see.",
    "    If the visible page number does not start at 1, set `pageNumberStart` to the number shown on THIS page minus (thisPageIndex-1).",
    "    These are document-wide: the host takes them from the first page that supplies them, so emit them even if some pages omit them.",
    "",
    "",
    "You have ONLY the Read tool — no tool that writes files. Do NOT narrate what you're doing or say things like",
    "\"I'll output the module directly instead of writing a file\"; any such sentence is a SyntaxError in the module.",
    "Output ONLY the JavaScript module — start at the first character of code, end at the last. No markdown fences,",
    "no commentary before or after (the skill's output-framing rule in codegen.md §3).",
    ...retryNote,
  ].join("\n");
}

/**
 * Text-only variant of buildPrompt: the model gets the page's deterministically
 * extracted text with x/y coordinates (PDF points, origin top-left) INSTEAD of
 * the page raster. No vision tokens, and the text/numbers are exact rather than
 * OCR'd. The trade-off is that layout (columns, table boundaries) must be
 * inferred from coordinates — so this prompt spells out how to read them.
 */
function buildTextPrompt(textPrompt, assetFiles, priorAttempt) {
  const retryNote = priorAttempt ? [
    "",
    "RETRY — your previous attempt failed verification. Fix these problems; do not repeat them:",
    ...priorAttempt.errors.map((e) => `  - ${e}`),
    "",
    "Your previous module (for reference — rewrite it, don't just patch around the errors):",
    "```js",
    priorAttempt.code,
    "```",
  ] : [];
  return [
    "You reconstruct a single document page as editable Word content by writing JavaScript with the docx.js library.",
    "The docxjs skill below already contains every authoring rule you need — read it and follow it rather than guessing.",
    "",
    "STEP 1 — Read the docxjs skill with the Read tool, in this order:",
    `  ${SKILL_DIR}/references/codegen.md     (READ FIRST — the buildPage return contract, double-quote string-escaping, output framing, the API mistakes that silently corrupt output, and a pre-emit self-check.)`,
    `  ${SKILL_DIR}/SKILL.md                   (building blocks + reusable helpers — runsFromText, bodyPara, heading, caption.)`,
    `  ${SKILL_DIR}/references/layout.md        (ONLY if the page mixes column counts — sections and CONTINUOUS breaks.)`,
    `  ${SKILL_DIR}/references/tables.md        (ONLY if the page has a table — borders, spans, cell margins, numeric/decimal-aligned columns, totals rows.)`,
    "",
    "STEP 2 — The page's text has already been extracted for you deterministically. Reconstruct THIS page from it.",
    "  Each line below is:  x=<left>  y=<top>  | <text>   with sizes in points and origin at the TOP-LEFT.",
    "  Read the geometry to recover layout — there is NO image, so the coordinates ARE the layout:",
    "    - COLUMNS: lines whose x-values cluster into two (or more) well-separated bands that both run down the page are",
    "      SIDE-BY-SIDE COLUMNS. Emit them as separate section columns / a layout table (see layout.md) — do NOT interleave",
    "      them into one stream by y alone. Read each column top-to-bottom in full before moving to the next.",
    "    - TABLES: a run of lines at the same y with text at several repeated x-positions is a TABLE ROW; the repeated",
    "      x-positions are the columns. Rebuild it as a real docx table with those columns, one row per y. Right-align",
    "      numeric columns. A short bold line spanning the width just above such a block is its heading, not a cell.",
    "    - HEADINGS: lines tagged [HEADING ~Npt] or [bold] are headings/labels, not body text.",
    "    - FILLS / BORDERS: if a FILLS section is present, those are the page's EXACT background colors (with the text each",
    "      sits behind). Apply each as the `shading:{fill}` on the matching heading/section bar or table cell, and set that",
    "      block's TextRun color to contrast (white/light text on a dark fill). A colored bar spanning a column (e.g. a green",
    "      \"TRANSACTION RATIONALE\" bar) is a one-cell full-width shaded table/paragraph, not a plain heading. If a BORDERS",
    "      line is present, draw matching table gridlines in that color/weight. NOTE: docx border `size` is in EIGHTHS of a point, so a 1pt PDF border is `size: 8` (NOT 4). Use ONLY the colors given — never invent others.",
    "",
    "----- EXTRACTED PAGE -----",
    textPrompt,
    "----- END EXTRACTED PAGE -----",
    "",
    "STEP 3 — Emit an ES module for THIS host (the skill's \"Mode B\"):",
    "  export function buildPage(docx, ctx) { ...; return { sections, styles, header, footer, pageNumberStart }; }",
    "  - `sections` is an ARRAY of Word section objects, each { properties, children }.",
    "  - Put page size/margins on the FIRST section; later sections use type: docx.SectionType.CONTINUOUS.",
    "  - Reproduce the table/column STRUCTURE the coordinates imply. Transcribe the text EXACTLY as given (it is already",
    "    correct — do not 'fix', reflow, or summarize it). Since you cannot see colors, use clean borderless tables and",
    "    plain headings unless a fill is clearly implied; do NOT invent fills or colors.",
    "  - RUNNING HEADER/FOOTER: text pinned to the very top/bottom margin that would recur every page (a running head, a",
    "    page number like \"Page 3 of 20\", a confidentiality line) is FURNITURE — return it as `header`/`footer`, never in",
    "    `children`. Use docx.PageNumber.CURRENT for a live page number; never hard-code the digit.",
    "",
    "You have ONLY the Read tool. Do NOT narrate. Output ONLY the JavaScript module — start at the first character of code,",
    "end at the last. No markdown fences, no commentary before or after.",
    ...retryNote,
  ].join("\n");
}

/** Strip a whole-output ```fence``` if the model wrapped everything in one. */
function stripOuterFence(md) {
  const t = (md || "").trim();
  const m = t.match(/^```[a-zA-Z]*\n([\s\S]*?)\n```\s*$/);
  return m ? m[1].trim() : t;
}

function buildMarkdownPrompt(imageAbsPath, assetFiles) {
  const assets = assetFiles && assetFiles.length ? assetFiles.map((f) => `"${f}"`).join(", ") : "(none)";
  return [
    "Reconstruct a single document page as pure GitHub-Flavored Markdown (GFM) by LOOKING AT the page image.",
    "The Markdown is converted to Word by the markdown-docx library, which is a CommonMark+GFM converter and",
    "renders NO raw HTML — any HTML tag would appear in the document as literal visible text. So output PURE GFM ONLY.",
    "",
    `Read the page image: ${imageAbsPath}`,
    "",
    "STRICT RULES:",
    "  - NO HTML whatsoever: no <div>, <p>, <span>, <b>, <i>, <u>, <sup>, <sub>, <br>, and no &nbsp; / &amp; entities.",
    "  - Bold = **text**, italic = *text*, bold+italic = ***text***. Headings/section titles = # .. ###### by level.",
    "  - Superscripts have no GFM form — write them inline as plain text: \"30th October\", not 30<sup>th</sup>.",
    "  - Indented / nested sub-points (a), b), c) style): use blockquotes (> ) or nested list indentation, NOT spaces or margins.",
    "  - Tabular or column-aligned data (e.g. a salary breakup, any label:value grid) MUST be a GFM table.",
    "    Right-align numeric/amount columns using the ---: separator. Never fake alignment with spaces or &nbsp;.",
    "  - Enumerations = ordered (1.) or unordered (-) lists. Separate blocks with a blank line.",
    `  - Images: reference ONLY these provided asset files by name — ![alt](filename). Available: ${assets}.`,
    "    Add no image that isn't in that list; if a figure has no asset, describe it in one italic line instead.",
    "",
    "Output ONLY the Markdown — start at the first character of content, end at the last.",
    "No code fences around the whole thing, no commentary before or after.",
  ].join("\n");
}

function buildHtmlPrompt(imageAbsPath, assetFiles) {
  const assets = assetFiles && assetFiles.length ? assetFiles.map((f) => `"${f}"`).join(", ") : "(none)";
  return [
    "Reconstruct a single document page as clean SEMANTIC HTML by LOOKING AT the page image.",
    "The HTML is converted to Word by pandoc, which PARSES HTML structure but IGNORES CSS.",
    "So encode meaning with TAGS, never with style/color/flex — CSS attributes are silently dropped.",
    "",
    `Read the page image: ${imageAbsPath}`,
    "",
    "RULES:",
    "  - Use semantic tags: <h1>..<h6> for headings, <p>, <strong>, <em>, <ul>/<ol> + <li>, <a href=\"...\">.",
    "  - MULTI-COLUMN layouts (side-by-side text, a sidebar, a resume card, two-up columns) MUST be a <table>",
    "    with ONE <tr> and one <td> per column. This is the ONLY way columns survive into Word. Never use",
    "    CSS float/flex/columns for layout — pandoc ignores them and everything collapses into one column.",
    "  - Real data tables: <table> with <tr> and <th>/<td>. Keep header cells as <th>.",
    `  - Images: <img src=\"FILENAME\" alt=\"...\"> using ONLY these provided asset files: ${assets}.`,
    "    Add no image not in that list; if a figure has no asset, describe it in one <em>...</em> line instead.",
    "  - Output an HTML FRAGMENT only — no <html>, <head>, <body>, <style>, <script>, or DOCTYPE.",
    "  - You MAY keep light inline styling for readability, but never depend on it for structure or meaning.",
    "",
    "  SEMANTIC CLASSES (a downstream filter maps these to real Word styles — colors/shading come from them, so USE them):",
    "  - A colored/filled SIDEBAR column: put its blocks in the FIRST <td>, and wrap them:",
    "      * the dark-filled header block (photo, name, contact) → <div class=\"sidebar-dark\"> ... </div>",
    "      * the lighter-filled remainder (skills, certifications, awards) → <div class=\"sidebar\"> ... </div>",
    "  - A tinted intro/summary box (usually atop the main column) → <div class=\"intro\"> ... </div>",
    "  - Skill/rating DOTS or bars (e.g. ●●●●○): reproduce the glyphs literally inside <span class=\"dots\">●●●●○</span>.",
    "  - Do NOT invent these classes for plain single-column pages; use them only where the page truly has that treatment.",
    "",
    "Output ONLY the HTML — start at the first tag, end at the last. No markdown fences, no commentary.",
  ].join("\n");
}

/** Strip a leading ```html fence / trailing ``` if the model wrapped the HTML. */
function stripHtmlFence(s) {
  const t = stripOuterFence(s);
  return t.replace(/^```html\s*/i, "").replace(/```\s*$/, "").trim();
}

/**
 * Ask headless Claude to reconstruct one page image as semantic HTML for pandoc.
 * @returns {Promise<string>} HTML fragment for the page
 */
export function generatePageHtml(imageAbsPath, opts) {
  const args = [
    "-p", buildHtmlPrompt(imageAbsPath, opts.assetFiles),
    "--model", opts.model,
    "--allowedTools", "Read",
    "--output-format", "text",
  ];
  return new Promise((resolve, reject) => {
    execFile("claude", args, { timeout: opts.timeoutMs, maxBuffer: 32 * 1024 * 1024 }, (err, stdout, stderr) => {
      if (err) {
        reject(new Error(`claude failed: ${err.message}${stderr ? ` | ${stderr.slice(0, 400)}` : ""}`));
        return;
      }
      const html = stripHtmlFence(stdout);
      if (!html) { reject(new Error("model returned empty html")); return; }
      resolve(html);
    });
  });
}

/**
 * Ask headless Claude to reconstruct one page image as pure GFM Markdown.
 * @returns {Promise<string>} GFM Markdown for the page (outer fence stripped)
 */
export function generatePageMarkdown(imageAbsPath, opts) {
  const args = [
    "-p", buildMarkdownPrompt(imageAbsPath, opts.assetFiles),
    "--model", opts.model,
    "--allowedTools", "Read",
    "--output-format", "text",
  ];
  return new Promise((resolve, reject) => {
    execFile("claude", args, { timeout: opts.timeoutMs, maxBuffer: 32 * 1024 * 1024 }, (err, stdout, stderr) => {
      if (err) {
        reject(new Error(`claude failed: ${err.message}${stderr ? ` | ${stderr.slice(0, 400)}` : ""}`));
        return;
      }
      const md = stripOuterFence(stdout);
      if (!md) { reject(new Error("model returned empty markdown")); return; }
      resolve(md);
    });
  });
}

/**
 * Ask headless Claude to generate the docx-building module for one page image.
 * @param {object} opts
 * @param {{code:string, errors:string[]}} [opts.priorAttempt]  a previous attempt that
 *   failed verification — its errors and source are fed back into the prompt as a
 *   targeted retry hint instead of asking Claude to start over blind.
 * @returns {Promise<string>} ES module source exporting `buildPage(docx, ctx)`
 */
export function generatePageCode(imageAbsPath, opts) {
  return runCodegen(buildPrompt(imageAbsPath, opts.assetFiles, opts.priorAttempt), opts);
}

/**
 * Text-only counterpart of generatePageCode: reconstructs the page from the
 * deterministically extracted text+coordinates (pdfextract.py) with NO page
 * image — zero vision tokens, exact text. Used by the --py-extract LLM path.
 * @returns {Promise<string>} ES module source exporting `buildPage(docx, ctx)`
 */
export function generatePageCodeFromText(textPrompt, opts) {
  return runCodegen(buildTextPrompt(textPrompt, opts.assetFiles, opts.priorAttempt), opts);
}

// ---- Compact layout-JSON worker (the ~2×-faster path) -----------------------
// Instead of writing ~6KB of docx.js code (the output-length bottleneck, ≈2 min
// per page), the model emits a SMALL layout JSON — transcription + block
// assembly only. A deterministic renderer (src/docx.js renderLayoutJson) expands
// it and STAMPS the fills/borders/frame we already extracted, so the model never
// spends output tokens (or gets a chance to be wrong) on colors or the frame.

function buildJsonPrompt(source, assetFiles, priorAttempt) {
  const assets = assetFiles && assetFiles.length ? assetFiles.map((f) => `"${f}"`).join(", ") : "(none)";
  const retryNote = priorAttempt ? [
    "",
    "RETRY — your previous JSON failed verification. Fix these and do not repeat them:",
    ...priorAttempt.errors.map((e) => `  - ${e}`),
  ] : [];
  const sourceBlock = source.textPrompt
    ? [
        "The page's text has ALREADY been extracted for you deterministically. Reconstruct THIS page from it.",
        "Each line is:  x=<left>  y=<top>  | <text>   (points, origin TOP-LEFT). The coordinates ARE the layout:",
        "  - Two (or more) well-separated x-bands that both run down the page = SIDE-BY-SIDE COLUMNS → one entry per column in `columns`.",
        "  - A run of lines sharing a y with text at several repeated x-positions = a TABLE ROW; repeated x-positions are the columns.",
        "  - Lines tagged [HEADING ~Npt] are headings; [bold] lines are bold labels/section bars.",
        "  - IGNORE any FILLS/BORDERS section for COLOR — the host injects exact colors and the frame deterministically. But DO use a",
        "    labeled colored bar (e.g. a \"TRANSACTION RATIONALE\" band) as a `{t:\"bar\",text}` block so the host can bind its fill.",
        "",
        "----- EXTRACTED PAGE -----",
        source.textPrompt,
        "----- END EXTRACTED PAGE -----",
      ]
    : [
        `Read the page image and reconstruct it: ${source.imageAbsPath}`,
        "Transcribe the text EXACTLY. Do NOT choose colors or draw a frame — the host injects fills/borders/frame from the PDF's",
        "vector layer. Still mark a full-width colored heading band as a `{t:\"bar\",text}` block so the host can bind its fill.",
      ];
  return [
    "You reconstruct ONE document page as a COMPACT LAYOUT JSON (NOT code, NOT prose). A deterministic host renderer turns",
    "your JSON into an editable Word page and stamps the exact fills/borders/frame itself. Your job is transcription + block",
    "assembly ONLY — that is what keeps this fast. Output the JSON and nothing else.",
    "",
    ...sourceBlock,
    "",
    "SCHEMA — output exactly this shape (omit optional keys you don't need):",
    "  { \"title\"?: string,",
    "    \"frame\"?: boolean,                 // true only if page content sits inside an outer box/border",
    "    \"styles\"?: { \"font\"?: string, \"size\"?: number },   // size is HALF-points, e.g. 22 = 11pt",
    "    \"columns\": [ { \"blocks\": [ Block, ... ] }, ... ] }   // ONE object per side-by-side column",
    "  Block =",
    "    { \"t\":\"bar\",  \"text\": string }                      // full-width colored section band (host fills the color)",
    "    { \"t\":\"h\",    \"text\": string, \"level\"?: 1|2|3|4 }   // heading",
    "    { \"t\":\"p\",    \"text\": string }                      // paragraph",
    "    { \"t\":\"list\", \"ordered\"?: boolean, \"items\": [string,...] }",
    "    { \"t\":\"table\",\"header\"?: boolean, \"cols\"?: [number,...], \"rows\": [ [cell,...], ... ] }",
    "        cell = string | { \"text\": string, \"bold\"?: boolean, \"align\"?: \"left\"|\"right\"|\"center\" }",
    "        header defaults true (row 0 is the header); cols are RELATIVE widths; numeric columns are auto right-aligned.",
    "    { \"t\":\"img\",  \"file\": string }                      // ONLY a provided asset filename (see below)",
    "",
    `IMAGES: the page's figures/photos have ALREADY been cropped to these asset files: ${assets}. If the page has a figure,`,
    "  place a `{\"t\":\"img\",\"file\":\"…\"}` block where it appears, using ONLY a filename from that list. Do NOT invent filenames;",
    "  do NOT drop a figure — every provided asset should appear once, in reading order, in the column where it sits.",
    "",
    "RULES: Transcribe text EXACTLY (it is already correct — do not reflow, fix, or summarize). Use `columns` ONLY for genuine",
    "side-by-side columns; a normal single-column page is one column. Do NOT invent fills, colors, or a frame — leave them to the host.",
    "Output ONLY minified JSON — start with { and end with }. No markdown fences, no commentary.",
    ...retryNote,
  ].join("\n");
}

/** Strip a ```json fence if the model wrapped the object, then isolate {...}. */
function parseLayoutJson(raw) {
  let t = stripOuterFence(raw || "").trim();
  const first = t.indexOf("{"), last = t.lastIndexOf("}");
  if (first === -1 || last === -1 || last <= first) throw new Error("no JSON object in model output");
  t = t.slice(first, last + 1);
  let obj;
  try { obj = JSON.parse(t); } catch (e) { throw new Error(`invalid JSON: ${e.message}`); }
  if (!obj || typeof obj !== "object") throw new Error("parsed JSON is not an object");
  if (!Array.isArray(obj.columns) && !Array.isArray(obj.blocks)) throw new Error("JSON has neither `columns` nor `blocks`");
  return obj;
}

/**
 * Ask headless Claude to emit the compact layout JSON for one page — from the
 * py-extract text (opts.textPrompt) when available, else from the page image.
 * @returns {Promise<object>} the parsed layout JSON
 */
export function generatePageJson(source, opts) {
  const args = ["-p", buildJsonPrompt(source, opts.assetFiles, opts.priorAttempt), "--model", opts.model, "--allowedTools", "Read", "--output-format", "text"];
  return new Promise((resolve, reject) => {
    execFile("claude", args, { timeout: opts.timeoutMs, maxBuffer: 32 * 1024 * 1024 }, (err, stdout, stderr) => {
      if (err) { reject(new Error(`claude failed: ${err.message}${stderr ? ` | ${stderr.slice(0, 400)}` : ""}`)); return; }
      try { resolve(parseLayoutJson(stdout)); }
      catch (e) { reject(new Error(`could not parse layout JSON: ${e.message}`)); }
    });
  });
}

/** Shared claude -p codegen call: run the prompt, extract the buildPage module. */
function runCodegen(prompt, opts) {
  const args = ["-p", prompt, "--model", opts.model, "--allowedTools", "Read", "--output-format", "text"];
  return new Promise((resolve, reject) => {
    execFile(
      "claude",
      args,
      { timeout: opts.timeoutMs, maxBuffer: 32 * 1024 * 1024 },
      (err, stdout, stderr) => {
        if (err) {
          reject(new Error(`claude failed: ${err.message}${stderr ? ` | ${stderr.slice(0, 400)}` : ""}`));
          return;
        }
        // extractModule strips fences AND any prose the model wraps around the
        // module (e.g. a leading "I'll output the module directly..." sentence),
        // re-injects a canonical `import * as docx from "docx"`, and throws if the
        // buildPage export is missing. This is the layer that actually eliminates
        // narration-in-output — prompt rules only reduce it (codegen.md §6).
        let code;
        try {
          code = extractModule(stdout || "", { injectDocxImport: true, requireExport: "buildPage" });
        } catch (e) {
          reject(new Error(`could not extract module from model output: ${e.message}`));
          return;
        }
        resolve(code);
      }
    );
  });
}

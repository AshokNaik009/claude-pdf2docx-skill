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
  const args = [
    "-p",
    buildPrompt(imageAbsPath, opts.assetFiles, opts.priorAttempt),
    "--model",
    opts.model,
    "--allowedTools",
    "Read",
    "--output-format",
    "text",
  ];
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

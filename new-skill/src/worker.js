// Headless Claude worker. Per the architecture, Claude LOOKS AT THE PAGE IMAGE
// and GENERATES the docx-library JavaScript that reconstructs the page. We do
// not parse Markdown — Claude is the code generator. The merge step executes
// the generated module and stitches pages into the final .docx.
import { execFile } from "node:child_process";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const HERE = dirname(fileURLToPath(import.meta.url));
const SKILL_DIR = resolve(HERE, "..", ".claude", "skills", "docxjs");

function buildPrompt(imageAbsPath, assetFiles) {
  const assets = assetFiles && assetFiles.length
    ? assetFiles.map((f) => `"${f}"`).join(", ")
    : "(none)";
  return [
    "You reconstruct a single document page as editable Word content by writing JavaScript with the docx.js library.",
    "",
    "STEP 1 — Read the docx.js authoring skill (use the Read tool) and follow it:",
    `  ${SKILL_DIR}/SKILL.md`,
    `  ${SKILL_DIR}/references/layout.md      (sections, continuous breaks, mixing column counts — READ THIS)`,
    `  ${SKILL_DIR}/references/tables.md       (borders, spans, cell margins, numeric columns — read if the page has a table)`,
    `STEP 2 — Read the page image to reconstruct: ${imageAbsPath}`,
    "",
    "STEP 3 — Output an ES module exporting exactly one function:",
    "  export function buildPage(docx, ctx) { ...; return { sections, styles }; }",
    "",
    "Return contract:",
    "- Return { sections, styles } where `sections` is an ARRAY of Word section objects and `styles`",
    "  (optional) is a document-wide default, e.g. { default: { document: { run: { font: 'Palatino Linotype', size: 22 } } } }.",
    "  Set `styles` to match the page's body font/size if you can identify it. (A single section object",
    "  { properties, children } or a bare children array are also accepted, but prefer { sections, styles }.)",
    "- Each section = { properties, children }. properties may include:",
    "    page:   { size: { width: 12240, height: 15840 }, margin: { top:1440,bottom:1440,left:1440,right:1440 } }  // put on the FIRST section",
    "    column: { count: N, space: 560, separate: false }   // N = number of TEXT COLUMNS you SEE in that band",
    "    type:   docx.SectionType.CONTINUOUS                  // on later sections so they DON'T start a new page",
    "",
    "CRITICAL — mixed column layouts (this is what the skill's references/layout.md is for):",
    "- If the page is two-column text with a FULL-WIDTH table or figure in the middle (very common in",
    "  academic papers), emit MULTIPLE sections, exactly like the skill shows:",
    "    section 1: { properties: { column: { count: 2, space: 560 }, page: {...} }, children: [ ...two-column body... ] }",
    "    section 2: { properties: { type: docx.SectionType.CONTINUOUS }, children: [ caption, fullWidthTable, figure ] }",
    "    section 3: { properties: { type: docx.SectionType.CONTINUOUS, column: { count: 2, space: 560 } }, children: [ ...resume two-column... ] }",
    "  A full-width table or figure must NEVER sit inside a two-column section.",
    "- Within a column band, output paragraphs in reading order (entire left column, then entire right);",
    "  Word flows them automatically. Never interleave columns line by line.",
    "",
    "Fidelity:",
    "- Do NOT import/require anything, do NOT construct a Document, do NOT call Packer. Use only `docx` and `ctx`.",
    "- Reuse the skill's helpers (runsFromText auto-italicizes 'et al.', bodyPara, heading, caption).",
    "- Tables: visible single-line borders, header row bold + tableHeader:true, cell padding, rowSpan/columnSpan when merged.",
    `- Images extracted from this page (embed with ctx.image("<file>") -> ImageRun, wrapped in a centered Paragraph): ${assets}.`,
    "- FIGURES/DIAGRAMS: if a matching image file is listed above, embed it with ctx.image(). If NONE is listed",
    "  (a vector diagram was not extracted), insert a single centered placeholder paragraph like",
    "  '[ Figure: <caption text> ]'. Do NOT fabricate a table of labels to imitate the diagram.",
    "- Transcribe ALL visible text verbatim. Do not summarize, translate, or invent content.",
    "",
    "Output ONLY the JavaScript module. No markdown code fences, no commentary before or after.",
  ].join("\n");
}

/**
 * Strip accidental ``` fences and any docx import/require Claude wrote, then
 * prepend a real `import * as docx from "docx"` so `docx` resolves everywhere
 * in the module (including any helper Claude defines outside buildPage). The
 * buildPage(docx, ctx) param still shadows it locally with the same object.
 */
function sanitize(code) {
  let c = code.trim();
  const fence = c.match(/^```[a-zA-Z]*\n([\s\S]*?)\n```$/);
  if (fence) c = fence[1];
  c = c
    .split("\n")
    .filter((ln) => !/^\s*(import\s.*from\s*['"]docx['"]|const\s.*=\s*require\(['"]docx['"]\))/.test(ln))
    .join("\n")
    .trim();
  return `import * as docx from "docx";\n\n${c}`;
}

/**
 * Ask headless Claude to generate the docx-building module for one page image.
 * @returns {Promise<string>} ES module source exporting `buildPage(docx, ctx)`
 */
export function generatePageCode(imageAbsPath, opts) {
  const args = [
    "-p",
    buildPrompt(imageAbsPath, opts.assetFiles),
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
        const code = sanitize(stdout || "");
        if (!/export\s+function\s+buildPage/.test(code)) {
          reject(new Error("generated code missing `export function buildPage`"));
          return;
        }
        resolve(code);
      }
    );
  });
}

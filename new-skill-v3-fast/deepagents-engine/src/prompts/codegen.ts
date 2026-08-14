// Distilled buildPage(docx, ctx) contract — from vendor/src/worker.js
// buildPrompt/buildTextPrompt + vendor/skills/docxjs/references/codegen.md
// ("Mode B"). Used as the codegen-fallback section of the page_builder and
// gate_fixer prompts; the full references stay on disk (decision D9).
export function CODEGEN_RULES(skillDir: string): string {
  return [
    "Emit an ES module whose whole text you pass as the `source` argument of validate_module:",
    "  export function buildPage(docx, ctx) { ...; return { sections, styles }; }",
    "- `sections` is an ARRAY of Word section objects, each { properties, children }. children = Paragraphs/Tables.",
    "- The FIRST section may omit `properties` (the host applies US-Letter/1in defaults); every LATER section MUST set",
    "  properties: { type: docx.SectionType.CONTINUOUS } or it starts a new page.",
    "- `styles` (optional): { default: { document: { run: { font: \"...\", size: 22 } } } }  // size in HALF-points",
    "- Mode B hard rules: do NOT import or require ANYTHING (docx arrives as the parameter); never construct",
    "  new docx.Document(...) or call docx.Packer; no fs/fetch/console.log; define helpers INSIDE buildPage so `docx`",
    "  is in scope; export exactly the one function buildPage.",
    "- Images: ONLY via ctx.image(\"<asset filename>\") inside a Paragraph's children — the host sizes it correctly.",
    "- STRING SAFETY (the #1 failure): double-quoted strings; escape exactly two characters: \" → \\\" and \\ → \\\\.",
    "  One paragraph = ONE source line — never put a literal line break inside a string.",
    "- UNITS: font size = HALF-points (22 = 11pt); border size = EIGHTHS of a point (1pt border ⇒ size: 8);",
    "  twips everywhere else (1440 = 1 inch).",
    "- Tables: every cell needs a Paragraph child; covered cells of a columnSpan are OMITTED, not empty cells;",
    "  give the Table width: { size: 9360, type: docx.WidthType.DXA } and columnWidths that sum to it.",
    "- The `source` string must be pure code: no markdown fences, no commentary before or after.",
    `Full reference on disk if the host lets you read files: ${skillDir}/references/codegen.md (also tables.md, layout.md, images.md).`,
  ].join("\n");
}

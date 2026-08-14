// gate_fixer subagent prompt (spec §7.7): repairs the specific pages the gate
// flagged, using the same validators as page_builder; the orchestrator
// re-merges and re-gates afterwards.
import { LAYOUT_SCHEMA_TEXT } from "./page-builder.js";
import { CODEGEN_RULES } from "./codegen.js";
import type { EngineConfig } from "../config.js";

export function gateFixerPrompt(cfg: EngineConfig, skillDir: string): string {
  return [
    "You are gate_fixer. The whole-document verification gate FAILED and your task message contains the failing",
    "checks plus perPageHints naming the suspect pages. Repair those specific pages, then report — do NOT merge or",
    "run the gate yourself; the orchestrator re-merges after you.",
    "",
    "PROCEDURE:",
    "1. Call get_session to see each page's status/mode/reason.",
    "2. For each hinted page (hint page 0 = document-wide, e.g. a table overflow — find the culprit via get_session",
    "   modes and the check detail):",
    "   a. Call read_page_input {page} for its textPrompt, injected facts, assetFiles, and priorErrors.",
    "   b. Rebuild the page's layout JSON and call validate_layout {page, layout}, fixing per the hint:",
    '      - "injected fill #XXXXXX missing" → the page must carry a {"t":"bar","text":...} block whose text matches the',
    "        labeled band in textPrompt (see injected.fills[].text) so the host can bind that fill — add or fix it.",
    "      - low word coverage → transcribe MORE of textPrompt into the layout; do not summarize.",
    '      - image not embedded → add {"t":"img","file":"..."} blocks for every assetFiles entry.',
    "      - table overflow → simplify that table: fewer columns, or split it; keep cols as small relative widths.",
    `   c. At most ${cfg.maxBuilderAttempts} validate_layout fix rounds per page; then ONE validate_module attempt with`,
    "      docx.js code (CODEGEN RULES below); then, ONLY if the page is truly unfixable, mark_page_image_fallback",
    "      {page, reason}.",
    "3. Report per page: what you changed, ok/failed, and the mode it ended in.",
    "",
    LAYOUT_SCHEMA_TEXT,
    "",
    "RULES: transcribe text EXACTLY from textPrompt; never invent fills/colors/frames (the host stamps them); use",
    "`columns` only for genuine side-by-side columns; sizes are HALF-points.",
    "",
    "CODEGEN RULES (only for the single validate_module fallback):",
    CODEGEN_RULES(skillDir),
  ].join("\n");
}

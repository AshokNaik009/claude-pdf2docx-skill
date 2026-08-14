// The fix-loop heart (spec §7.5.2, decision D6): render + verify + coverage in
// one call, returning errors the model can act on. validateLayoutForPage is the
// single write path for mode:"json" manifests — the page_builder tool, the
// gate_fixer tool, AND vision.ts (M3) all go through it, so a page completed by
// any rung is byte-identical in shape to the parent engine's manifests.
import { existsSync, readFileSync, rmSync, writeFileSync, mkdirSync } from "node:fs";
import { join } from "node:path";
// @ts-ignore vendored JS
import { ensureImages, layoutText, coverageOfText } from "../../vendor/src/pipeline.js";
// @ts-ignore vendored JS
import { renderLayoutJson, makeCtx, executeModule } from "../../vendor/src/docx.js";
// @ts-ignore vendored JS
import { verifySections } from "../../vendor/src/verify.js";
// @ts-ignore vendored JS
import { STATUS, loadSession, saveSession } from "../../vendor/src/session.js";
// @ts-ignore vendored JS
import { extractModule } from "../../vendor/skills/docxjs/scripts/extract-code.mjs";
import { pagesRelFor, type Injected } from "./deterministic-pass.js";
import type { EngineConfig } from "../config.js";

export interface ValidateResult {
  ok: boolean;
  coverage?: number;
  errors: string[];
  warnings?: string[];
}

interface PageFiles {
  pageState: any;
  session: any;
  pagesAbs: string;
  assetsAbs: string;
  meta: { reason: string; assets: string[]; assetFiles: string[] };
  injected: Injected;
  origText: string;
}

/** Split origtext.txt the way pipeline.js fallbackParagraphs splits page text. */
const toFallbackParagraphs = (origText: string) =>
  origText
    .split(/\n\s*\n/)
    .map((s) => s.replace(/\s+/g, " ").trim())
    .filter(Boolean);

/**
 * Load the per-page prewritten inputs (deterministic-pass wrote them for every
 * needs_llm page). Throws a model-actionable error if the page has no inputs
 * (i.e. it was never routed to the LLM).
 */
export function loadPageFiles(workDir: string, page: number): PageFiles {
  const session = loadSession(workDir);
  const pageState = session.pages.find((p: any) => p.page === page);
  if (!pageState) throw new Error(`page ${page} is not in this document (1..${session.totalPages})`);
  const pagesAbs = join(workDir, pagesRelFor(pageState));
  const metaPath = join(pagesAbs, "meta.json");
  if (!existsSync(metaPath)) {
    throw new Error(
      `page ${page} has no pre-extracted LLM inputs (status: ${pageState.status}) — only pages the deterministic pass flagged as complex can be rebuilt`
    );
  }
  return {
    pageState,
    session,
    pagesAbs,
    assetsAbs: join(workDir, pageState.assets),
    meta: JSON.parse(readFileSync(metaPath, "utf8")),
    injected: JSON.parse(readFileSync(join(pagesAbs, "injected.json"), "utf8")),
    origText: readFileSync(join(pagesAbs, "origtext.txt"), "utf8"),
  };
}

const errorsPath = (pagesAbs: string) => join(pagesAbs, "errors.json");

function recordFailure(pagesAbs: string, errors: string[], coverage?: number) {
  writeFileSync(errorsPath(pagesAbs), JSON.stringify({ errors, ...(coverage !== undefined ? { coverage } : {}) }, null, 2));
}

function markCompleted(workDir: string, f: PageFiles, mode: string, tokens: number) {
  f.pageState.status = STATUS.COMPLETED;
  f.pageState.mode = mode;
  f.pageState.reason = f.meta.reason;
  f.pageState.tokens = tokens;
  f.pageState.error = null;
  saveSession(workDir, f.session);
  rmSync(errorsPath(f.pagesAbs), { force: true });
}

/**
 * One compact-layout-JSON validation attempt (mirror of pipeline.js jsonOnce):
 * ensureImages → renderLayoutJson (stamps injected fills/frame) → verifySections
 * → coverage fidelity. PASS writes layout.json + the mode:"json" manifest and
 * marks the page completed; FAIL records errors.json for the retry loop.
 */
export async function validateLayoutForPage(
  workDir: string,
  cfg: EngineConfig,
  page: number,
  layout: any,
  opts: { model?: string; llmInput?: "text" | "image" } = {}
): Promise<ValidateResult> {
  const f = loadPageFiles(workDir, page);
  if (!layout || typeof layout !== "object" || (!Array.isArray(layout.columns) && !Array.isArray(layout.blocks))) {
    const errors = ['layout must be an object with a `columns` array (one entry per side-by-side column), e.g. {"columns":[{"blocks":[...]}]}'];
    recordFailure(f.pagesAbs, errors);
    return { ok: false, errors };
  }

  const full = ensureImages(layout, f.meta.assetFiles);
  let verified: { ok: boolean; errors: string[]; warnings: string[] };
  try {
    const { sections, styles } = renderLayoutJson(full, makeCtx(f.assetsAbs), f.injected);
    verified = await verifySections({ sections, styles });
  } catch (e) {
    verified = { ok: false, errors: [`layout failed to render: ${(e as Error).message}`], warnings: [] };
  }
  const cov: number = coverageOfText(f.origText, layoutText(full), cfg);
  if (verified.ok && cov < cfg.coverageMin) {
    verified.ok = false;
    verified.errors = [...verified.errors, `low fidelity: only ${(cov * 100).toFixed(0)}% of the page's words reconstructed`];
  }

  if (!verified.ok) {
    recordFailure(f.pagesAbs, verified.errors, cov);
    return { ok: false, coverage: cov, errors: verified.errors, warnings: verified.warnings };
  }

  mkdirSync(f.pagesAbs, { recursive: true });
  writeFileSync(join(f.pagesAbs, "layout.json"), JSON.stringify(full, null, 2));
  const manifest = {
    page,
    mode: "json",
    llmInput: opts.llmInput ?? "text",
    reason: f.meta.reason,
    model: opts.model ?? "page_builder",
    layout: full,
    injected: f.injected,
    fallbackText: toFallbackParagraphs(f.origText),
    assets: f.meta.assets,
    assetsDir: f.pageState.assets,
  };
  writeFileSync(join(workDir, f.pageState.manifest), JSON.stringify(manifest, null, 2));
  markCompleted(workDir, f, "json", Math.round(JSON.stringify(full).length / 4));
  return { ok: true, coverage: cov, errors: [], warnings: verified.warnings };
}

/**
 * The docx.js code-gen fallback rung: extract the buildPage module from the
 * model's source, execute it, structurally verify (NO coverage check — parent
 * parity). PASS writes page.mjs + the mode:"llm" manifest.
 */
export async function validateModuleForPage(
  workDir: string,
  cfg: EngineConfig,
  page: number,
  source: string,
  opts: { model?: string } = {}
): Promise<ValidateResult> {
  const f = loadPageFiles(workDir, page);

  let code: string;
  try {
    code = extractModule(source, { injectDocxImport: true, requireExport: "buildPage" });
  } catch (e) {
    const errors = [`could not extract module: ${(e as Error).message}`];
    recordFailure(f.pagesAbs, errors);
    return { ok: false, errors };
  }

  mkdirSync(f.pagesAbs, { recursive: true });
  const codeRel = join(pagesRelFor(f.pageState), "page.mjs");
  const codeAbs = join(workDir, codeRel);
  writeFileSync(codeAbs, code);

  let verified: { ok: boolean; errors: string[]; warnings: string[] };
  try {
    const executed = await executeModule(codeAbs, f.assetsAbs);
    verified = await verifySections(executed);
  } catch (e) {
    verified = { ok: false, errors: [`module threw while executing: ${(e as Error).message}`], warnings: [] };
  }

  if (!verified.ok) {
    recordFailure(f.pagesAbs, verified.errors);
    return { ok: false, errors: verified.errors, warnings: verified.warnings };
  }

  const manifest = {
    page,
    mode: "llm",
    llmInput: "text",
    reason: f.meta.reason,
    model: opts.model ?? "page_builder",
    codeFile: codeRel,
    fallbackText: toFallbackParagraphs(f.origText),
    assets: f.meta.assets,
    assetsDir: f.pageState.assets,
  };
  writeFileSync(join(workDir, f.pageState.manifest), JSON.stringify(manifest, null, 2));
  markCompleted(workDir, f, "llm", Math.round(code.length / 4));
  return { ok: true, errors: [], warnings: verified.warnings };
}

/** read_page_input payload: everything a page_builder needs, from the prewrites. */
export function readPageInput(workDir: string, page: number) {
  const f = loadPageFiles(workDir, page);
  const textPrompt = readFileSync(join(f.pagesAbs, "textprompt.txt"), "utf8");
  const assetFiles: string[] = JSON.parse(readFileSync(join(f.pagesAbs, "assetFiles.json"), "utf8"));
  let priorErrors: string[] | undefined;
  const ep = errorsPath(f.pagesAbs);
  if (existsSync(ep)) {
    try { priorErrors = JSON.parse(readFileSync(ep, "utf8")).errors; } catch { /* stale/corrupt — ignore */ }
  }
  return {
    page,
    reason: f.meta.reason,
    textPrompt,
    injected: f.injected,
    assetFiles,
    ...(priorErrors?.length ? { priorErrors } : {}),
  };
}

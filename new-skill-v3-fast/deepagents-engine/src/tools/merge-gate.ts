// Merge + gate plain functions (spec §7.5.3). The tool() wrappers live in
// src/agent.ts's makeTools; the --no-agent CLI path calls these directly.
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { basename, join, resolve } from "node:path";
// @ts-ignore vendored JS
import { buildDocumentModel } from "../../vendor/src/merge.js";
// @ts-ignore vendored JS
import { renderDocx } from "../../vendor/src/docx.js";
// @ts-ignore vendored JS
import { runGate } from "../../vendor/src/gate.js";
// @ts-ignore vendored JS
import { STATUS, loadSession, saveSession } from "../../vendor/src/session.js";
// @ts-ignore vendored JS
import { openPdf, renderPage } from "../../vendor/src/pdf.js";
// @ts-ignore vendored JS
import { fallbackParagraphs } from "../../vendor/src/pipeline.js";
import { pagesRelFor } from "./deterministic-pass.js";
import type { EngineConfig } from "../config.js";

export interface MergeResult {
  docxPath: string;
  mergedPages: number;
  missingPages: { page: number; status: string }[];
}

export interface GateCheck { name: string; ok: boolean; detail: string }
export interface GateResult {
  ok: boolean;
  checks: GateCheck[];
  counts?: Record<string, number>;
  perPageHints: { page: number; hint: string }[];
}

/** Order + merge completed pages into outputs/<outName>.docx. */
export async function mergeDocument(workDir: string, outName: string, cfg: EngineConfig): Promise<MergeResult> {
  const model = buildDocumentModel(workDir);
  mkdirSync(resolve(cfg.outputsDir), { recursive: true });
  const stem = (outName || basename(workDir)).replace(/\.docx$/i, "");
  const buffer = await renderDocx(workDir, model, cfg);
  const docxPath = resolve(cfg.outputsDir, `${stem}.docx`);
  writeFileSync(docxPath, buffer);
  return {
    docxPath,
    mergedPages: model.pages.length,
    missingPages: model.missing.map((p: any) => ({ page: p.page, status: p.status })),
  };
}

const normHex = (h: unknown) => String(h || "").replace(/^#/, "").toUpperCase();

/** Load every completed page's manifest, keyed by page number. */
function loadManifests(workDir: string): Map<number, any> {
  const session = loadSession(workDir);
  const out = new Map<number, any>();
  for (const p of session.pages) {
    if (p.status !== STATUS.COMPLETED) continue;
    try {
      out.set(p.page, JSON.parse(readFileSync(join(workDir, p.manifest), "utf8")));
    } catch {
      /* unreadable manifest — no hints derivable for it */
    }
  }
  return out;
}

/**
 * Run the extractandVerify.py gate on the docx and derive per-page hints from
 * failing checks so a fixer knows WHERE to look (spec §7.5.3).
 */
export function runGateOn(docxPath: string, workDir: string, cfg: EngineConfig): GateResult {
  const model = buildDocumentModel(workDir);
  const gate = runGate(docxPath, model, cfg);
  const manifests = loadManifests(workDir);
  const perPageHints: { page: number; hint: string }[] = [];

  for (const check of gate.checks as GateCheck[]) {
    if (check.ok) continue;
    if (check.name === "injected fills present") {
      // "missing from output: #16A085, #2C3E50" → pages whose injected fills carry those hexes.
      const missing = [...(check.detail.matchAll(/#([0-9A-Fa-f]{6})/g))].map((m) => normHex(m[1]));
      for (const hex of missing) {
        for (const [page, m] of manifests) {
          const fills = (m.injected?.fills || []).map((f: any) => normHex(f.fill));
          const frame = m.injected?.frame?.color ? [normHex(m.injected.frame.color)] : [];
          if (fills.includes(hex) || frame.includes(hex)) {
            perPageHints.push({ page, hint: `injected fill #${hex} missing from output — this page's injected.fills carry it` });
          }
        }
      }
    } else if (check.name === "images embedded") {
      for (const [page, m] of manifests) {
        if ((m.assets || []).length) {
          perPageHints.push({ page, hint: `page has ${(m.assets || []).length} extracted image(s) — verify each appears as an img block` });
        }
      }
    } else if (check.name === "word coverage") {
      for (const [page, m] of manifests) {
        if (m.mode !== "text") {
          perPageHints.push({ page, hint: `word coverage low overall — this ${m.mode}-mode page is a likely source of dropped text` });
        }
      }
    } else if (check.name === "no table exceeds its width") {
      perPageHints.push({ page: 0, hint: `table overflow: ${check.detail}` });
    }
  }

  return { ok: gate.ok, checks: gate.checks, counts: gate.counts, perPageHints };
}

export interface PageSummary {
  page: number;
  status: string;
  mode: string | null;
  reason: string | null;
  error?: string | null;
}

/** Compact per-page status summary (never the raw session file — keep context small). */
export function sessionSummary(workDir: string): { totalPages: number; pages: PageSummary[] } {
  const session = loadSession(workDir);
  return {
    totalPages: session.totalPages,
    pages: session.pages.map((p: any) => ({
      page: p.page,
      status: p.status,
      mode: p.mode,
      reason: p.reason,
      ...(p.error ? { error: p.error } : {}),
    })),
  };
}

/**
 * The ladder's floor: ship the page as its own picture + fallback text. The
 * blocks:[{type:"image"}] manifest shape hits docx.js sectionsForPage's final
 * blocksToChildren branch, which renders via ctx.image against assetsDir.
 */
export function markPageImageFallback(
  workDir: string,
  page: number,
  reason: string,
  cfg: EngineConfig
): { ok: boolean; page: number; mode: string } {
  const session = loadSession(workDir);
  const pageState = session.pages.find((p: any) => p.page === page);
  if (!pageState) throw new Error(`page ${page} not in session (1..${session.totalPages})`);

  const assetsRel = pageState.assets;
  const assetsAbs = join(workDir, assetsRel);
  const pagesAbs = join(workDir, pagesRelFor(pageState));
  mkdirSync(assetsAbs, { recursive: true });
  const pngAbs = join(assetsAbs, "page.png");

  // page.png normally exists (deterministic pass pre-writes it for needs_llm
  // pages); render it if the page failed before the prewrite.
  if (!existsSync(pngAbs)) {
    const doc = openPdf(session.document);
    const render = renderPage(doc, page - 1, cfg.dpi);
    writeFileSync(pngAbs, render.fullPng);
  }

  let fallbackText: string[];
  const origTxtPath = join(pagesAbs, "origtext.txt");
  if (existsSync(origTxtPath)) {
    fallbackText = readFileSync(origTxtPath, "utf8")
      .split(/\n\s*\n/)
      .map((s) => s.replace(/\s+/g, " ").trim())
      .filter(Boolean);
  } else {
    fallbackText = fallbackParagraphs(openPdf(session.document), page - 1);
  }

  // Only page.png is actually embedded in image-fallback mode (the extracted
  // crops are visible inside it) — listing more would make the gate expect
  // media files that were never placed.
  const manifest = {
    page,
    mode: "image",
    reason,
    blocks: [{ type: "image", path: `${assetsRel}/page.png` }],
    fallbackText,
    assets: [`${assetsRel}/page.png`],
    assetsDir: assetsRel,
  };
  writeFileSync(join(workDir, pageState.manifest), JSON.stringify(manifest, null, 2));
  pageState.status = STATUS.COMPLETED;
  pageState.mode = "image";
  pageState.reason = reason;
  pageState.error = null;
  saveSession(workDir, session);
  return { ok: true, page, mode: "image" };
}

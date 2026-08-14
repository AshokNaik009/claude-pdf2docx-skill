// The coarse deterministic-pass (spec §7.5.1, decision D3): process ALL pages
// deterministically in one call. Clean pages are fully reconstructed (zero LLM)
// and marked completed; complex pages are marked needs_llm with their LLM
// inputs pre-written under pages/page_XXXX/ so subagents never re-extract.
// Port of vendor/src/pipeline.js processPage MINUS the mode==="llm" branch.
import { mkdirSync, writeFileSync } from "node:fs";
import { join, basename } from "node:path";
// @ts-ignore vendored JS
import { initSession, buildInjected, fallbackParagraphs } from "../../vendor/src/pipeline.js";
// @ts-ignore vendored JS
import { openPdf, pageText, renderPage, pageStructuredPy } from "../../vendor/src/pdf.js";
// @ts-ignore vendored JS
import { classifyPage } from "../../vendor/src/quality.js";
// @ts-ignore vendored JS
import { extractImageAssets, extractFigureRegions, buildPageSections } from "../../vendor/src/extract.js";
// @ts-ignore vendored JS
import { renderJsonSections, makeCtx } from "../../vendor/src/docx.js";
// @ts-ignore vendored JS
import { verifySections, coverage } from "../../vendor/src/verify.js";
// @ts-ignore vendored JS
import { STATUS, saveSession } from "../../vendor/src/session.js";
import type { EngineConfig } from "../config.js";

export interface Injected {
  fills?: { fill: string; text?: string; bbox?: unknown }[];
  border?: { color: string; width: number } | null;
  frame?: { color: string; width: number } | null;
}

export interface ComplexPage {
  page: number;
  reason: string;
  hasFills: boolean;
  hasFrame: boolean;
  assetFiles: string[];
}

export interface DeterministicPassResult {
  workDir: string;
  totalPages: number;
  deterministicDone: number[];
  complexPages: ComplexPage[];
  failed: { page: number; error: string }[];
}

export type Log = (msg: string) => void;

/** pages/page_XXXX/ dir (rel) for a page state whose assets dir is assets/page_XXXX. */
export const pagesRelFor = (pageState: { assets: string }) => pageState.assets.replace(/^assets/, "pages");

/**
 * Pre-write everything a page_builder / vision worker needs for one complex
 * page, so the LLM loop only ever reads files that already exist.
 */
function prewriteLlmInputs(
  workDir: string,
  pageState: any,
  args: {
    textPrompt: string | null;
    injected: any;
    assetFiles: string[];
    assetPaths: string[];
    origText: string;
    reason: string;
    fullPng: Buffer;
  }
) {
  const pagesRel = pagesRelFor(pageState);
  const pagesAbs = join(workDir, pagesRel);
  const assetsAbs = join(workDir, pageState.assets);
  mkdirSync(pagesAbs, { recursive: true });
  mkdirSync(assetsAbs, { recursive: true });

  writeFileSync(join(assetsAbs, "page.png"), args.fullPng);
  writeFileSync(join(pagesAbs, "textprompt.txt"), args.textPrompt ?? "");
  writeFileSync(join(pagesAbs, "injected.json"), JSON.stringify(args.injected, null, 2));
  writeFileSync(join(pagesAbs, "assetFiles.json"), JSON.stringify(args.assetFiles, null, 2));
  writeFileSync(join(pagesAbs, "origtext.txt"), args.origText);
  writeFileSync(
    join(pagesAbs, "meta.json"),
    JSON.stringify(
      {
        page: pageState.page,
        reason: args.reason,
        hasFills: !!args.injected.fills?.length,
        hasFrame: !!args.injected.frame,
        assetFiles: args.assetFiles,
        assets: args.assetPaths,
      },
      null,
      2
    )
  );
}

/**
 * Process every non-completed page deterministically. Returns the worklist of
 * complex (needs_llm) pages. Idempotent: completed pages are skipped; needs_llm
 * pages get their inputs re-prewritten (cheap, safe).
 */
export async function deterministicPass(
  pdfPath: string,
  cfg: EngineConfig,
  log: Log = console.log
): Promise<DeterministicPassResult> {
  const { workDir, session } = initSession(pdfPath, cfg);
  const doc = openPdf(pdfPath);

  const deterministicDone: number[] = [];
  const complexPages: ComplexPage[] = [];
  const failed: { page: number; error: string }[] = [];

  for (const pageState of session.pages) {
    if (pageState.status === STATUS.COMPLETED) {
      deterministicDone.push(pageState.page);
      continue;
    }
    const started = Date.now();
    const i = pageState.page - 1;
    try {
      const relDir = pageState.assets;
      const pageDir = join(workDir, relDir);
      mkdirSync(pageDir, { recursive: true });

      const { structured, width, textPrompt, drawings } = pageStructuredPy(pdfPath, i, cfg);
      const render = renderPage(doc, i, cfg.dpi);
      const injected = buildInjected(drawings) as Injected;

      const cls = classifyPage(structured, width, cfg);
      const imageAssets = extractImageAssets(render, structured, pageDir, relDir);

      // Vector-drawn figures produce no image block — recover them from caption
      // gaps exactly like pipeline.js does (lines 208–217).
      const figureRegions = extractFigureRegions(structured, render, pageDir, relDir, imageAssets, width);
      const augmentedBlocks = [...(structured.blocks || [])];
      for (const fig of figureRegions) {
        imageAssets.set(augmentedBlocks.length, fig.asset);
        augmentedBlocks.push({ type: "image", bbox: fig.bbox });
      }
      const structuredWithFigures = figureRegions.length ? { ...structured, blocks: augmentedBlocks } : structured;
      const assetFiles: string[] = [...imageAssets.values()].map((a: any) => basename(a.path));
      const assetPaths: string[] = [...imageAssets.values()].map((a: any) => a.path);

      let mode: string = cls.mode;
      let reason: string = cls.reason;

      // Fills-routing rule (pipeline.js 225–229): a "clean" text page carrying
      // vector fills or an outer frame can only be reproduced via injection.
      if (
        mode === "text" && cfg.compactJson && cfg.llmMode !== "none" && cfg.llmMode !== "all" &&
        (injected.fills?.length || injected.frame)
      ) {
        mode = "llm";
        reason = `vector fills/frame present → inject via compact-JSON (${reason})`;
      }

      if (mode === "text") {
        const { sections: jsonSections } = buildPageSections(structuredWithFigures, imageAssets, cfg);
        const allBlocks = jsonSections.flatMap((s: any) => s.blocks);
        const liveSections = renderJsonSections(jsonSections, makeCtx(pageDir));
        const verified = await verifySections({ sections: liveSections });
        const cov = verified.ok ? coverage(pageText(doc, i), allBlocks, cfg) : { ok: false, reasons: [] as string[] };

        const failureReason = !verified.ok
          ? `deterministic build failed verification: ${verified.errors.join("; ")}`
          : !cov.ok
          ? `deterministic build low fidelity: ${cov.reasons.join("; ")}`
          : null;

        if (!failureReason || cfg.llmMode === "none") {
          // Pass — or --llm none, where an unverified text page still ships
          // (parent behavior: never worse than plain text).
          const manifest = {
            page: pageState.page,
            mode: "text",
            reason: failureReason
              ? `${reason} (unverified: ${failureReason}; --llm none forbids escalation)`
              : reason,
            sections: jsonSections,
            fallbackText: fallbackParagraphs(doc, i),
            assets: assetPaths,
            assetsDir: relDir,
          };
          writeFileSync(join(workDir, pageState.manifest), JSON.stringify(manifest, null, 2));
          pageState.status = STATUS.COMPLETED;
          pageState.mode = "text";
          pageState.reason = manifest.reason;
          pageState.chars = pageText(doc, i).length;
          pageState.processingTimeMs = Date.now() - started;
          pageState.error = null;
          saveSession(workDir, session);
          deterministicDone.push(pageState.page);
          log(`  ✓ page ${pageState.page}  [text: ${manifest.reason}]  ${pageState.processingTimeMs}ms`);
          continue;
        }
        reason = failureReason;
        mode = "llm";
      }

      // Complex page → needs_llm: pre-write the LLM inputs for the builders.
      prewriteLlmInputs(workDir, pageState, {
        textPrompt,
        injected,
        assetFiles,
        assetPaths,
        origText: pageText(doc, i),
        reason,
        fullPng: render.fullPng,
      });
      pageState.status = STATUS.NEEDS_LLM;
      pageState.mode = null;
      pageState.reason = reason;
      pageState.chars = pageText(doc, i).length;
      pageState.processingTimeMs = Date.now() - started;
      pageState.error = null;
      saveSession(workDir, session);
      complexPages.push({
        page: pageState.page,
        reason,
        hasFills: !!injected.fills?.length,
        hasFrame: !!injected.frame,
        assetFiles,
      });
      log(`  → page ${pageState.page}  needs LLM  [${reason}]  ${pageState.processingTimeMs}ms`);
    } catch (e) {
      const msg = (e as Error).message;
      pageState.status = STATUS.FAILED;
      pageState.error = msg;
      saveSession(workDir, session);
      failed.push({ page: pageState.page, error: msg });
      log(`  ✗ page ${pageState.page} FAILED: ${msg}`);
    }
  }

  return {
    workDir,
    totalPages: session.totalPages,
    deterministicDone,
    complexPages,
    failed,
  };
}

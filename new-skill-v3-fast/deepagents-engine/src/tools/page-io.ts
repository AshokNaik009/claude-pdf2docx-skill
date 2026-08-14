// Closure tool factory (spec §7.5): every tool binds pdfPath/workDir/cfg from
// the per-run closure — the model NEVER supplies paths. All tools return
// JSON-serialized objects; failures come back as {ok:false, error} strings the
// model can act on instead of run-killing throws.
import { basename } from "node:path";
import { tool } from "langchain";
import { z } from "zod";
import { deterministicPass, type Log } from "./deterministic-pass.js";
import { mergeDocument, runGateOn, sessionSummary, markPageImageFallback } from "./merge-gate.js";
import { readPageInput, validateLayoutForPage, validateModuleForPage } from "./validate.js";
import { visionReconstruct } from "./vision.js";
import type { EngineConfig } from "../config.js";

const pageArg = z.number().int().min(1).describe("1-based page number");

// LOOSE layout schema (HANDOFF M2 note): validation failures must come from
// verifySections' actionable wording, not from zod rejecting the tool call.
const layoutSchema = z
  .looseObject({
    title: z.string().optional(),
    frame: z.boolean().optional(),
    styles: z.looseObject({ font: z.string().optional(), size: z.number().optional() }).optional(),
    columns: z.array(z.looseObject({ blocks: z.array(z.any()) })).optional(),
    blocks: z.array(z.any()).optional(),
  })
  .describe("compact layout JSON: { title?, frame?, styles?, columns:[{blocks:[Block]}] }");

const asJson = (v: unknown) => JSON.stringify(v);
const errJson = (e: unknown) => asJson({ ok: false, error: (e as Error).message });

export interface EngineTools {
  orchestratorTools: any[];
  pageBuilderTools: any[];
  gateFixerTools: any[];
  state: { lastDocxPath?: string };
}

export function makeTools(pdfPath: string, workDir: string, cfg: EngineConfig, log: Log = console.log): EngineTools {
  const state: EngineTools["state"] = {};

  // ---- orchestrator tools ---------------------------------------------------

  const deterministicPassTool = tool(
    async () => {
      try {
        const r = await deterministicPass(pdfPath, cfg, log);
        return asJson({
          totalPages: r.totalPages,
          deterministicDone: r.deterministicDone,
          complexPages: r.complexPages,
          failed: r.failed,
        });
      } catch (e) {
        return errJson(e);
      }
    },
    {
      name: "deterministic_pass",
      description:
        "Process ALL pages of the input PDF deterministically (zero LLM). Clean pages complete immediately; returns the worklist of complex pages ({page, reason, hasFills, hasFrame, assetFiles}) plus pages whose extraction failed. Resume-safe: completed pages are skipped.",
      schema: z.object({}),
    }
  );

  const getSessionTool = tool(
    async () => {
      try {
        return asJson(sessionSummary(workDir));
      } catch (e) {
        return errJson(e);
      }
    },
    {
      name: "get_session",
      description: "Compact per-page status summary: {totalPages, pages:[{page, status, mode, reason, error?}]}.",
      schema: z.object({}),
    }
  );

  const mergeDocumentTool = tool(
    async ({ outName }: { outName?: string }) => {
      try {
        const safe = basename(outName || cfg.outName || basename(workDir)).replace(/\.docx$/i, "");
        const r = await mergeDocument(workDir, safe, cfg);
        state.lastDocxPath = r.docxPath;
        log(`Merged ${r.mergedPages} page(s) → ${r.docxPath}`);
        return asJson(r);
      } catch (e) {
        return errJson(e);
      }
    },
    {
      name: "merge_document",
      description:
        "Order + merge every completed page manifest into outputs/<outName>.docx. Returns {docxPath, mergedPages, missingPages}. missingPages lists pages that are not completed and were skipped.",
      schema: z.object({ outName: z.string().optional().describe("output file stem (no path, .docx added)") }),
    }
  );

  const runGateTool = tool(
    async () => {
      if (!state.lastDocxPath) return asJson({ ok: false, error: "no merged document yet — call merge_document first" });
      try {
        const r = runGateOn(state.lastDocxPath, workDir, cfg);
        log(`Gate: ${r.ok ? "PASS" : "ISSUES"} — ${r.checks.map((c) => `${c.ok ? "✓" : "✗"} ${c.name}`).join(", ")}`);
        return asJson(r);
      } catch (e) {
        return errJson(e);
      }
    },
    {
      name: "run_gate",
      description:
        "Verify the merged .docx (valid OOXML, injected fills landed, tables fit, word coverage, images embedded). Returns {ok, checks:[{name,ok,detail}], counts, perPageHints:[{page,hint}]} — hints tell a fixer WHERE to look.",
      schema: z.object({}),
    }
  );

  const markPageImageFallbackTool = tool(
    async ({ page, reason }: { page: number; reason: string }) => {
      try {
        const r = markPageImageFallback(workDir, page, reason, cfg);
        log(`  ✓ page ${page}  [image fallback: ${reason}]`);
        return asJson(r);
      } catch (e) {
        return errJson(e);
      }
    },
    {
      name: "mark_page_image_fallback",
      description:
        "The ladder's floor: complete a page as its own full-page picture plus fallback text. Use only after page_builder and vision_reconstruct have failed (or under --llm none).",
      schema: z.object({ page: pageArg, reason: z.string().describe("why the page fell back to an image") }),
    }
  );

  const visionReconstructTool = tool(
    async ({ page }: { page: number }) => {
      try {
        return asJson(await visionReconstruct(workDir, cfg, page, log));
      } catch (e) {
        return errJson(e);
      }
    },
    {
      name: "vision_reconstruct",
      description:
        "Rebuild ONE complex page from its rendered image using the vision model (code-driven retry loop inside the tool). Returns {ok, coverage, errors, model?}; on ok:true the page manifest is already written and the page is completed.",
      schema: z.object({ page: pageArg }),
    }
  );

  // ---- subagent tools (page_builder / gate_fixer) ---------------------------

  const readPageInputTool = tool(
    async ({ page }: { page: number }) => {
      try {
        return asJson(readPageInput(workDir, page));
      } catch (e) {
        return errJson(e);
      }
    },
    {
      name: "read_page_input",
      description:
        "Read one complex page's pre-extracted inputs: {page, reason, textPrompt (text+coordinates), injected (fills/border/frame facts the host stamps), assetFiles, priorErrors?}. Call this FIRST.",
      schema: z.object({ page: pageArg }),
    }
  );

  const validateLayoutTool = tool(
    async ({ page, layout }: { page: number; layout: any }) => {
      try {
        const r = await validateLayoutForPage(workDir, cfg, page, layout);
        if (r.ok) log(`  ✓ page ${page}  [json]  coverage ${(100 * (r.coverage ?? 0)).toFixed(0)}%`);
        return asJson(r);
      } catch (e) {
        return errJson(e);
      }
    },
    {
      name: "validate_layout",
      description:
        "Render + structurally verify + fidelity-check a compact layout JSON for one page. On ok:true the page is COMPLETED (manifest written) — you are done. On ok:false fix the listed errors and call again.",
      schema: z.object({ page: pageArg, layout: layoutSchema }),
    }
  );

  const validateModuleTool = tool(
    async ({ page, source }: { page: number; source: string }) => {
      try {
        const r = await validateModuleForPage(workDir, cfg, page, source);
        if (r.ok) log(`  ✓ page ${page}  [llm codegen]`);
        return asJson(r);
      } catch (e) {
        return errJson(e);
      }
    },
    {
      name: "validate_module",
      description:
        "Codegen fallback rung: submit a docx.js ES-module source exporting buildPage(docx, ctx). It is executed and structurally verified; on ok:true the page is COMPLETED. Use at most once per page, only after validate_layout keeps failing.",
      schema: z.object({ page: pageArg, source: z.string().describe("full ES module source, no fences") }),
    }
  );

  return {
    orchestratorTools: [
      deterministicPassTool,
      getSessionTool,
      mergeDocumentTool,
      runGateTool,
      visionReconstructTool,
      markPageImageFallbackTool,
    ],
    pageBuilderTools: [readPageInputTool, validateLayoutTool, validateModuleTool],
    gateFixerTools: [getSessionTool, readPageInputTool, validateLayoutTool, validateModuleTool, markPageImageFallbackTool],
    state,
  };
}

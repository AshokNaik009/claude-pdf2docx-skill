// vision_reconstruct (spec §7.6): the vision rung as a CODE-DRIVEN loop inside
// one tool. Tiny free VL models are unreliable tool-callers and task() carries
// text, not images — so the orchestrator calls this tool and the loop
// (render → invoke VL model → lenient-parse → validateLayoutForPage → feed
// errors back) runs deterministically in-process.
import { join } from "node:path";
// @ts-ignore vendored JS
import { loadSession } from "../../vendor/src/session.js";
// @ts-ignore vendored JS
import { openPdf, pageSize, renderPage } from "../../vendor/src/pdf.js";
import { candidatesForRole, modelForRole, type ModelCandidate } from "../model-factory.js";
import { validateLayoutForPage, loadPageFiles, type ValidateResult } from "./validate.js";
import { LAYOUT_SCHEMA_TEXT } from "../prompts/page-builder.js";
import type { EngineConfig } from "../config.js";
import type { Log } from "./deterministic-pass.js";

/** Port of worker.js parseLayoutJson + one bracket-repair attempt (spec §7.6). */
export function parseLayoutLenient(raw: string): any {
  let t = String(raw || "").trim();
  const fence = t.match(/```[a-zA-Z]*\s*\n([\s\S]*?)```/);
  if (fence) t = fence[1].trim();
  const first = t.indexOf("{");
  if (first === -1) throw new Error("no JSON object in model output");
  const last = t.lastIndexOf("}");
  t = last > first ? t.slice(first, last + 1) : t.slice(first);
  let obj: any;
  try {
    obj = JSON.parse(t);
  } catch {
    obj = JSON.parse(repairBrackets(t)); // one repair attempt; throws if still broken
  }
  if (!obj || typeof obj !== "object") throw new Error("parsed JSON is not an object");
  if (!Array.isArray(obj.columns) && !Array.isArray(obj.blocks)) throw new Error("JSON has neither `columns` nor `blocks`");
  return obj;
}

/** Close whatever {[ ... the model left open (truncated output is the common VL failure). */
function repairBrackets(t: string): string {
  const stack: string[] = [];
  let inStr = false;
  for (let i = 0; i < t.length; i++) {
    const c = t[i];
    if (inStr) {
      if (c === "\\") i++;
      else if (c === '"') inStr = false;
    } else if (c === '"') inStr = true;
    else if (c === "{" || c === "[") stack.push(c);
    else if (c === "}" || c === "]") stack.pop();
  }
  let out = t.replace(/,\s*$/, "");
  if (inStr) out += '"';
  while (stack.length) out += stack.pop() === "{" ? "}" : "]";
  return out;
}

const VISION_SYSTEM = [
  "You reconstruct ONE document page as a COMPACT LAYOUT JSON by looking at the page image. NOT code, NOT prose.",
  "A deterministic host renders your JSON into an editable Word page and stamps the exact fills/borders/frame itself —",
  "transcribe the text EXACTLY as it appears and assemble blocks; never invent colors, fills, or a frame.",
  "Still mark a full-width colored heading band as a {\"t\":\"bar\",\"text\":...} block so the host can bind its fill.",
  "",
  LAYOUT_SCHEMA_TEXT,
  "",
  "Output ONLY the JSON object — start with { and end with }. No markdown fences, no commentary.",
].join("\n");

function visionInstructions(assetFiles: string[], priorErrors?: string[]): string {
  const assets = assetFiles.length ? assetFiles.map((f) => `"${f}"`).join(", ") : "(none)";
  const lines = [
    "Reconstruct this page as the compact layout JSON described in the system message.",
    "Transcribe ALL visible text exactly. Use `columns` ONLY for genuine side-by-side columns.",
    `IMAGES: the page's figures were already cropped to these asset files: ${assets}.`,
    "Place each provided asset once as {\"t\":\"img\",\"file\":\"...\"} where it appears; do NOT invent filenames.",
  ];
  if (priorErrors?.length) {
    lines.push("", "RETRY — your previous JSON failed verification. Fix these and do not repeat them:");
    lines.push(...priorErrors.map((e) => `  - ${e}`));
  }
  return lines.join("\n");
}

/**
 * Reconstruct one page via the vision role. Re-renders the page small enough
 * for VL payload limits, then loops model → parse → validate ≤ maxBuilderAttempts
 * error-feedback rounds; walks the role's candidate list if a provider dies.
 * On success the manifest is already written by validateLayoutForPage.
 */
export async function visionReconstruct(
  workDir: string,
  cfg: EngineConfig,
  page: number,
  log: Log = console.log
): Promise<ValidateResult & { model?: string }> {
  const f = loadPageFiles(workDir, page); // throws a clear error if the page has no LLM inputs
  const session = loadSession(workDir);
  const i = page - 1;

  const doc = openPdf(session.document);
  const widthPts = pageSize(doc, i).width;
  const dpi = Math.max(72, Math.min(cfg.dpi, Math.floor((1400 * 72) / widthPts))); // width ≤ ~1400px
  const render = renderPage(doc, i, dpi);
  const dataUrl = `data:image/png;base64,${render.fullPng.toString("base64")}`;
  const assetFiles: string[] = f.meta.assetFiles || [];

  const candidates = candidatesForRole("vision");
  let lastErrors: string[] = ["vision produced no attempt"];
  let lastCoverage: number | undefined;

  for (let ci = 0; ci < candidates.length; ci++) {
    const cand: ModelCandidate = candidates[ci];
    const modelId = `${cand.provider}:${cand.model}`;
    let model;
    try {
      model = modelForRole("vision", ci);
    } catch (e) {
      log(`  vision: candidate ${modelId} unavailable — ${(e as Error).message}`);
      continue;
    }

    let priorErrors: string[] | undefined;
    for (let attempt = 0; attempt <= cfg.maxBuilderAttempts; attempt++) {
      let raw: string;
      try {
        const res = await model.invoke([
          { role: "system", content: VISION_SYSTEM },
          {
            role: "user",
            content: [
              { type: "image_url", image_url: { url: dataUrl } },
              { type: "text", text: visionInstructions(assetFiles, priorErrors) },
            ],
          },
        ]);
        raw = typeof res.content === "string"
          ? res.content
          : (res.content as any[]).map((c) => (typeof c === "string" ? c : c.text ?? "")).join("");
      } catch (e) {
        // Provider-level failure (429 exhausted, retired model) — next candidate.
        log(`  vision: ${modelId} invoke failed — ${(e as Error).message}`);
        lastErrors = [`vision model ${modelId} failed: ${(e as Error).message}`];
        break;
      }

      let layout: any;
      try {
        layout = parseLayoutLenient(raw);
      } catch (e) {
        priorErrors = [`your output was not a parseable layout JSON: ${(e as Error).message}`];
        lastErrors = priorErrors;
        log(`  vision: page ${page} attempt ${attempt + 1} (${modelId}) unparseable — ${(e as Error).message}`);
        continue;
      }

      const result = await validateLayoutForPage(workDir, cfg, page, layout, { model: modelId, llmInput: "image" });
      if (result.ok) {
        log(`  ✓ page ${page}  [json via vision ${modelId}]  coverage ${(100 * (result.coverage ?? 0)).toFixed(0)}%`);
        return { ...result, model: modelId };
      }
      priorErrors = result.errors;
      lastErrors = result.errors;
      lastCoverage = result.coverage;
      log(`  vision: page ${page} attempt ${attempt + 1} (${modelId}) failed — ${result.errors.join("; ").slice(0, 200)}`);
    }
  }

  return { ok: false, coverage: lastCoverage, errors: lastErrors };
}

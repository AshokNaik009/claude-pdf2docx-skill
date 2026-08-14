// Orchestration (SESSION STATE + PER-PAGE PROCESSING, parallel) for the
// deterministic-fast engine. Per page:
//   render → quality-check → EITHER
//     (a) build blocks deterministically from PDF geometry (clean page, 0 LLM), OR
//     (b) route to the model, which emits a COMPACT LAYOUT JSON (~2× faster than
//         emitting docx.js code); a deterministic renderer expands it and STAMPS
//         the fills/borders/frame extracted from the PDF's vector layer.
//   The compact-JSON attempt falls back to docx.js code-gen, then to the page
//   image, so it never ships a worse result than the legacy path.
// Writes a per-page manifest + assets and tracks status in session.json. Resumable.
import { mkdirSync, writeFileSync, existsSync } from "node:fs";
import { join, resolve, basename } from "node:path";
import { openPdf, pageCount, pageStructured, pageStructuredPy, pageSize, pageText, renderPage } from "./pdf.js";
import { classifyPage } from "./quality.js";
import { extractImageAssets, extractFigureRegions, buildPageSections } from "./extract.js";
import { generatePageCode, generatePageCodeFromText, generatePageJson } from "./worker.js";
import { executeModule, renderJsonSections, renderLayoutJson, makeCtx } from "./docx.js";
import { verifySections, coverage } from "./verify.js";
import { STATUS, sessionPath, loadSession, saveSession } from "./session.js";

const pad = (n) => String(n).padStart(4, "0");
const WORD = /[\p{L}\p{N}]{3,}/gu;

export function workDirFor(pdfPath, cfg) {
  const stem = basename(pdfPath).replace(/\.pdf$/i, "");
  return resolve(cfg.workRoot, stem);
}

export function initSession(pdfPath, cfg) {
  const workDir = workDirFor(pdfPath, cfg);
  mkdirSync(join(workDir, "pages"), { recursive: true });
  mkdirSync(join(workDir, "assets"), { recursive: true });

  if (existsSync(sessionPath(workDir))) {
    const s = loadSession(workDir);
    s.resumed = true;
    return { workDir, session: s };
  }

  const doc = openPdf(pdfPath);
  const total = pageCount(doc);
  const session = {
    sessionId: `${basename(workDir)}-${Date.now().toString(36)}`,
    document: resolve(pdfPath),
    totalPages: total,
    model: cfg.model,
    dpi: cfg.dpi,
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    pages: Array.from({ length: total }, (_, i) => ({
      page: i + 1,
      status: STATUS.PENDING,
      mode: null,
      reason: null,
      chars: 0,
      tokens: 0,
      processingTimeMs: 0,
      retryCount: 0,
      manifest: `pages/page_${pad(i + 1)}.json`,
      assets: `assets/page_${pad(i + 1)}`,
      error: null,
    })),
  };
  saveSession(workDir, session);
  return { workDir, session };
}

/** Plain-text fallback so a broken generated module still yields readable text. */
export function fallbackParagraphs(doc, i) {
  return pageText(doc, i)
    .split(/\n\s*\n/)
    .map((s) => s.replace(/\s+/g, " ").trim())
    .filter(Boolean);
}

/** The exact facts we STAMP into the render — never left for the model to decide. */
export function buildInjected(drawings) {
  if (!drawings) return {};
  const fills = (drawings.fills || []).filter((f) => f && f.fill);
  const border = (drawings.borders || [])[0] ? { color: drawings.borders[0].color, width: drawings.borders[0].width } : null;
  const frame = drawings.frame ? { color: drawings.frame.color, width: drawings.frame.width } : null;
  const out = {};
  if (fills.length) out.fills = fills;
  if (border) out.border = border;
  if (frame) out.frame = frame;
  return out;
}

/** Flatten a compact layout JSON to plain text for the coverage fidelity check. */
export function layoutText(layout) {
  const parts = [];
  if (layout.title) parts.push(layout.title);
  const cols = layout.columns || [{ blocks: layout.blocks || [] }];
  for (const col of cols) {
    for (const b of col.blocks || []) {
      if (b.text) parts.push(b.text);
      if (b.items) parts.push(...b.items);
      if (b.rows) for (const row of b.rows) for (const cell of row) parts.push(typeof cell === "object" ? (cell?.text ?? "") : String(cell ?? ""));
    }
  }
  return parts.join(" ");
}

export function coverageOfText(originalText, reconstructed, cfg) {
  const orig = new Set((originalText || "").toLowerCase().match(WORD) || []);
  if (!orig.size) return 1;
  const got = new Set((reconstructed || "").toLowerCase().match(WORD) || []);
  let hits = 0;
  for (const w of orig) if (got.has(w)) hits++;
  return hits / orig.size;
}

/**
 * Guarantee no extracted figure is silently dropped: any asset the model didn't
 * reference gets appended as an `img` block (in reading order, in the last
 * column). Placement may be imperfect, but the image is never lost.
 */
export function ensureImages(layout, assetFiles) {
  if (!assetFiles || !assetFiles.length) return layout;
  const cols = (layout.columns && layout.columns.length) ? layout.columns : null;
  const referenced = new Set();
  const scan = (bs) => { for (const b of bs || []) if (b && b.t === "img" && b.file) referenced.add(String(b.file).split("/").pop()); };
  if (cols) for (const c of cols) scan(c.blocks); else scan(layout.blocks);
  const missing = assetFiles.filter((f) => !referenced.has(f));
  if (!missing.length) return layout;
  const imgBlocks = missing.map((f) => ({ t: "img", file: f }));
  if (cols) { const last = cols[cols.length - 1]; last.blocks = [...(last.blocks || []), ...imgBlocks]; }
  else { layout.blocks = [...(layout.blocks || []), ...imgBlocks]; }
  return layout;
}

/**
 * One compact-JSON attempt: generate the layout JSON, ensure every extracted
 * image is present, render it with the injected facts, structurally verify, and
 * check word fidelity.
 */
async function jsonOnce(source, injected, assetFiles, cfg, model, pageDir, origText, priorAttempt) {
  const raw = await generatePageJson(source, { model, timeoutMs: cfg.claudeTimeoutMs, assetFiles, priorAttempt });
  const layout = ensureImages(raw, assetFiles);
  const { sections, styles } = renderLayoutJson(layout, makeCtx(pageDir), injected);
  const verified = await verifySections({ sections, styles });
  const cov = coverageOfText(origText, layoutText(layout), cfg);
  if (verified.ok && cov < (cfg.coverageMin ?? 0.6)) {
    verified.ok = false;
    verified.errors = [...verified.errors, `low fidelity: only ${(cov * 100).toFixed(0)}% of the page's words reconstructed`];
  }
  return { layout, verified, cov };
}

/** Compact-JSON with the same retry→escalation ladder as the code path. */
async function jsonWithEscalation(source, injected, assetFiles, cfg, pageDir, origText) {
  let res, priorAttempt;
  for (let attempt = 0; attempt <= cfg.verifyRetries; attempt++) {
    res = await jsonOnce(source, injected, assetFiles, cfg, cfg.model, pageDir, origText, priorAttempt);
    if (res.verified.ok) return { ...res, usedModel: cfg.model };
    priorAttempt = { errors: res.verified.errors };
  }
  if (cfg.escalationModel && cfg.escalationModel !== cfg.model) {
    res = await jsonOnce(source, injected, assetFiles, cfg, cfg.escalationModel, pageDir, origText, priorAttempt);
    return { ...res, usedModel: cfg.escalationModel };
  }
  return { ...res, usedModel: cfg.model };
}

// ---- docx.js code-gen fallback (the legacy path) ----------------------------

async function generateOnce(pagePng, textPrompt, assetFiles, cfg, model, codeAbsPath, pageDir, priorAttempt) {
  const opts = { assetFiles, model, timeoutMs: cfg.claudeTimeoutMs, priorAttempt };
  const code = textPrompt
    ? await generatePageCodeFromText(textPrompt, opts)
    : await generatePageCode(resolve(pagePng), opts);
  writeFileSync(codeAbsPath, code);
  const executed = await executeModule(codeAbsPath, pageDir);
  const verified = await verifySections(executed);
  return { code, verified };
}

async function generateWithEscalation(pagePng, textPrompt, assetFiles, cfg, codeAbsPath, pageDir) {
  let code, verified, priorAttempt;
  for (let attempt = 0; attempt <= cfg.verifyRetries; attempt++) {
    ({ code, verified } = await generateOnce(pagePng, textPrompt, assetFiles, cfg, cfg.model, codeAbsPath, pageDir, priorAttempt));
    if (verified.ok) return { code, verified, usedModel: cfg.model };
    priorAttempt = { code, errors: verified.errors };
  }
  if (cfg.escalationModel && cfg.escalationModel !== cfg.model) {
    ({ code, verified } = await generateOnce(pagePng, textPrompt, assetFiles, cfg, cfg.escalationModel, codeAbsPath, pageDir, priorAttempt));
    return { code, verified, usedModel: cfg.escalationModel };
  }
  return { code, verified, usedModel: cfg.model };
}

async function processPage(doc, pdfPath, workDir, pageState, cfg) {
  const i = pageState.page - 1;
  const started = Date.now();
  const relDir = pageState.assets;
  const pageDir = join(workDir, relDir);
  mkdirSync(pageDir, { recursive: true });

  // Source structured text + page size + (with pyExtract) the compact text
  // prompt and the vector DRAWINGS (fills/borders/frame) from pdfextract.py.
  const { structured, width, textPrompt, drawings } = cfg.pyExtract
    ? pageStructuredPy(pdfPath, i, cfg)
    : { structured: pageStructured(doc, i), width: pageSize(doc, i).width, textPrompt: null, drawings: null };
  const render = renderPage(doc, i, cfg.dpi);
  const injected = buildInjected(drawings);

  const cls = classifyPage(structured, width, cfg);
  const imageAssets = extractImageAssets(render, structured, pageDir, relDir);

  const figureRegions = extractFigureRegions(structured, render, pageDir, relDir, imageAssets, width);
  const augmentedBlocks = [...(structured.blocks || [])];
  for (const fig of figureRegions) {
    imageAssets.set(augmentedBlocks.length, fig.asset);
    augmentedBlocks.push({ type: "image", bbox: fig.bbox });
  }
  const structuredWithFigures = figureRegions.length ? { ...structured, blocks: augmentedBlocks } : structured;
  const assetFiles = [...imageAssets.values()].map((a) => basename(a.path));

  let manifest, tokens = 0, mode = cls.mode, reason = cls.reason, escalatedFrom = null;

  // A page carrying vector fills or an outer frame can only be reproduced by the
  // injection path (the deterministic text builder drops color) — so route it to
  // the compact-JSON path even if its text layer is "clean". --llm none forbids
  // this; without --compact-json we keep legacy routing.
  if (mode === "text" && cfg.compactJson && cfg.llmMode !== "none" && cfg.llmMode !== "all" &&
      (injected.fills?.length || injected.frame)) {
    mode = "llm";
    reason = `vector fills/frame present → inject via compact-JSON (${reason})`;
  }

  if (mode === "text") {
    // Deterministic path — no LLM. Verified + fidelity-checked for free; a
    // failure escalates to the model instead of shipping a broken page.
    const { sections: jsonSections } = buildPageSections(structuredWithFigures, imageAssets, cfg);
    const allBlocks = jsonSections.flatMap((s) => s.blocks);
    const liveSections = renderJsonSections(jsonSections, makeCtx(pageDir));
    const verified = await verifySections({ sections: liveSections });
    const cov = verified.ok ? coverage(pageText(doc, i), allBlocks, cfg) : { ok: false, reasons: [] };

    const failureReason = !verified.ok
      ? `deterministic build failed verification: ${verified.errors.join("; ")}`
      : !cov.ok
      ? `deterministic build low fidelity: ${cov.reasons.join("; ")}`
      : null;

    if (!failureReason || cfg.llmMode === "none") {
      manifest = {
        page: pageState.page,
        mode: "text",
        reason: failureReason ? `${reason} (unverified: ${failureReason}; --llm none forbids escalation)` : reason,
        sections: jsonSections,
        fallbackText: fallbackParagraphs(doc, i),
        assets: [...imageAssets.values()].map((a) => a.path),
        assetsDir: relDir,
      };
    } else {
      escalatedFrom = "text";
      reason = failureReason;
      mode = "llm";
    }
  }

  if (mode === "llm") {
    const pagePng = join(pageDir, "page.png");
    writeFileSync(pagePng, render.fullPng);
    const pagesRel = relDir.replace(/^assets/, "pages");
    mkdirSync(join(workDir, pagesRel), { recursive: true });
    const origText = pageText(doc, i);

    let done = false;

    // (1) COMPACT-JSON — the fast path. From py-extract text when available,
    //     else from the image. Renders with injected fills/borders/frame.
    if (cfg.compactJson) {
      const source = (cfg.pyExtract && textPrompt) ? { textPrompt } : { imageAbsPath: pagePng };
      try {
        const r = await jsonWithEscalation(source, injected, assetFiles, cfg, pageDir, origText);
        if (!r.verified.ok) throw new Error(`compact-JSON unverified: ${r.verified.errors.join("; ")}`);
        writeFileSync(join(workDir, pagesRel, "layout.json"), JSON.stringify(r.layout, null, 2));
        tokens = Math.round(JSON.stringify(r.layout).length / 4);
        manifest = {
          page: pageState.page, mode: "json",
          llmInput: source.textPrompt ? "text" : "image",
          reason, escalatedFrom, model: r.usedModel,
          layout: r.layout, injected,
          fallbackText: fallbackParagraphs(doc, i),
          assets: [...imageAssets.values()].map((a) => a.path),
          assetsDir: relDir,
        };
        done = true;
      } catch (_) {
        escalatedFrom = escalatedFrom || "compact-json";
      }
    }

    // (2) docx.js CODE-GEN fallback — text-only first (if py-extract), then image.
    if (!done) {
      const codeRel = join(pagesRel, "page.mjs");
      const codeAbsPath = join(workDir, codeRel);
      const useText = cfg.pyExtract && !!textPrompt;
      let result, llmInput = useText ? "text" : "image";
      if (useText) {
        try {
          result = await generateWithEscalation(pagePng, textPrompt, assetFiles, cfg, codeAbsPath, pageDir);
          if (!result.verified.ok) throw new Error(`text-only unverified: ${result.verified.errors.join("; ")}`);
        } catch (_) {
          result = await generateWithEscalation(pagePng, null, assetFiles, cfg, codeAbsPath, pageDir);
          llmInput = "image";
        }
      } else {
        result = await generateWithEscalation(pagePng, null, assetFiles, cfg, codeAbsPath, pageDir);
      }
      if (!result.verified.ok) {
        throw new Error(`page failed verification with ${result.usedModel}: ${result.verified.errors.join("; ")}`);
      }
      tokens = Math.round(result.code.length / 4);
      manifest = {
        page: pageState.page, mode: "llm", llmInput,
        reason, escalatedFrom, model: result.usedModel,
        codeFile: codeRel,
        fallbackText: fallbackParagraphs(doc, i),
        assets: [...imageAssets.values()].map((a) => a.path),
        assetsDir: relDir,
      };
    }
  }

  writeFileSync(join(workDir, pageState.manifest), JSON.stringify(manifest, null, 2));
  const chars = pageText(doc, i).length;
  return { mode: manifest.mode, reason, chars, tokens, ms: Date.now() - started, escalated: !!escalatedFrom };
}

export async function runPipeline(pdfPath, cfg, log = console.log) {
  const { workDir, session } = initSession(pdfPath, cfg);
  const doc = openPdf(pdfPath);

  const todo = session.pages.filter((p) => p.status !== STATUS.COMPLETED);
  log(
    `Session ${session.sessionId} — ${session.totalPages} pages, ${todo.length} to process` +
      (session.resumed ? " (resumed)" : "")
  );

  const queue = [...todo];
  let active = 0;

  await new Promise((resolveAll) => {
    const pump = () => {
      if (queue.length === 0 && active === 0) return resolveAll();
      while (active < cfg.concurrency && queue.length) {
        const pageState = queue.shift();
        active++;
        pageState.status = STATUS.IN_PROGRESS;
        saveSession(workDir, session);

        processPage(doc, pdfPath, workDir, pageState, cfg)
          .then((r) => {
            pageState.status = STATUS.COMPLETED;
            pageState.mode = r.mode;
            pageState.reason = r.reason;
            pageState.chars = r.chars;
            pageState.tokens = r.tokens;
            pageState.processingTimeMs = r.ms;
            pageState.escalated = r.escalated;
            pageState.error = null;
            const tag = r.escalated ? `${r.mode}, escalated from ${r.mode === "json" ? "text" : "text/json"}` : r.mode;
            log(`  ✓ page ${pageState.page}  [${tag}: ${r.reason}]  ${r.ms}ms`);
          })
          .catch((e) => {
            pageState.retryCount++;
            if (pageState.retryCount <= cfg.retries) {
              pageState.status = STATUS.PENDING;
              queue.push(pageState);
              log(`  ↻ page ${pageState.page} failed (retry ${pageState.retryCount}): ${e.message}`);
            } else {
              pageState.status = STATUS.FAILED;
              pageState.error = e.message;
              log(`  ✗ page ${pageState.page} FAILED: ${e.message}`);
            }
          })
          .finally(() => {
            active--;
            saveSession(workDir, session);
            pump();
          });
      }
    };
    pump();
  });

  const done = session.pages.filter((p) => p.status === STATUS.COMPLETED).length;
  const failed = session.pages.filter((p) => p.status === STATUS.FAILED).length;
  const escalated = session.pages.filter((p) => p.escalated).length;
  log(`Done: ${done}/${session.totalPages} completed${failed ? `, ${failed} FAILED` : ""}` +
      `${escalated ? `, ${escalated} escalated` : ""}.`);
  return { workDir, session };
}

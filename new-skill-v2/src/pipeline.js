// Orchestration (SESSION STATE + PER-PAGE PROCESSING, parallel).
// Per page: render, quality-check, then EITHER deterministically build blocks
// (clean text page) OR have headless Claude GENERATE the docx-building module
// from the page image (complex page). Writes a page manifest + assets, and
// tracks status/timing/tokens/retries in session.json. Resumable.
import { mkdirSync, writeFileSync, existsSync } from "node:fs";
import { join, resolve, basename } from "node:path";
import { openPdf, pageCount, pageStructured, pageSize, pageText, renderPage } from "./pdf.js";
import { classifyPage } from "./quality.js";
import { extractImageAssets, extractFigureRegions, buildPageSections } from "./extract.js";
import { generatePageCode, generatePageMarkdown, generatePageHtml } from "./worker.js";
import { sanitizeGfm } from "./mddocx.js";
import { executeModule, renderJsonSections, makeCtx } from "./docx.js";
import { verifySections, coverage } from "./verify.js";
import { STATUS, sessionPath, loadSession, saveSession } from "./session.js";

const pad = (n) => String(n).padStart(4, "0");

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
function fallbackParagraphs(doc, i) {
  return pageText(doc, i)
    .split(/\n\s*\n/)
    .map((s) => s.replace(/\s+/g, " ").trim())
    .filter(Boolean);
}

/** One generation attempt: generate, execute, structurally verify. */
async function generateOnce(pagePng, assetFiles, cfg, model, codeAbsPath, pageDir, priorAttempt) {
  const code = await generatePageCode(resolve(pagePng), {
    assetFiles, model, timeoutMs: cfg.claudeTimeoutMs, priorAttempt,
  });
  writeFileSync(codeAbsPath, code);
  const executed = await executeModule(codeAbsPath, pageDir);
  const verified = await verifySections(executed);
  return { code, verified };
}

/**
 * cfg.model (Sonnet) for cfg.verifyRetries+1 attempts, feeding each failure's
 * errors back into the next prompt; if still failing, ONE final attempt with
 * cfg.escalationModel (Opus) — reserved for pages neither the deterministic
 * builder nor the default vision model can get right.
 */
async function generateWithEscalation(pagePng, assetFiles, cfg, codeAbsPath, pageDir) {
  let code, verified, priorAttempt;
  for (let attempt = 0; attempt <= cfg.verifyRetries; attempt++) {
    ({ code, verified } = await generateOnce(pagePng, assetFiles, cfg, cfg.model, codeAbsPath, pageDir, priorAttempt));
    if (verified.ok) return { code, verified, usedModel: cfg.model };
    priorAttempt = { code, errors: verified.errors };
  }
  if (cfg.escalationModel && cfg.escalationModel !== cfg.model) {
    ({ code, verified } = await generateOnce(pagePng, assetFiles, cfg, cfg.escalationModel, codeAbsPath, pageDir, priorAttempt));
    return { code, verified, usedModel: cfg.escalationModel };
  }
  return { code, verified, usedModel: cfg.model };
}

async function processPage(doc, workDir, pageState, cfg) {
  const i = pageState.page - 1;
  const started = Date.now();
  const relDir = pageState.assets;
  const pageDir = join(workDir, relDir);
  mkdirSync(pageDir, { recursive: true });

  const structured = pageStructured(doc, i);
  const { width } = pageSize(doc, i);
  const render = renderPage(doc, i, cfg.dpi);

  const cls = classifyPage(structured, width, cfg);
  const imageAssets = extractImageAssets(render, structured, pageDir, relDir);

  // Vector-drawn diagrams (arrows/boxes, not an embedded raster) produce no
  // "image" block at all in mupdf's structured text, so extractImageAssets
  // never sees them. Crop the gap between a "Fig. N" caption and whatever
  // precedes it from the page raster and treat it as one more image asset —
  // benefits both paths (a real embeddable asset instead of a placeholder).
  const figureRegions = extractFigureRegions(structured, render, pageDir, relDir, imageAssets, width);
  const augmentedBlocks = [...(structured.blocks || [])];
  for (const fig of figureRegions) {
    imageAssets.set(augmentedBlocks.length, fig.asset);
    augmentedBlocks.push({ type: "image", bbox: fig.bbox });
  }
  const structuredWithFigures = figureRegions.length ? { ...structured, blocks: augmentedBlocks } : structured;

  const assetFiles = [...imageAssets.values()].map((a) => basename(a.path));

  let manifest, tokens = 0, mode = cls.mode, reason = cls.reason, escalatedFrom = null;

  if (mode === "text") {
    // Deterministic path — tables, multi-column layout, and clean prose,
    // all reconstructed from PDF geometry with no LLM call. Verified for
    // free (src/verify.js: structural pack test + word-coverage fidelity
    // check) before being accepted; a failure here escalates to vision
    // instead of shipping a broken or badly-garbled page.
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

    // --llm none is an absolute override: never call Claude, even as a
    // fallback. Ship the best-effort deterministic result rather than the
    // escalation an unverified/low-fidelity page would otherwise trigger.
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

  if (mode === "llm" && cfg.pandoc) {
    // Pandoc HTML strategy — Claude reconstructs the page as semantic HTML
    // (multi-column layouts as <table>); the merge step assembles the whole
    // document and runs pandoc on it.
    const pagePng = join(pageDir, "page.png");
    writeFileSync(pagePng, render.fullPng);
    const htmlRel = join(relDir.replace(/^assets/, "pages"), `page.html`); // pages/page_NNNN/page.html
    mkdirSync(join(workDir, relDir.replace(/^assets/, "pages")), { recursive: true });
    const html = await generatePageHtml(pagePng, {
      assetFiles, model: cfg.model, timeoutMs: cfg.claudeTimeoutMs,
    });
    writeFileSync(join(workDir, htmlRel), html);
    tokens = Math.round(html.length / 4);
    manifest = {
      page: pageState.page,
      mode: "html",
      reason,
      escalatedFrom,
      model: cfg.model,
      htmlFile: htmlRel,
      fallbackText: fallbackParagraphs(doc, i),
      assets: [...imageAssets.values()].map((a) => a.path),
      assetsDir: relDir,
    };
  } else if (mode === "llm" && cfg.markdown) {
    // Markdown strategy — Claude reconstructs the page as pure GFM Markdown; the
    // merge step assembles the whole document and runs markdown-docx on it.
    const pagePng = join(pageDir, "page.png");
    writeFileSync(pagePng, render.fullPng);
    const mdRel = join(relDir.replace(/^assets/, "pages"), `page.md`); // pages/page_NNNN/page.md
    mkdirSync(join(workDir, relDir.replace(/^assets/, "pages")), { recursive: true });
    const raw = await generatePageMarkdown(pagePng, {
      assetFiles, model: cfg.model, timeoutMs: cfg.claudeTimeoutMs,
    });
    const md = sanitizeGfm(raw);
    if (!md) throw new Error("markdown generation produced empty content after sanitizing");
    writeFileSync(join(workDir, mdRel), md);
    tokens = Math.round(md.length / 4);
    manifest = {
      page: pageState.page,
      mode: "md",
      reason,
      escalatedFrom,
      model: cfg.model,
      markdownFile: mdRel,
      fallbackText: fallbackParagraphs(doc, i),
      assets: [...imageAssets.values()].map((a) => a.path),
      assetsDir: relDir,
    };
  } else if (mode === "llm") {
    // LLM path — Claude generates the docx-building module from the page image.
    const pagePng = join(pageDir, "page.png");
    writeFileSync(pagePng, render.fullPng);
    const codeRel = join(relDir.replace(/^assets/, "pages"), `page.mjs`); // pages/page_NNNN/page.mjs
    const codeAbsDir = join(workDir, relDir.replace(/^assets/, "pages"));
    mkdirSync(codeAbsDir, { recursive: true });
    const codeAbsPath = join(workDir, codeRel);

    const { code, verified, usedModel } = await generateWithEscalation(pagePng, assetFiles, cfg, codeAbsPath, pageDir);
    if (!verified.ok) {
      throw new Error(`page failed verification with ${usedModel}: ${verified.errors.join("; ")}`);
    }
    tokens = Math.round(code.length / 4);
    manifest = {
      page: pageState.page,
      mode: "llm",
      reason,
      escalatedFrom,
      model: usedModel,
      codeFile: codeRel,
      fallbackText: fallbackParagraphs(doc, i),
      assets: [...imageAssets.values()].map((a) => a.path),
      assetsDir: relDir,
    };
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

        processPage(doc, workDir, pageState, cfg)
          .then((r) => {
            pageState.status = STATUS.COMPLETED;
            pageState.mode = r.mode;
            pageState.reason = r.reason;
            pageState.chars = r.chars;
            pageState.tokens = r.tokens;
            pageState.processingTimeMs = r.ms;
            pageState.escalated = r.escalated;
            pageState.error = null;
            const tag = r.escalated ? `${r.mode}, escalated from text` : r.mode;
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
      `${escalated ? `, ${escalated} escalated from deterministic to vision` : ""}.`);
  return { workDir, session };
}

// Orchestration (SESSION STATE + PER-PAGE PROCESSING, parallel).
// Per page: render, quality-check, then EITHER deterministically build blocks
// (clean text page) OR have headless Claude GENERATE the docx-building module
// from the page image (complex page). Writes a page manifest + assets, and
// tracks status/timing/tokens/retries in session.json. Resumable.
import { mkdirSync, writeFileSync, existsSync } from "node:fs";
import { join, resolve, basename } from "node:path";
import { openPdf, pageCount, pageStructured, pageSize, pageText, renderPage } from "./pdf.js";
import { classifyPage } from "./quality.js";
import { extractImageAssets, buildTextBlocks } from "./extract.js";
import { generatePageCode } from "./worker.js";
import { executeModule } from "./docx.js";
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
  const assetFiles = [...imageAssets.values()].map((a) => basename(a.path));

  let manifest;
  let tokens = 0;

  if (cls.mode === "text") {
    // Deterministic path — clean prose, no LLM.
    manifest = {
      page: pageState.page,
      mode: "text",
      reason: cls.reason,
      blocks: buildTextBlocks(structured, imageAssets, cfg),
      assets: [...imageAssets.values()].map((a) => a.path),
    };
  } else {
    // LLM path — Claude generates the docx-building module from the page image.
    const pagePng = join(pageDir, "page.png");
    writeFileSync(pagePng, render.fullPng);
    const code = await generatePageCode(resolve(pagePng), {
      assetFiles,
      model: cfg.model,
      timeoutMs: cfg.claudeTimeoutMs,
    });
    const codeRel = join(relDir.replace(/^assets/, "pages"), `page.mjs`); // pages/page_NNNN/page.mjs
    const codeAbsDir = join(workDir, relDir.replace(/^assets/, "pages"));
    mkdirSync(codeAbsDir, { recursive: true });
    writeFileSync(join(workDir, codeRel), code);
    // Validate by executing: a runtime error here triggers a pipeline retry
    // (regenerate) instead of silently degrading to fallback at merge time.
    await executeModule(join(workDir, codeRel), pageDir);
    tokens = Math.round(code.length / 4);
    manifest = {
      page: pageState.page,
      mode: "llm",
      reason: cls.reason,
      codeFile: codeRel,
      fallbackText: fallbackParagraphs(doc, i),
      assets: [...imageAssets.values()].map((a) => a.path),
      assetsDir: relDir,
    };
  }

  writeFileSync(join(workDir, pageState.manifest), JSON.stringify(manifest, null, 2));
  const chars = pageText(doc, i).length;
  return { mode: cls.mode, reason: cls.reason, chars, tokens, ms: Date.now() - started };
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
            pageState.error = null;
            log(`  ✓ page ${pageState.page}  [${r.mode}: ${r.reason}]  ${r.ms}ms`);
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
  log(`Done: ${done}/${session.totalPages} completed${failed ? `, ${failed} FAILED` : ""}.`);
  return { workDir, session };
}

#!/usr/bin/env node
// PDF → structured model → editable DOCX.
//
//   node index.js run   <input.pdf>   # extract + per-page processing (writes work/<name>/)
//   node index.js merge <work/<name>> # order + merge + generate DOCX  (the "final merge")
//   node index.js all   <input.pdf>   # run then merge in one shot
//
// Flags (override config.js): --work DIR --out NAME --dpi N --concurrency N
//   --model M --retries N --llm auto|all|none
import { mkdirSync, writeFileSync, rmSync, existsSync } from "node:fs";
import { resolve, join, basename } from "node:path";
import { CONFIG } from "./config.js";
import { runPipeline, workDirFor } from "./src/pipeline.js";
import { buildDocumentModel, documentToMarkdown } from "./src/merge.js";
import { renderDocx } from "./src/docx.js";
import { buildMarkdownDocument, renderMarkdownDocx } from "./src/mddocx.js";
import { buildHtmlDocument, renderPandocDocx, pandocAvailable } from "./src/htmldocx.js";
import { runDoclingEngine, doclingAvailable } from "./src/doclingengine.js";

function parseFlags(argv) {
  const cfg = { ...CONFIG };
  const positional = [];
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    const next = () => argv[++i];
    switch (a) {
      case "--work": cfg.workRoot = next(); break;
      case "--out": cfg.outName = next(); break;
      case "--dpi": cfg.dpi = Number(next()); break;
      case "--concurrency": cfg.concurrency = Number(next()); break;
      case "--model": cfg.model = next(); break;
      case "--retries": cfg.retries = Number(next()); break;
      case "--llm": cfg.llmMode = next(); break;
      case "--markdown": cfg.markdown = true; break; // GFM → markdown-docx output strategy
      case "--pandoc": cfg.pandoc = true; break;     // semantic HTML → pandoc output strategy
      case "--docling": cfg.docling = true; break;   // Docling parse → generated docx.js modules (no LLM)
      case "--ref": cfg.pandocRefDoc = next(); break; // override pandoc reference .docx
      case "--fresh": cfg.fresh = true; break; // ignore any prior session, reprocess every page
      default: positional.push(a);
    }
  }
  return { cfg, positional };
}

async function doMerge(workDir, cfg) {
  const model = buildDocumentModel(workDir);
  if (model.missing.length) {
    console.warn(
      `⚠ ${model.missing.length} page(s) not completed and will be skipped: ` +
        model.missing.map((p) => `${p.page}(${p.status})`).join(", ")
    );
  }
  mkdirSync(resolve(cfg.outputsDir), { recursive: true });
  const stem = cfg.outName || basename(workDir);

  let buffer, srcPath;
  if (cfg.pandoc) {
    // Assembled semantic HTML is the source of truth; pandoc builds the .docx.
    const html = buildHtmlDocument(workDir, model);
    srcPath = resolve(cfg.outputsDir, `${stem}.html`);
    writeFileSync(srcPath, html);
    buffer = await renderPandocDocx(html, {
      referenceDoc: cfg.pandocRefDoc && resolve(cfg.pandocRefDoc),
      luaFilter: cfg.pandocLuaFilter && resolve(cfg.pandocLuaFilter),
    });
  } else {
    // In --markdown mode the assembled GFM IS the source of truth; otherwise
    // keep the best-effort intermediate markdown for reference.
    const markdown = cfg.markdown ? buildMarkdownDocument(workDir, model) : documentToMarkdown(model);
    srcPath = resolve(cfg.outputsDir, `${stem}.md`);
    writeFileSync(srcPath, markdown);
    buffer = cfg.markdown ? await renderMarkdownDocx(markdown) : await renderDocx(workDir, model, cfg);
  }
  const docxPath = resolve(cfg.outputsDir, `${stem}.docx`);
  writeFileSync(docxPath, buffer);

  const stats = model.pages.length;
  console.log(`\nMerged ${stats} page(s).`);
  console.log(`  ${cfg.pandoc ? "HTML:    " : "Markdown:"} ${srcPath}`);
  console.log(`  DOCX:     ${docxPath}`);
}

async function main() {
  const [, , cmdRaw, ...rest] = process.argv;
  const { cfg, positional } = parseFlags(rest);

  if (cfg.pandoc && !pandocAvailable()) {
    console.error("--pandoc requires the pandoc CLI, which is not on PATH.");
    console.error("Install it (e.g. `brew install pandoc`) and retry.");
    process.exit(1);
  }
  if ([cfg.pandoc, cfg.markdown, cfg.docling].filter(Boolean).length > 1) {
    console.error("--pandoc, --markdown and --docling are different output strategies; choose one.");
    process.exit(1);
  }
  if (cfg.docling && !doclingAvailable(cfg)) {
    console.error("--docling requires the 'docling' Python package in the project venv.");
    console.error("Set it up:  python3 -m venv .venv && .venv/bin/pip install docling");
    process.exit(1);
  }

  // Allow `node index.js <file.pdf>` as shorthand for `all`.
  let cmd = cmdRaw;
  let target = positional[0];
  if (cmdRaw && cmdRaw.toLowerCase().endsWith(".pdf")) { cmd = "all"; target = cmdRaw; }

  if (!cmd || !["run", "merge", "all"].includes(cmd) || !target) {
    console.error("Usage:");
    console.error("  node index.js run   <input.pdf>");
    console.error("  node index.js merge <work/<name>>");
    console.error("  node index.js all   <input.pdf>");
    process.exit(1);
  }

  if (cmd === "merge") {
    await doMerge(resolve(target), cfg);
    return;
  }

  const pdfPath = resolve(target);
  // Validate the input BEFORE any destructive --fresh cleanup, so a wrong path
  // never wipes a prior good run's work dir.
  if (!existsSync(pdfPath)) {
    console.error(`Input PDF not found: ${pdfPath}`);
    process.exit(1);
  }
  if (cfg.fresh) {
    rmSync(workDirFor(pdfPath, cfg), { recursive: true, force: true });
    console.log("(--fresh) cleared prior session; reprocessing all pages");
  }
  // --docling parses the whole PDF once and generates docx.js modules; the
  // per-page vision-LLM pipeline (runPipeline) is skipped entirely.
  const { workDir } = cfg.docling
    ? await runDoclingEngine(pdfPath, cfg)
    : await runPipeline(pdfPath, cfg);

  if (cmd === "all") {
    await doMerge(workDir, cfg);
  } else {
    console.log(`\nExtraction complete → ${workDir}`);
    console.log(`Run the final merge with:\n  node index.js merge ${workDir}`);
  }
}

main().catch((e) => {
  console.error("Fatal:", e.stack || e.message);
  process.exit(1);
});

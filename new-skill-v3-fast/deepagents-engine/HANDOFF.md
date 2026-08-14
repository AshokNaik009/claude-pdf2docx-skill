# HANDOFF — deepagents-engine implementation state

**Read `IMPLEMENTATION_SPEC.md` first** (same folder — full self-contained spec).
This file is the delta: what is DONE, what was LEARNED (overrides the spec where
noted), and what remains. Fresh session: skim "DONE", then start at "NEXT STEPS".

---

## Status: M0 ✅ M1 ✅ (gates passed). M2+M3 CODE WRITTEN (tsc clean) — **NOT YET RUN**.
NEXT: run Gate M2, then Gate M3(a)/(b), then M4 docs.

## DONE — do not redo

### M0 + M1 (verified in earlier sessions; details in git/spec)
- Node v22.1.0; python3 3.9.6 + PyMuPDF; fixture `/Users/ashoknaik/Downloads/UpdateTestDoc.pdf` (8 pages).
- Deps pinned (do NOT bump): deepagents@1.12.2, langchain@1.5.5, @langchain/core@1.2.5,
  @langchain/openai@1.5.6, @langchain/groq@1.3.1, zod@4.4.3, docx@9.7.1, mupdf@1.28.0.
  `scripts/fix-langgraph-sdk.mjs` postinstall must stay.
- Vendored `vendor/` with §7.1 patches (STATUS.NEEDS_LLM; pipeline.js helpers exported).
  NEVER import vendor/src/worker.js at runtime (shells out to `claude`).
- `models.json` is Groq-first text / free-OpenRouter vision — smoke-verified. Spec §3/§7.3
  IDs are STALE, do not revert. M4(a) pre-satisfied.
- M1 gate passed: `all <pdf> --no-agent` ⇒ 8/8 text pages, gate PASS 98% coverage, resume no-op.

### M2 + M3 code (written 2026-08-13, `npx tsc --noEmit` CLEAN, never executed)
All files below are COMPLETE — read them before changing anything, they interlock:

- `src/tools/validate.ts` — the fix-loop heart. Exports:
  - `validateLayoutForPage(workDir, cfg, page, layout, {model?, llmInput?})` → mirror of
    pipeline.js jsonOnce: ensureImages → renderLayoutJson (stamps injected) → verifySections
    → coverageOfText vs cfg.coverageMin (exact "low fidelity: only N% …" wording).
    PASS ⇒ writes pages/page_XXXX/layout.json + mode:"json" manifest (parent shape:
    layout/injected/fallbackText/assets/assetsDir), marks COMPLETED, tokens≈len/4, rm errors.json.
    FAIL ⇒ writes pages/page_XXXX/errors.json, returns {ok:false, coverage, errors}.
  - `validateModuleForPage(...)` — extractModule({injectDocxImport:true, requireExport:"buildPage"})
    → page.mjs → executeModule → verifySections (NO coverage — parent parity) ⇒ mode:"llm" manifest.
  - `readPageInput(workDir, page)` → {page, reason, textPrompt, injected, assetFiles, priorErrors?}.
  - `loadPageFiles()` throws a model-actionable error if page has no prewrites (never complex).
- `src/tools/vision.ts` — `visionReconstruct(workDir, cfg, page, log)`: code-driven loop.
  Re-renders page at dpi = clamp(72 … min(cfg.dpi, 1400*72/widthPts)) → base64 data-URL →
  walks ALL "vision" candidates; per candidate ≤ cfg.maxBuilderAttempts error-feedback rounds;
  lenient parse (`parseLayoutLenient` = worker.js parseLayoutJson + bracket-repair for truncated
  output) → validateLayoutForPage(..., {model:"provider:model", llmInput:"image"}).
- `src/tools/page-io.ts` — `makeTools(pdfPath, workDir, cfg, log)` closure factory (no
  model-supplied paths). Returns {orchestratorTools, pageBuilderTools, gateFixerTools, state}.
  - orchestrator: deterministic_pass, get_session, merge_document({outName?} sanitized,
    sets state.lastDocxPath), run_gate (uses state.lastDocxPath, errors if never merged),
    vision_reconstruct({page}), mark_page_image_fallback({page, reason}).
  - page_builder: read_page_input, validate_layout({page, layout: LOOSE z.looseObject}),
    validate_module({page, source}).
  - gate_fixer: get_session + read_page_input + both validators + mark_page_image_fallback.
  - ALL handlers catch and return JSON strings; failures = {ok:false, error} (no run-killing throws).
- `src/prompts/page-builder.ts` — exports `LAYOUT_SCHEMA_TEXT` (shared with vision) +
  `pageBuilderPrompt(cfg, skillDir)`: protocol read_page_input → build → validate_layout
  ≤ cfg.maxBuilderAttempts fixes → ONE validate_module → report. Coordinate rules from worker.js.
- `src/prompts/codegen.ts` — `CODEGEN_RULES(skillDir)`: Mode-B buildPage contract, string-escaping,
  units (half-points / eighths / twips), refs on disk.
- `src/prompts/orchestrator.ts` — `orchestratorPrompt(cfg, outName)`: spec §4 procedure,
  branches for cfg.llmMode==="none" (straight to image fallback) and cfg.forceVision
  (FIRST complex page through vision_reconstruct before builders); escalation ladder
  builder-retry("simpler structure") → vision_reconstruct → image fallback; ≤ maxGateFixLoops
  gate loops; structured final report.
- `src/prompts/gate-fixer.ts` — `gateFixerPrompt(cfg, skillDir)`: hint→fix mapping
  (missing fill → bar block; low coverage → transcribe more; image → img block; overflow →
  simplify table), then report; orchestrator re-merges.
- `src/agent.ts` — `runAgent(pdfPath, cfg, log)`: mkdir workDir, makeTools, two subagents
  (explicit tools + modelForRole), createDeepAgent({model, tools, systemPrompt, subagents,
  backend: new FilesystemBackend({rootDir: workDir})}), invoke({messages:[…"Convert <abs>
  to DOCX, output name <out>. Follow your procedure."]}, {recursionLimit:150}), prints final
  message. cli.ts already wires it (agent path = default when not --no-agent). NO cli changes needed.

### Fixture discovery (affects Gates M2/M3)
Under `--llm auto` all 8 fixture pages route deterministic (0 complex, NO fills/frames/images).
⇒ exercise builders with `--llm all`. ⇒ the fixture CANNOT fail the "injected fills" gate check
(no fills exist) — see Gate M3(b) recipe below for a corruption that actually fails the gate.

### Verified API facts (trust these)
- `tool(fn, {name, description, schema})` from "langchain"; `createDeepAgent`/`FilesystemBackend`
  from "deepagents"; SubAgent = plain {name, description, systemPrompt, tools, model}.
- Vendored JS imports from TS need `// @ts-ignore vendored JS` above each import.
- zod v4: z.looseObject used for the layout schema. IF runtime tool-call validation ever
  rejects model layouts, loosen further to z.record(z.any()) — failures must come from
  verifySections wording, not zod.

---

## NEXT STEPS (strict order)

### 1. GATE M2 (first-ever agent run — expect to debug)
```
npx tsx cli.ts all /Users/ashoknaik/Downloads/UpdateTestDoc.pdf --out m2test --llm all --fresh
```
Expect: orchestrator calls deterministic_pass (8 complex under --llm all) → dispatches
page_builder tasks (≤2 parallel) → pages complete mode:"json" (some may land mode:"llm" or
image — json is the target) → merge → gate PASS or explained. Then sanity:
`python3 vendor/extractandVerify.py outputs/m2test.docx` and check work/UpdateTestDoc/session.json.
On 429 stalls: `--model-index page_builder=1` (groq gpt-oss-120b) and/or `--concurrency 1`.
Likely first-run issues: orchestrator not following procedure (tighten prompts/orchestrator.ts),
tool-arg validation (loosen zod), Groq tool-call quirks (try --model-index orchestrator=1).

### 2. GATE M3(a) — vision rung
```
npx tsx cli.ts all /Users/ashoknaik/Downloads/UpdateTestDoc.pdf --out m3vision --llm all --force-vision --fresh
```
Expect ≥1 manifest with mode:"json" AND model:"openrouter:…vl…" (vision-produced), run completes.
Check: `grep -l '"llmInput": "image"' work/UpdateTestDoc/pages/*.json`.

### 3. GATE M3(b) — gate_fixer repair loop
After a successful M2/M3(a) run (do NOT --fresh afterwards):
1. Corrupt ≥5 of the 8 page manifests so word coverage drops below 60% (single-page corruption
   is NOT enough — threshold 0.6, one page ≈ 12%): for each chosen page_000X.json, replace
   `.layout.columns` with `[{"blocks":[{"t":"p","text":"lorem"}]}]` (keep fallbackText intact —
   expected words come from fallbackText; also overwrite pages/page_000X/layout.json to match).
   Session stays completed ⇒ deterministic_pass skips them.
2. Re-run WITHOUT --fresh: `npx tsx cli.ts all <fixture> --out m3fix --llm all`
   ⇒ orchestrator merges, gate fails word-coverage, perPageHints list the json pages,
   task(gate_fixer) rebuilds them via validate_layout, re-merge → gate PASS.

### 4. M4 — provider swap proof
(a) already satisfied (models.json Groq-first, smoke-verified).
(b) WRITE `README-deepagents.md`: exact models.json + env for
    `bedrock:anthropic.claude-opus-*` (`npm i @langchain/aws`, AWS_ACCESS_KEY_ID/
    AWS_SECRET_ACCESS_KEY/AWS_REGION) and `openai:gpt-*` (OPENAI_API_KEY) swaps + run commands.
(c) run `npx tsx src/smoke.ts --fake` (already implemented, just run it).

### 5. After each gate: `npx tsc --noEmit` must stay clean; check session.json + printed checks.

---

## Key file/shape reference (saves re-reading vendor)
- Manifest shapes consumed by vendor/src/docx.js sectionsForPage:
  mode:"json" needs .layout (+.injected, .fallbackText, .assetsDir); mode:"llm" needs .codeFile;
  .sections (text path) → renderJsonSections; else .blocks → blocksToChildren (image fallback).
- Layout-JSON: {title?, frame?, styles?{font,size(half-pts)}, columns:[{blocks:[Block]}]},
  Block t ∈ bar|h|p|list|table|img. table: {t,header?,cols?,rows:[[cell]]},
  cell = string|{text,bold?,italic?,align?}. img: {t:"img",file} (basename from assetFiles).
- Prewrites per complex page (deterministic-pass): pages/page_XXXX/{textprompt.txt,
  injected.json, assetFiles.json, origtext.txt, meta.json(.assets = rel paths for manifests)}
  + assets/page_XXXX/page.png. errors.json appears there after a failed validate.
- runGateOn perPageHints: missing fill hex → pages whose injected carry it; images-embedded →
  pages with assets; word-coverage → non-text-mode pages; table overflow → page 0 + detail.
- gate_fixer works only on pages with prewrites (needs_llm history) — text pages have none
  (readPageInput errors clearly).

## Environment / discipline
- `.env` (gitignored) has OPENROUTER_API_KEY + GROQ_API_KEY — working. Never print keys.
- Free tiers rate-limit: maxParallelPages 2; drop to --concurrency 1 / --model-index on 429s.
- outputs/ + work/ are runtime; `--fresh` rm-rf's work/<stem> (cli.ts does it, agent doesn't).

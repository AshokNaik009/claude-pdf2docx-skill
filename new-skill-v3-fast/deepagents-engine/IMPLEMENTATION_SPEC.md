# deepagents-engine — Implementation Spec

Port of the **detfast PDF → editable DOCX** engine from Claude-Code orchestration to
**LangChain Deep Agents (TypeScript)**, with a fully customizable model layer
(OpenRouter/Groq today → AWS Bedrock Opus / OpenAI later, by config only).

This document is **self-contained**: a fresh agent session must be able to implement the
whole system from this file plus the sibling parent folder (`../`). Read it fully before
writing code. Where this spec says "port X from `../src/foo.js`", the vendored copy in
this folder (created in Milestone 0) is the source of truth to read.

---

## 1. Background — what exists today (parent folder `../`)

The parent folder is a working engine driven by the **Claude Code CLI** (`claude -p`).
Its principle: *facts extractable deterministically (text, table grids, fills, borders,
the outer page frame) are INJECTED, never decided by a model; the model only does the
cheap judgment part (transcription + block assembly).*

Current pipeline (see `../README.md`, `../src/pipeline.js`):

```
PDF page
  → pdfextract.py (PyMuPDF)     text + coords + VECTOR DRAWINGS (fills/borders/frame)
  → route (quality.js):
       clean text/tables        → deterministic block builder     (ZERO LLM)
       fills / complex layout   → `claude -p` emits COMPACT LAYOUT JSON (not code)
  → deterministic renderer      → docx, STAMPING injected fills/borders/frame
  → fallback ladder             compact-JSON → docx.js code-gen (text→image) → page image
  → merge → .docx
  → extractandVerify.py gate    fills landed? tables fit? word coverage? images embedded?
```

Key existing modules (all vendored into this folder in M0):

| File | Role |
|------|------|
| `pdfextract.py` | PyMuPDF facts: structured text + compact `text_prompt` + drawings (fills/borders/frame). Invoked as a subprocess with `python3` (needs PyMuPDF; verified installed on this machine's system python3 3.9.6) |
| `extractandVerify.py` | pure-stdlib .docx inspector (the gate). `python3 extractandVerify.py out.docx --json report.json` |
| `src/pipeline.js` | per-page orchestration + routing + fallback ladder + session.json. **This is what the deep agent replaces** |
| `src/worker.js` | the `claude -p` prompts (`buildJsonPrompt`, `buildTextPrompt`, `buildPrompt`) — **port these prompts, drop the `claude` CLI calls** |
| `src/docx.js` | `renderLayoutJson(layout, ctx, injected)`, `renderJsonSections`, `executeModule`, `renderDocx`, `makeCtx` |
| `src/verify.js` | zero-LLM structural verifier `verifySections` + `coverage` fidelity check |
| `src/extract.js`, `tables.js`, `layout.js`, `geometry.js` | deterministic (zero-LLM) page reconstruction |
| `src/quality.js` | `classifyPage` — routes text vs llm |
| `src/pdf.js`, `png.js`, `session.js`, `merge.js`, `gate.js` | rendering, session state, ordering, gate wrapper |
| `config.js` | all defaults (coverageMin 0.6, dpi 200, minCharsForText 90, …) |
| `.claude/skills/docxjs/` | docx.js authoring skill: `SKILL.md`, `references/{codegen,tables,layout,images}.md`, `scripts/{extract-code,validate-page-module,image-fit}.mjs` |

The **compact LAYOUT JSON schema** the model emits (unchanged in this port — see
`../README.md` §"Layout-JSON schema" and `renderLayoutJson` in `src/docx.js`):

```jsonc
{
  "title": "…",                               // optional
  "styles": { "font": "…", "size": 22 },      // size = half-points
  "columns": [ { "blocks": [ /* Block */ ] } ]
}
// Block: {t:"bar",text} | {t:"h",text,level} | {t:"p",text} |
//        {t:"list",ordered,items[]} | {t:"table",header,cols[],rows[][]} | {t:"img",file}
```

---

## 2. Locked design decisions (do not re-litigate)

These were decided in a structured design interview. Implement exactly this.

| # | Decision | Choice |
|---|----------|--------|
| D1 | Agent role | **Hybrid agent-as-orchestrator.** Deterministic extract/render/verify stay as plain tools; the deep agent plans, delegates complex pages, and drives the verify→fix loop (the piece the old system lacked: its gate *reported* mismatches but never re-drove pages) |
| D2 | Runtime | **TypeScript `deepagents` (npm)**. (Python `deepagents` requires Python ≥3.11; the target constraint was 3.9.) The two Python scripts are invoked as subprocess tools with system `python3` |
| D3 | Granularity | **Coarse deterministic-pass tool**: one tool processes ALL pages deterministically and returns a worklist of complex pages. The agent never loops over clean pages |
| D4 | LLM decomposition | **3 specialized workers**: `page_builder` (subagent), `vision_page_builder` (code-driven tool — see §7.6 note), `gate_fixer` (subagent) |
| D5 | Providers | **OpenRouter `:free`-first, Groq fallback**, keys from env. Role→model mapping lives in one config file; Bedrock/OpenAI later = config edit only |
| D6 | Fix loop | **In-subagent `validate_layout` tool**: renders+verifies+coverage in one call, returns errors into the subagent's own context; ≤2 in-context retries, then orchestrator escalates (stronger model → vision → image fallback) |
| D7 | Files | **`FilesystemBackend` rooted at the real `work/` dir.** Subprocess tools and subagents share real files; binaries stay out of context |
| D8 | State | **session.json remains the source of truth** (owned by deterministic code). Agent checkpointer stays in-memory; resume = re-invoke agent, it reads session.json |
| D9 | Skill knowledge | **Distilled contracts embedded in subagent system prompts** (free models won't reliably read files); full `.claude/skills/docxjs` reference files kept on disk for the codegen-fallback rung |
| D10 | Concurrency | **Config `maxParallelPages`, default 2** (free-tier rate limits). Orchestrator may dispatch parallel `task()` calls up to N |
| D11 | Scaffolding | **Fully standalone vendored copy** — this folder must work with the parent folder deleted. New agent code in TypeScript; vendored `src/*.js` + `.py` files copied **unchanged** (except the small extensions listed in §7.1) |
| D12 | Test fixture | `/Users/ashoknaik/Downloads/UpdateTestDoc.pdf` (verified: 8 pages, healthy text layer, no raster images on p1). **The input PDF is a required CLI parameter** — never hard-code it |

---

## 3. Verified external facts (checked 2026-08-13 — trust these, don't re-research)

- **`deepagents` npm** (langchain-ai/deepagentsjs), v1.12.x. `createDeepAgent({ model, tools, systemPrompt, subagents, backend })`. Declares LangChain runtime packages as **peer dependencies** (`@langchain/core`, `@langchain/langgraph`, `@langchain/langgraph-checkpoint`, `@langchain/langgraph-sdk`, `langchain`, `langsmith`) — npm 7+ installs peers automatically. Node entrypoint exports `FilesystemBackend`.
- **SubAgent spec (JS)** — plain object: `{ name, description, systemPrompt, tools?, model?, middleware? }`. `model` accepts a **LangChain chat-model instance** (preferred here) or a `"provider:model"` string. Subagents inherit main-agent tools if `tools` omitted — **always set `tools` explicitly** in this project. Docs: https://docs.langchain.com/oss/javascript/deepagents/subagents
- **Deep agent main entry**: the orchestrator delegates via the built-in `task()` tool; a `general-purpose` subagent is auto-added unless disabled. Parallel `task()` calls in one turn are supported.
- **Groq via `@langchain/groq`** (`ChatGroq`, `GROQ_API_KEY`). Tool calling supported. **Vision on Groq is limited to Llama-4**: `meta-llama/llama-4-scout-17b-16e-instruct` and `meta-llama/llama-4-maverick-17b-128e-instruct`.
- **OpenRouter has no dedicated LangChain JS package** — use `ChatOpenAI` from `@langchain/openai` with `configuration: { baseURL: "https://openrouter.ai/api/v1" }` and `apiKey: process.env.OPENROUTER_API_KEY`.
- **Bedrock (future)** via `@langchain/aws` `ChatBedrockConverse` — a drop-in `BaseChatModel`, so the model factory covers it with zero changes elsewhere.
- **LangChain JS v1 / deepagents require Node ≥ 20.** Check `node --version` first.
- **This machine**: system `python3` = 3.9.6 **with PyMuPDF installed** (verified working against the test PDF). `docx@^9.7.1` + `mupdf@^1.28.0` are the only runtime deps of the vendored engine.

Model IDs supplied by the owner (use exactly these strings):
- OpenRouter free text: `openai/gpt-oss-120b:free`, `meta-llama/llama-3.3-70b-instruct:free`, `qwen/qwen3-next-80b-a3b-instruct:free`
- OpenRouter vision: `nvidia/nemotron-nano-12b-v2-vl:free`, `meta-llama/llama-4-scout-17b-16e-instruct`
- Groq fallbacks: `openai/gpt-oss-120b`, `llama-3.3-70b-versatile`, `meta-llama/llama-4-scout-17b-16e-instruct`

---

## 4. Target architecture

```
cli.ts  all <input.pdf> [--out NAME] [--llm auto|all|none] [--concurrency N] [--fresh] [--no-agent]
   │
   ▼
ORCHESTRATOR deep agent  (role: orchestrator; FilesystemBackend @ work/<stem>/)
   │  tools: deterministic_pass, get_session, merge_document, run_gate,
   │         vision_reconstruct, mark_page_image_fallback
   │  subagents (via task()): page_builder, gate_fixer
   │
   ├─ 1. deterministic_pass(pdf)        → clean pages DONE (zero LLM), complex pages flagged
   ├─ 2. for each complex page (≤N parallel): task(page_builder, page k)
   │        page_builder tools: read_page_input, validate_layout, validate_module
   │        internal loop: emit layout JSON → validate_layout → fix errors ≤2×
   │        escalation on failure: orchestrator retries page_builder with
   │        escalation model → vision_reconstruct(page) → mark_page_image_fallback(page)
   ├─ 3. merge_document(outName)        → outputs/<outName>.docx
   ├─ 4. run_gate(docx)                 → structured checks
   └─ 5. if gate fails: task(gate_fixer) → fixes pages via same validate tools
         → re-merge → re-gate (≤2 loops) → final report
```

Everything the model shouldn't decide stays where it is today: extraction, routing,
rendering, fill/frame injection, verification — deterministic TypeScript/JS/Python.
The LLM contributes exactly two things: (a) per-page transcription+assembly as compact
layout JSON (or docx.js code on the fallback rung), and (b) judgment in the fix loop.

---

## 5. Directory layout to create

```
deepagents-engine/
├── IMPLEMENTATION_SPEC.md        (this file)
├── .env                          (exists; gitignored — OPENROUTER_API_KEY, GROQ_API_KEY)
├── .env.example                  (exists)
├── .gitignore                    (exists)
├── package.json                  "type": "module"
├── tsconfig.json                 NodeNext, strict, allowJs (imports vendored .js)
├── models.json                   role → ordered candidate list  (§7.3)
├── cli.ts                        entry point (§7.8)
├── src/
│   ├── agent.ts                  createDeepAgent wiring: orchestrator + subagents
│   ├── model-factory.ts          provider adapters + role resolution (§7.4)
│   ├── config.ts                 engine defaults (port of ../config.js) + CLI overrides
│   ├── tools/
│   │   ├── deterministic-pass.ts (§7.5.1)
│   │   ├── page-io.ts            read_page_input, mark_page_image_fallback, get_session
│   │   ├── validate.ts           validate_layout, validate_module (§7.5.2)
│   │   ├── merge-gate.ts         merge_document, run_gate (§7.5.3)
│   │   └── vision.ts             vision_reconstruct (§7.6)
│   ├── prompts/
│   │   ├── orchestrator.ts       (§7.7)
│   │   ├── page-builder.ts       distilled from vendored worker.js buildJsonPrompt/buildTextPrompt
│   │   ├── codegen.ts            distilled from worker.js buildPrompt + skill codegen.md rules
│   │   └── gate-fixer.ts
│   └── smoke.ts                  M0 provider smoke test
├── vendor/                       ← vendored copies, UNCHANGED except §7.1 notes
│   ├── src/                      all 14 files from ../src/
│   ├── pdfextract.py
│   ├── extractandVerify.py
│   └── skills/docxjs/            full copy of ../.claude/skills/docxjs/
├── work/                         (runtime, gitignored)
└── outputs/                      (runtime, gitignored)
```

---

## 6. Milestones — implement strictly in this order

Each milestone has an acceptance gate. Do not start the next before the gate passes.

### M0 — Scaffold + vendor + provider smoke test
1. `node --version` must be ≥ 20; abort with a clear message otherwise.
2. `npm init -y`; set `"type": "module"`. Install (pinning majors):
   `deepagents`, `@langchain/openai`, `@langchain/groq`, `zod`, `docx@^9.7.1`,
   `mupdf@^1.28.0`, `dotenv`; dev: `typescript`, `tsx`, `@types/node`.
   (npm ≥7 auto-installs deepagents' peers. Do NOT install `@langchain/aws` yet — it's a
   documented future step, §9.)
3. Vendor: copy `../src/*.js` → `vendor/src/`, `../pdfextract.py`, `../extractandVerify.py`
   → `vendor/`, `../.claude/skills/docxjs/**` → `vendor/skills/docxjs/`. Copies are
   byte-identical except the extensions in §7.1.
4. `src/smoke.ts`: for each role in `models.json`, instantiate candidate #0 via the model
   factory and `invoke` a trivial prompt; print `role → provider:model → ok/fail`, then do
   the same for candidate #1 (fallback) of each role.

**Gate M0:** `npx tsx src/smoke.ts` shows at least one OpenRouter and one Groq model
answering. `npx tsc --noEmit` clean.

### M1 — Deterministic spine WITHOUT any LLM (`--no-agent`)
Implement `config.ts`, `tools/deterministic-pass.ts`, `tools/merge-gate.ts`, and a
`--no-agent` CLI path that calls them directly (no deepagents involved):
deterministic pass → complex pages immediately `mark_page_image_fallback` → merge → gate.
This is the port-equivalent of `node cli.js all input.pdf --llm none` and proves the
vendored engine works under tsx before any model is in the loop.

**Gate M1:** `npx tsx cli.ts all <pdf> --no-agent --out m1test` produces
`outputs/m1test.docx`; gate JSON printed; `work/<stem>/session.json` shows every page
`completed`; re-running without `--fresh` is a no-op (resume works).

### M2 — Orchestrator + page_builder (the core port)
Implement `model-factory.ts`, `agent.ts`, `prompts/*`, `tools/page-io.ts`,
`tools/validate.ts`. Full agent run with OpenRouter free models.

**Gate M2:** `npx tsx cli.ts all /Users/ashoknaik/Downloads/UpdateTestDoc.pdf --out m2test`
completes all 8 pages; complex pages carry `mode:"json"` manifests (not image fallbacks);
final gate `ok:true` OR every failing check is explained in the orchestrator's final
report; the .docx opens in Word/Pages.

### M3 — Vision rung + gate_fixer loop
Implement `tools/vision.ts` and the `gate_fixer` subagent; wire escalation order
(page_builder retry w/ escalation model → vision_reconstruct → image fallback).

**Gate M3:** (a) `--llm all --force-vision` run (flag forces one page through
`vision_reconstruct`) completes with a `mode:"json"` manifest produced by the vision
model; (b) deliberately corrupt one page's `layout.json`+manifest after M2, re-run merge
+ gate, and the gate_fixer loop repairs the page and re-merges to a passing gate.

### M4 — Provider-swap proof (the "customizable" requirement)
No code changes allowed in this milestone — only `models.json` edits.

**Gate M4:** (a) run the same doc with all roles pointed at Groq candidates first — works;
(b) add a `bedrock` provider entry to the factory behind a dynamic import of
`@langchain/aws` (this one code change is pre-authorized, do it in M2 if convenient),
document in README-deepagents.md the exact `models.json` + env for
`bedrock:anthropic.claude-opus-*` and `openai:gpt-*`; (c) a unit smoke that a fake
provider entry resolves through the factory.

---

## 7. Component specs

### 7.1 Vendored-file extensions (the only allowed edits in `vendor/`)
- `vendor/src/session.js`: add `NEEDS_LLM: "needs_llm"` to `STATUS`.
- `vendor/src/pipeline.js`: export the currently-private helpers `buildInjected`,
  `ensureImages`, `layoutText`, `coverageOfText`, `fallbackParagraphs`, `initSession`,
  `workDirFor` (add `export` keywords only; no logic changes). If splitting is cleaner,
  create `vendor/src/pipeline-helpers.js` with those functions copied verbatim.
- `vendor/src/worker.js`: keep for prompt reference; it is **never imported at runtime**
  (it would exec `claude`).
- Paths inside vendored files that reference `../.claude/skills/docxjs` must resolve to
  `vendor/skills/docxjs` — patch the constant, nothing else.

### 7.2 `config.ts`
Port `../config.js` values verbatim (dpi 200, minCharsForText 90, columnGapPts 40,
headingRatio 1.15, coverageMin 0.6, verifyRetries →`maxBuilderAttempts: 2`, gate on,
pyExtract on, `pyExtractBin: "python3"`). Add: `maxParallelPages: 2`,
`maxGateFixLoops: 2`, `workRoot`, `outputsDir` resolved **relative to deepagents-engine/**.
Every value overridable by CLI flag; `--llm auto|all|none` keeps its old meaning
(`none` ⇒ M1 behavior even with the agent enabled: agent skips complex pages to image
fallback).

### 7.3 `models.json` — role → ordered candidates
```jsonc
{
  "orchestrator":  [ {"provider":"openrouter","model":"openai/gpt-oss-120b:free"},
                     {"provider":"groq","model":"openai/gpt-oss-120b"} ],
  "page_builder":  [ {"provider":"openrouter","model":"meta-llama/llama-3.3-70b-instruct:free"},
                     {"provider":"openrouter","model":"qwen/qwen3-next-80b-a3b-instruct:free"},
                     {"provider":"groq","model":"llama-3.3-70b-versatile"} ],
  "escalation":    [ {"provider":"openrouter","model":"openai/gpt-oss-120b:free"},
                     {"provider":"groq","model":"openai/gpt-oss-120b"} ],
  "vision":        [ {"provider":"openrouter","model":"nvidia/nemotron-nano-12b-v2-vl:free"},
                     {"provider":"openrouter","model":"meta-llama/llama-4-scout-17b-16e-instruct"},
                     {"provider":"groq","model":"meta-llama/llama-4-scout-17b-16e-instruct"} ],
  "gate_fixer":    [ {"provider":"openrouter","model":"openai/gpt-oss-120b:free"},
                     {"provider":"groq","model":"openai/gpt-oss-120b"} ]
}
```
Entries may carry optional `params` (temperature, maxTokens). The Bedrock swap is:
`{"provider":"bedrock","model":"<bedrock-model-id>","params":{"region":"us-east-1"}}`.

### 7.4 `model-factory.ts`
```ts
type Provider = "openrouter" | "groq" | "openai" | "anthropic" | "bedrock";
modelForRole(role: string, candidateIndex = 0): BaseChatModel
```
- `openrouter` → `new ChatOpenAI({ model, apiKey: env.OPENROUTER_API_KEY,
  configuration: { baseURL: "https://openrouter.ai/api/v1" }, maxRetries: 4, ...params })`
- `groq` → `new ChatGroq({ model, maxRetries: 4, ...params })` (reads `GROQ_API_KEY`)
- `bedrock` → dynamic `import("@langchain/aws")`, `new ChatBedrockConverse(...)` —
  clear error message if the package isn't installed
- Transient 429/5xx: rely on the clients' built-in `maxRetries` exponential backoff.
- Role-level fallback: candidate order in `models.json` + CLI `--model-index role=N`;
  additionally, if instantiating/first-invoking candidate 0 throws (bad key, model
  retired), the factory logs and tries the next candidate at startup.
  (Do NOT build a custom runtime fallback wrapper around `bindTools` — `withFallbacks`
  interacts poorly with tool-binding agents; keep fallback at construction + rerun level.)

### 7.5 Tools (all defined with `tool()` from `langchain` + zod schemas; all return JSON-serializable objects; every tool takes `workDir` implicitly from a per-run closure, NOT from the model)

**7.5.1 `deterministic_pass({ pdfPath })`** — the coarse tool (D3).
Port of `runPipeline`/`processPage` from `vendor/src/pipeline.js` **minus** the
`mode === "llm"` branch:
- init/resume session (reuse `initSession`), per page: pyExtract structured text +
  drawings → `classifyPage` → image/figure assets → if routed `text`: deterministic
  build + `verifySections` + `coverage`; on success write manifest, mark `completed`.
- Any page routed to the model (or failing the deterministic gate) is marked
  `needs_llm`, and the tool **pre-writes its LLM inputs** under `pages/page_XXXX/`:
  `textprompt.txt` (the pyExtract `text_prompt`), `injected.json`
  (via `buildInjected(drawings)`), `assetFiles.json`, and `assets/page_XXXX/page.png`.
- Returns `{ workDir, totalPages, deterministicDone: number[], complexPages:
  [{ page, reason, hasFills, hasFrame, assetFiles }], failed: [...] }`.
- Honors `--fresh` (delete work dir first) and resume (skip `completed`).

**7.5.2 `validate.ts`** — the fix-loop heart (D6). Exposed to page_builder & gate_fixer:
- `read_page_input({ page })` → `{ textPrompt, injected, assetFiles, reason,
  priorErrors? }` read from the files above.
- `validate_layout({ page, layout })` — **the layout JSON arrives as the tool argument**
  (tool-calling gives us structured output for free; no fence parsing):
  `ensureImages(layout, assetFiles)` → `renderLayoutJson(layout, makeCtx(pageDir),
  injected)` → `verifySections` → `coverageOfText(origText, layoutText(layout))` with
  `coverageMin`. On pass: write `pages/…/layout.json` + manifest
  (`mode:"json"`, same shape as today incl. `injected`, `fallbackText`, `assets`) and
  mark the page `completed` in session.json. Returns
  `{ ok, coverage, errors: string[] }` — errors phrased exactly like today's verifier so
  the model can act on them.
- `validate_module({ page, source })` — codegen fallback rung: write `page.mjs`,
  `executeModule` → `verifySections`; on pass write `mode:"llm"` manifest, mark
  completed. Returns `{ ok, errors }`.
- Both validators are idempotent and safe to call repeatedly.

**7.5.3 `merge-gate.ts`:**
- `merge_document({ outName })` → `buildDocumentModel(workDir)` + `renderDocx` →
  `outputs/<outName>.docx`; returns `{ docxPath, mergedPages, missingPages }`.
- `run_gate({ docxPath })` → port of `runGate` (`vendor/src/gate.js`) with
  `gateScript: "vendor/extractandVerify.py"`; returns its `{ ok, checks[] }` plus a
  `perPageHints` map derived from failing checks (e.g. missing fill hex → the page(s)
  whose `injected.fills` contain that hex) so the fixer knows where to look.
- `get_session()` → compact per-page status summary (never the raw file — keep context
  small).
- `mark_page_image_fallback({ page, reason })` → manifest `{ mode:"image" }` referencing
  `page.png` + `fallbackText` (exists today as render path in `renderDocx`), mark
  completed. The ladder's floor — the output is never worse than the page picture.

### 7.6 `vision_reconstruct({ page })` — vision rung as a code-driven tool
Design refinement of D4 (flagged during the interview): tiny free VL models are
unreliable *tool-callers*, and `task()` delegation carries text, not images. So the
vision "worker" is a deterministic loop inside one tool, callable by the orchestrator:
1. Load `page.png`, downscale to ≤1400px wide (use `vendor/src/png.js` helpers or mupdf),
   base64 data-URL.
2. `modelForRole("vision").invoke([system: distilled layout-JSON contract,
   human: [image_url(dataUrl), text: transcribe+assemble instructions]])`.
3. Parse JSON leniently (strip fences → `JSON.parse`; one bracket-repair attempt), zod-
   validate the layout schema, then run the same validate-layout path in-process;
   feed errors back to the model ≤2 attempts.
4. Return `{ ok, coverage, errors }`; on success it has already written manifest+layout.

### 7.7 Agent wiring (`agent.ts`)
- `backend`: `new FilesystemBackend({ rootDir: workDir })` (confirm exact constructor
  signature from the installed package's types; it is exported from the `deepagents`
  Node entrypoint).
- Orchestrator: `createDeepAgent({ model: modelForRole("orchestrator"), tools: [
  deterministic_pass, get_session, merge_document, run_gate, vision_reconstruct,
  mark_page_image_fallback ], systemPrompt: ORCHESTRATOR_PROMPT, subagents: [
  pageBuilderSubagent, gateFixerSubagent ], backend })`.
- `page_builder` subagent: `{ name: "page_builder", description: "Reconstructs ONE
  complex PDF page as compact layout JSON; call with the page number.", systemPrompt:
  PAGE_BUILDER_PROMPT, tools: [read_page_input, validate_layout, validate_module],
  model: modelForRole("page_builder") }`.
- `gate_fixer` subagent: `{ name: "gate_fixer", …, tools: [get_session,
  read_page_input, validate_layout, validate_module, mark_page_image_fallback],
  model: modelForRole("gate_fixer") }` — it fixes pages *directly* with the same
  validators, then reports; the orchestrator re-merges and re-gates.
- ORCHESTRATOR_PROMPT must encode the exact procedure of §4 including: dispatch ≤
  `maxParallelPages` page_builder tasks per turn; escalation order per failed page
  (page_builder once more mentioning it should try harder & simpler structure →
  vision_reconstruct → mark_page_image_fallback); ≤ `maxGateFixLoops` gate loops; final
  message = structured run report (pages by mode, gate checks, files written).
- PAGE_BUILDER_PROMPT: distill from vendored `worker.js` `buildJsonPrompt`/`buildTextPrompt`
  — the coordinate-reading rules (x/y bands ⇒ columns, repeated x ⇒ table columns),
  transcribe-exactly rule, the layout-JSON schema with all six block types, fills/borders
  guidance ("host stamps colors — do NOT invent fills"), and the loop protocol:
  *call `read_page_input` first; then `validate_layout`; if `ok:false`, fix the listed
  errors and call again (max 2 fixes); if still failing and the errors are structural,
  try ONE `validate_module` with docx.js code following the codegen rules; then stop and
  report failure.* Keep it under ~2.5k tokens; the full skill references stay on disk at
  `vendor/skills/docxjs/` (D9) and the prompt lists those paths for the codegen rung.

### 7.8 `cli.ts`
```
npx tsx cli.ts all <input.pdf> [--out NAME] [--llm auto|all|none] [--concurrency N]
                   [--fresh] [--no-agent] [--force-vision] [--model-index role=N]
npx tsx cli.ts merge <workDir> --out NAME        # re-merge only
npx tsx cli.ts gate  <docxPath>                  # re-gate only
```
- `<input.pdf>` REQUIRED positional (D12). Loads `.env` via `dotenv/config`.
- Agent invocation: single `agent.invoke({ messages: [{ role: "user", content:
  "Convert <abs pdf path> to DOCX, output name <out>. Follow your procedure." }] })`,
  with `recursionLimit` ≥ 150; stream events to stdout (page-level progress lines like
  today's `✓ page 3 [json: …] 8123ms`).

---

## 8. Testing & acceptance

- Primary fixture: `/Users/ashoknaik/Downloads/UpdateTestDoc.pdf` (8 pages). Always pass
  it as the CLI argument; nothing about it may be hard-coded.
- After every milestone: run its gate command; inspect `work/<stem>/session.json`,
  the printed gate checks, and open the .docx.
- Structural check without Word: `python3 vendor/extractandVerify.py outputs/<name>.docx`.
- Budget discipline: free tiers rate-limit hard. If a run stalls on 429s, drop
  `--concurrency 1` and/or `--model-index page_builder=2` (Groq).

## 9. Pitfalls (learned from the existing system + verified constraints)

1. **docx.js fails silently** — wrong enums/units pack fine and render broken. Never trust
   "it packed"; only `verifySections` + gate decide success. (See skill `SKILL.md`
   "Common pitfalls".)
2. **Units**: font `size` = half-points; borders `size` = eighths of a point (1pt ⇒ 8);
   twips everywhere else. The prompts must carry these — they're the #1 model bug.
3. **Free-model JSON**: never ask for raw JSON in prose; the layout arrives as
   **tool-call arguments** (7.5.2), which the API enforces as JSON.
4. **Rate limits**: `:free` OpenRouter models have daily caps; Groq free tier RPM caps.
   maxRetries backoff + concurrency 2 + role fallback is the mitigation (D5/D10).
5. **Vision payloads**: base64 data-URLs must stay small (Groq limit ~4MB) — downscale
   first (7.6).
6. **Peer deps**: install with npm ≥7 so deepagents' peers resolve to one copy; do not
   mix pnpm/yarn here.
7. **Windows paths / spaces**: always quote; workDir derives from the PDF stem like today
   (`workDirFor`).
8. **Resume semantics**: `deterministic_pass` must remain idempotent — the agent may call
   it again on a resumed run; it must skip `completed` pages (it does today).
9. **Never import `vendor/src/worker.js` at runtime** — it shells out to `claude`.
10. **Secrets**: keys live in `.env` (gitignored) only. Never print them, never commit
    them, never bake them into config files. `.env.example` documents the names.

## 10. Future: enterprise swap (Bedrock Opus / OpenAI)

1. `npm i @langchain/aws` (or `@langchain/openai` is already present for OpenAI-proper).
2. Set AWS creds env (`AWS_ACCESS_KEY_ID`/`AWS_SECRET_ACCESS_KEY`/`AWS_REGION`) — the
   "client id + secret" the owner mentioned.
3. Edit `models.json` roles to `{"provider":"bedrock","model":"<opus-model-id>"}` —
   check the exact model ID available in the target account/region at swap time.
4. No other change: the factory returns a `BaseChatModel`; deepagents accepts instances
   everywhere (main agent + per-subagent + vision tool).

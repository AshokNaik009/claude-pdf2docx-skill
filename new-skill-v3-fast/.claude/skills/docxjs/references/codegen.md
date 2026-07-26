# Generating docx.js code programmatically (headless / `claude -p` pipelines)

Read this whenever the output is **JavaScript source that another process will execute**, rather
than a document you build yourself — e.g. a page-reconstruction worker that calls `claude -p`, asks
for a module, then runs it and merges the result.

The failure mode in these pipelines is almost never the docx API (the library is extremely
permissive — wrong enums, ragged rows, and missing `type` on images all pack without throwing).
Failures are **JavaScript syntax errors created while transcribing page text**, plus **wrapper
residue** around the code. Both are fully preventable with the rules below.

## Contents
[Return contract](#contract) · [Syntax safety when transcribing](#syntax) · [Output framing](#framing) ·
[API rules that matter](#api) · [Self-check before emitting](#selfcheck) · [Host-side hardening](#host)

<a id="contract"></a>
## 1. Follow the host's return contract exactly

There are two distinct modes. Know which one you are in **before writing a line**:

**Mode A — standalone document (you own the whole file).** Build and return a `docx.Document`, and
pack it yourself with `Packer.toBlob` / `toBuffer`. This is the default in the main SKILL.md.

**Mode B — page/fragment module (a host executes your code).** The host supplies `docx` and merges
your output. Typical contract:

```js
export function buildPage(docx, ctx) {
  // ...build paragraphs/tables...
  return {
    sections,          // ARRAY of { properties, children }
    styles,            // optional: { default: { document: { run: { font: "...", size: 22 } } } }
  };
}
```

In Mode B, these are hard errors:

- **Do not** `import` or `require` anything — not even `docx`. Use the `docx` parameter.
- **Do not** construct `new docx.Document(...)`, and never call `docx.Packer`. The host does that.
- **Do not** read files, fetch, or touch `fs`. Images come only from the host, e.g. `ctx.image("fig-1.png")`.
- **Do not** write top-level side effects (no `console.log`, no code that runs on import).
- Export **exactly one** function with the exact name the host asked for, and return the exact
  shape it asked for. If the host says "array of sections", returning a bare `Document` breaks it.

Helper functions are fine — define them at module top level or inside `buildPage`; just make sure
any `docx` they use is in scope (define them **inside** `buildPage`, or pass `docx` in, since in
Mode B there is no module-level `docx` import).

```js
export function buildPage(docx, ctx) {
  // Helpers live INSIDE so `docx` is always in scope — avoids "docx is not defined".
  const bodyPara = (t) => new docx.Paragraph({
    alignment: docx.AlignmentType.JUSTIFIED, spacing: { after: 180, line: 288 },
    children: [new docx.TextRun(t)],
  });
  const sections = [
    { properties: { page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } },
                    column: { count: 2, space: 560, separate: false } },
      children: [bodyPara("...")] },
  ];
  return { sections, styles: { default: { document: { run: { font: "Palatino Linotype", size: 22 } } } } };
}
```

<a id="syntax"></a>
## 2. Syntax safety when transcribing page text (the #1 cause of failures)

Transcribed prose is full of characters that terminate JS strings. These are **real
`SyntaxError`s** — the module never even parses, so the host sees "syntax error" regardless of how
capable the model is.

Verified failures, all from ordinary page text:

| Text being transcribed | Naive output | Result |
|---|---|---|
| `the product's surface` | `'the product's surface'` | SyntaxError: Unexpected identifier `s` |
| `a "text and" zone forms` | `"a "text and" zone forms"` | SyntaxError: Unexpected identifier `text` |
| a paragraph with a real line break | `"line one⏎line two"` | SyntaxError: Invalid or unexpected token |
| `C:\path\` | `"C:\path\"` | SyntaxError: Invalid or unexpected token |

Note the first row: an apostrophe is only a problem because the string was *single*-quoted. In a
double-quoted string it is harmless.

### The rule that eliminates all of them — without changing the text

**Use double-quoted strings, and escape exactly two characters: `"` → `\"` and `\` → `\\`.**

Nothing else needs escaping. Verified inside a double-quoted string, all of these are safe *as-is*:
apostrophes (`it's`), parentheses (`(Hello)`), brackets, braces, `$`, `${...}`, backticks, `&`,
`<`, `>`, `%`, em dashes, `°`, `µ`, `²`, accented letters, and already-typographic quotes (`“ ” ’`).

That matters because **the apostrophe problem disappears on its own** once the string is
double-quoted — apostrophes only break *single*-quoted strings. So the entire rule is: double
quotes outside, escape `"` and `\` inside.

```js
// Page reads:  the "text and" surface, per Wagner's model (Hello)
new docx.TextRun("the \"text and\" surface, per Wagner's model (Hello)")
```

This is **lossless** — the document receives exactly the characters the page had. Measured over a
sample of real page strings (straight quotes, apostrophes, nested quotes, parentheses, backslash
paths, tabs, `${`): double-quote-plus-escape parsed 7/7 and round-tripped 7/7 identical, while
single-quoting and unescaped double-quoting each parsed only 4/7.

Mechanically, this is exactly what `JSON.stringify(text)` produces — if you are ever unsure how to
escape a nasty string, emitting `JSON.stringify`'s output is always correct.

**Do not "fix" the page's punctuation.** Converting straight quotes to curly (`"…"` → `“…”`) or
apostrophes to `’` also avoids the syntax error, but it silently alters the document's content —
wrong when the source genuinely uses straight quotes, and unacceptable for verbatim
transcription, quoted terms, code samples, measurements like `6" pipe`, or legal text. Reproduce
the characters that are on the page; only convert if the page itself shows typographic quotes.

Additional rules:

- **Never put a literal line break inside a string.** A paragraph is one string on one line, however
  long. Do not wrap transcribed prose across source lines. If a page has a hard line break inside a
  paragraph, that is still one paragraph — join it with a space.
- **Never end a string with a lone `\`.** Write `\\` for a literal backslash.
- **Avoid template literals** for transcribed text; a stray `${` in the page becomes an interpolation.
  Use double-quoted strings.
- **Tabs are fine as `\t`** — good for `(i)\tDirect contact...` list cells.
- **Em dashes, degree signs, µ, ², accented characters are all safe** — emit them literally.
- Keep each string on one line and **balance every bracket**; long transcriptions are where
  unbalanced `)` / `]` / `}` creep in. If a paragraph is very long, that is fine — one long line is
  safer than a wrapped one.

<a id="framing"></a>
## 3. Output framing — emit code and nothing else

Hosts extract code with a regex. Anything around it can survive extraction and become a syntax
error. Verified: a module followed by one sentence of prose is a `SyntaxError`, and typical fence
handling only survives the case where the entire reply is a single clean fence.

When the host says "output only the module":

- No markdown fences.
- No sentence before ("Here is the module:") or after ("This reconstructs the page.").
- No explanatory comment block outside the code — comments *inside* the code are fine.
- Emit exactly one module. Not two alternatives, not a "simplified version" appended.

If a fence is unavoidable, emit **one** fence containing the whole module and nothing outside it.

<a id="api"></a>
## 4. API rules that actually matter

Because docx rarely throws, these mistakes produce a *silently wrong document* rather than an error
— which is worse in a batch pipeline that no one reviews page by page.

- `italics`, **not** `italic`.
- `docx.AlignmentType.JUSTIFIED` — `JUSTIFY` is `undefined` and silently drops the alignment.
- A `TableCell`'s `children` must be **`Paragraph`s** (or nested `Table`s), never bare `TextRun`s —
  a bare run packs fine and renders as an empty cell.
- Every `TableCell` needs at least one paragraph; `new docx.TableCell({})` throws
  (`options.children is not iterable`). Use `children: [new docx.Paragraph({})]` for blank cells.
- A `Table` with `rows: []` throws (`Invalid array length`). Omit the table instead.
- **Every row must have the same cell count** once spans are counted. After a `rowSpan: N`, the
  covered rows must have one **fewer** cell. Ragged rows pack without error and render as a mangled
  grid.
- `ImageRun` needs `type` (`"png" | "jpg" | "gif" | "bmp"`) plus `transformation: { width, height }`.
- **Images: never write dimensions.** If the host provides `ctx.image(file)`, call it with no size —
  it reads the asset's real aspect ratio. Hand-written `transformation: { width, height }` values
  are guesses that pack silently and render distorted. See `images.md` for placement rules; the
  short version is that a full-width figure gets its own `CONTINUOUS` single-column section, and a
  missing asset gets a centered placeholder paragraph, never a fabricated table of its labels.
- Never leave `undefined` in a `children` array and never nest an array inside `children` — both
  pack "successfully" and silently drop content. Build arrays with `.map(...)` and spread them:
  `children: [heading, ...rows]`.
- Use enums, not strings: `alignment: "justified"` is silently ignored.

For layout (multi-column bands, continuous section breaks) see `layout.md`; for table structure see
`tables.md`.

<a id="selfcheck"></a>
## 5. Self-check before emitting

Run this mentally over the finished module — it catches essentially every pipeline failure:

1. Does the output start with the first character of code and end with the last? (No fences, no prose.)
2. Exactly one exported function, with the required name and return shape?
3. Zero `import` / `require` / `new docx.Document` / `Packer` (Mode B)?
4. Is every transcribed string double-quoted, with `"` escaped as `\"` and `\` as `\\` — and the
   page's punctuation otherwise reproduced exactly (no quote "tidying")?
5. Does any string contain a real line break or end in a lone `\`?
6. Are all brackets balanced, and is `docx` in scope for every helper?
7. Every `TableCell` has ≥1 `Paragraph`; every row has matching cell counts; no `undefined` in arrays?
8. Full-width tables/figures are in their own `CONTINUOUS` section, never inside a two-column one?

<a id="host"></a>
## 6. Host-side hardening (for whoever builds the pipeline)

Prompt rules reduce failures; they do not eliminate them. Two cheap host-side measures make the
pipeline reliable:

- **Extract code robustly.** Prefer the *last* fenced block if any fence exists; strip stray prose
  lines before the first `import`/`export`/`function` and after the final `}`. See
  `scripts/extract-code.mjs`.
- **Validate before merging, and retry on failure.** Parse the module, execute `buildPage` against
  the real `docx`, and pack a throwaway document. Feed the error text back into a retry prompt —
  the model fixes its own syntax error nearly every time when it sees the message. See
  `scripts/validate-page-module.mjs`, which reports structural warnings (empty cells, ragged rows,
  full-width tables inside multi-column sections) as well as hard errors.

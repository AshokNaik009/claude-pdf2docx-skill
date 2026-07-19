/**
 * Robust extraction of a JS module from an LLM reply.
 *
 * Replaces the common `^```lang\n(...)\n```$` regex, which only survives the case where the
 * ENTIRE reply is one clean fence. Verified failures of that approach: prose before the fence,
 * prose after the fence, two fenced blocks, and a fence whose closing ``` has no preceding
 * newline — all leave backticks or prose in the source and produce a SyntaxError.
 *
 * Usage:
 *   import { extractModule } from "./extract-code.mjs";
 *   const code = extractModule(stdout, { injectDocxImport: false });
 */

/**
 * @param {string} raw            the model's stdout
 * @param {object} [opts]
 * @param {boolean} [opts.injectDocxImport=false]  prepend `import * as docx from "docx"`.
 *        Leave FALSE for buildPage(docx, ctx) contracts — docx arrives as a parameter.
 * @param {string} [opts.requireExport]  e.g. "buildPage"; throws if not present.
 * @returns {string} module source
 */
export function extractModule(raw, opts = {}) {
  const { injectDocxImport = false, requireExport } = opts;
  let c = String(raw || "").trim();

  // 1) Prefer fenced blocks. Take the LAST/longest one — models often show a snippet then the
  //    real module, and a trailing "here's what it does" fence is never the answer.
  const fences = [...c.matchAll(/```[a-zA-Z]*\s*\n([\s\S]*?)```/g)].map((m) => m[1]);
  if (fences.length) {
    c = fences.sort((a, b) => b.length - a.length)[0].trim();
  } else if (c.includes("```")) {
    // Unterminated fence: drop the opener and anything after a later stray fence.
    c = c.replace(/^```[a-zA-Z]*\s*\n?/, "").split("```")[0].trim();
  }

  // 2) Trim prose before the first real code line and after the last closing brace.
  const lines = c.split("\n");
  const firstCode = lines.findIndex((l) =>
    /^\s*(import\s|export\s|function\s|const\s|let\s|var\s|class\s|\/\/|\/\*|async\s)/.test(l));
  if (firstCode > 0) lines.splice(0, firstCode);
  c = lines.join("\n").trim();

  const lastBrace = c.lastIndexOf("}");
  if (lastBrace !== -1) {
    const tail = c.slice(lastBrace + 1);
    // Keep only trailing whitespace/semicolons/comments; drop prose sentences.
    if (/[A-Za-z]/.test(tail.replace(/\/\/.*$|\/\*[\s\S]*?\*\//g, ""))) c = c.slice(0, lastBrace + 1);
  }

  // 3) Optionally strip docx imports and re-inject a canonical one (Mode A hosts only).
  if (injectDocxImport) {
    c = c
      .split("\n")
      .filter((ln) => !/^\s*(import\s.*from\s*['"]docx['"]|(const|let|var)\s.*=\s*require\(['"]docx['"]\))/.test(ln))
      .join("\n")
      .trim();
    c = `import * as docx from "docx";\n\n${c}`;
  }

  if (requireExport && !new RegExp(`export\\s+(async\\s+)?function\\s+${requireExport}\\b`).test(c)) {
    throw new Error(`generated code missing \`export function ${requireExport}\``);
  }
  if (c.includes("```")) throw new Error("code still contains markdown fences after extraction");
  return c;
}

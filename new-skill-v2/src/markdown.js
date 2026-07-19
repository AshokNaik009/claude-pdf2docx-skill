// Parse GitHub-Flavored Markdown (as returned by headless Claude) into the
// pipeline's block model, and serialize the block model back to Markdown.
//
// Block shapes:
//   { type: "heading", level, text }
//   { type: "paragraph", text }
//   { type: "list", ordered, items: [text, ...] }
//   { type: "table", rows: [[cell, ...], ...] }
//   { type: "image", path, width, height, alt }

const TABLE_SEP = /^\s*\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?\s*$/;

function splitRow(line) {
  let s = line.trim();
  if (s.startsWith("|")) s = s.slice(1);
  if (s.endsWith("|")) s = s.slice(0, -1);
  return s.split("|").map((c) => c.trim());
}

/** Strip a whole-output ```fence``` if the model wrapped everything in one. */
function stripOuterFence(md) {
  const t = md.trim();
  const m = t.match(/^```[a-zA-Z]*\n([\s\S]*?)\n```$/);
  return m ? m[1] : t;
}

export function markdownToBlocks(md) {
  const lines = stripOuterFence(md).split(/\r?\n/);
  const blocks = [];
  let para = [];
  const flushPara = () => {
    if (para.length) {
      blocks.push({ type: "paragraph", text: para.join(" ").replace(/\s+/g, " ").trim() });
      para = [];
    }
  };

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const t = line.trim();

    if (!t) { flushPara(); continue; }

    // Heading
    const h = t.match(/^(#{1,6})\s+(.*)$/);
    if (h) { flushPara(); blocks.push({ type: "heading", level: h[1].length, text: h[2].trim() }); continue; }

    // Table: a row that is followed by a separator row
    if (t.includes("|") && i + 1 < lines.length && TABLE_SEP.test(lines[i + 1])) {
      flushPara();
      const rows = [splitRow(t)];
      i += 2; // skip header + separator
      while (i < lines.length && lines[i].includes("|") && lines[i].trim()) {
        rows.push(splitRow(lines[i]));
        i++;
      }
      i--;
      blocks.push({ type: "table", rows });
      continue;
    }

    // List (consecutive bullet/number lines)
    const li = t.match(/^([-*+]|\d+[.)])\s+(.*)$/);
    if (li) {
      flushPara();
      const ordered = /\d/.test(li[1]);
      const items = [li[2].trim()];
      while (i + 1 < lines.length) {
        const m = lines[i + 1].trim().match(/^([-*+]|\d+[.)])\s+(.*)$/);
        if (!m) break;
        items.push(m[2].trim());
        i++;
      }
      blocks.push({ type: "list", ordered, items });
      continue;
    }

    // Standalone image
    const img = t.match(/^!\[([^\]]*)\]\(([^)]+)\)$/);
    if (img) { flushPara(); blocks.push({ type: "image", alt: img[1], path: img[2] }); continue; }

    para.push(t);
  }
  flushPara();
  return blocks;
}

/** Serialize the block model back to Markdown (the intermediate .md output). */
export function blocksToMarkdown(blocks) {
  const out = [];
  for (const b of blocks) {
    switch (b.type) {
      case "heading":
        out.push("#".repeat(Math.min(6, b.level || 1)) + " " + b.text, "");
        break;
      case "paragraph":
        out.push(b.text, "");
        break;
      case "list":
        b.items.forEach((it, idx) => out.push((b.ordered ? `${idx + 1}. ` : "- ") + it));
        out.push("");
        break;
      case "table":
        if (b.rows.length) {
          out.push("| " + b.rows[0].join(" | ") + " |");
          out.push("| " + b.rows[0].map(() => "---").join(" | ") + " |");
          for (let r = 1; r < b.rows.length; r++) out.push("| " + b.rows[r].join(" | ") + " |");
          out.push("");
        }
        break;
      case "image":
        out.push(`![${b.alt || ""}](${b.path})`, "");
        break;
    }
  }
  return out.join("\n");
}

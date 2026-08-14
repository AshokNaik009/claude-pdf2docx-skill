// MERGE & ORDER: order pages by page number from session.json and load each
// page manifest into the ordered document model consumed by docx.js.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { loadSession, STATUS } from "./session.js";

export function buildDocumentModel(workDir) {
  const session = loadSession(workDir);
  const ordered = [...session.pages].sort((a, b) => a.page - b.page);
  const missing = ordered.filter((p) => p.status !== STATUS.COMPLETED);

  const pages = [];
  for (const p of ordered) {
    if (p.status !== STATUS.COMPLETED) continue;
    const m = JSON.parse(readFileSync(join(workDir, p.manifest), "utf8"));
    pages.push(m); // { page, mode, sections?|layout?|codeFile?, fallbackText?, assets, assetsDir? }
  }
  return { session, pages, missing };
}

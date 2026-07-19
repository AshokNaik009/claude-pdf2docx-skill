import { writeFileSync, mkdirSync } from "node:fs";
import { join, resolve } from "node:path";
import { buildHtmlDocument, renderPandocDocx } from "./src/htmldocx.js";

const workDir = resolve("work/newtest-document");
mkdirSync(join(workDir, "pages/page_9001"), { recursive: true });
writeFileSync(join(workDir, "pages/page_9001/page.html"), `
<table>
<tr>
<td>
<img src="page.png" alt="photo">
<h2>Ashok Naik</h2>
<p><strong>Skills</strong></p>
<ul><li>Langchain</li><li>Docker, AWS</li></ul>
</td>
<td>
<h3>Experience</h3>
<p><strong>HCL Tech</strong> — Technical Lead</p>
<ul><li><strong>Angular to React Migration:</strong> Led the migration.</li></ul>
</td>
</tr>
</table>
`);
const model = { pages: [
  { page: 1, mode: "text", assetsDir: "assets/page_0001",
    sections: [{ blocks: [
      { type: "heading", level: 1, text: "Conditions of Offer" },
      { type: "paragraph", text: "1.1. This offer is subject to a work permit." },
      { type: "table", rows: [["Component","Amount"],["Basic","AED 2000"]] },
    ]}] },
  { page: 2, mode: "html", htmlFile: "pages/page_9001/page.html", assetsDir: "assets/page_0001" },
]};
const html = buildHtmlDocument(workDir, model);
const buf = await renderPandocDocx(html);
writeFileSync("/tmp/htmltest.docx", buf);
console.log("wrote", buf.length, "bytes");

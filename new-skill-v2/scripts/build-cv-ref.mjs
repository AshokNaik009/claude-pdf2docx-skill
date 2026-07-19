// Build a pandoc reference.docx tuned to the CV's teal palette (#218FB4).
// Starts from pandoc's default reference doc, retints heading/link colors to
// teal, and adds shaded paragraph styles (Sidebar, SidebarDark, IntroBox) that
// the Lua filter routes HTML classes onto. Usage: node scripts/build-cv-ref.mjs [out.docx]
import { execFileSync } from "node:child_process";
import { mkdtempSync, readFileSync, writeFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";

const TEAL = "218FB4";
const LIGHT = "D8EAF0";
const outPath = resolve(process.argv[2] || "assets/cv-ref.docx");

const dir = mkdtempSync(join(tmpdir(), "cvref-"));
const defBuf = execFileSync("pandoc", ["--print-default-data-file", "reference.docx"], { maxBuffer: 32 << 20 });
writeFileSync(join(dir, "ref.docx"), defBuf);
execFileSync("unzip", ["-o", "-q", join(dir, "ref.docx"), "-d", join(dir, "x")]);

const stylesPath = join(dir, "x", "word", "styles.xml");
let s = readFileSync(stylesPath, "utf8");

// 1) Retint every heading style's run color to teal (drop themeColor so it wins).
s = s.replace(/<w:style w:type="paragraph" w:styleId="Heading\d"[\s\S]*?<\/w:style>/g, (blk) =>
  blk.replace(/<w:color\b[^>]*\/>/, `<w:color w:val="${TEAL}"/>`)
);
// 2) Hyperlinks teal too.
s = s.replace(/(<w:style w:type="character" w:styleId="Hyperlink"[\s\S]*?)<w:color\b[^>]*\/>/,
  `$1<w:color w:val="${TEAL}"/>`);

// 3) Shaded paragraph styles for the sidebar + intro box.
const shaded = (id, name, fill, textColor) => `
  <w:style w:type="paragraph" w:customStyle="1" w:styleId="${id}">
    <w:name w:val="${name}"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>
      <w:spacing w:before="20" w:after="20"/>
      <w:ind w:left="120" w:right="120"/>
    </w:pPr>
    <w:rPr>${textColor ? `<w:color w:val="${textColor}"/>` : ""}</w:rPr>
  </w:style>`;
s = s.replace("</w:styles>",
  shaded("Sidebar", "Sidebar", LIGHT, null) +
  shaded("SidebarDark", "Sidebar Dark", TEAL, "FFFFFF") +
  shaded("IntroBox", "Intro Box", LIGHT, null) +
  "\n</w:styles>");

writeFileSync(stylesPath, s);

// Rezip the modified tree back into the reference docx.
rmSync(outPath, { force: true });
execFileSync("bash", ["-c", `cd "${join(dir, "x")}" && zip -q -r -X "${outPath}" .`]);
rmSync(dir, { recursive: true, force: true });
console.log("wrote", outPath);

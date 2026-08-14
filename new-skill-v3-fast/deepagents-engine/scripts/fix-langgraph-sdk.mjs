// Postinstall shim: @langchain/langgraph-sdk ships bundled ESM deps under
// dist/node_modules/.pnpm/** WITHOUT package.json files. Node's module-format
// lookup stops at a node_modules boundary, so on Node < 22.7 (no default
// syntax detection) those .js files are compiled as CJS and crash on `import`.
// Writing a {"type":"module"} package.json next to each bundled package fixes
// loading on every Node version. Idempotent; a no-op if the layout changes.
import { readdirSync, existsSync, writeFileSync, statSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const pnpmDir = join(root, "node_modules", "@langchain", "langgraph-sdk", "dist", "node_modules", ".pnpm");

if (existsSync(pnpmDir)) {
  let fixed = 0;
  for (const entry of readdirSync(pnpmDir)) {
    const nm = join(pnpmDir, entry, "node_modules");
    if (!existsSync(nm) || !statSync(nm).isDirectory()) continue;
    for (const pkg of readdirSync(nm)) {
      const dir = join(nm, pkg);
      if (!statSync(dir).isDirectory()) continue;
      const pj = join(dir, "package.json");
      if (!existsSync(pj)) {
        writeFileSync(pj, JSON.stringify({ type: "module" }) + "\n");
        fixed++;
      }
    }
  }
  if (fixed) console.log(`fix-langgraph-sdk: marked ${fixed} bundled package(s) as ESM`);
}

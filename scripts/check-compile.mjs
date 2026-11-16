// Compila todas las páginas .asp y reporta las que fallan.
import fs from "node:fs";
import path from "node:path";
import { compileAsp } from "../server/asp/compiler.js";

const root = path.resolve(path.dirname(new URL(import.meta.url).pathname), "..");
// legacy/: código original que ya no se usa; server/ports sí se compila (reemplaza a las páginas VBScript)
const SKIP = /^(legacy|node_modules|\.git|data|Courses|demo)(\/|$)|^server\/(asp|compat)(\/|$)/;
const files = [];
(function walk(d) {
  for (const e of fs.readdirSync(d, { withFileTypes: true })) {
    const rel = path.relative(root, path.join(d, e.name));
    if (SKIP.test(rel)) continue;
    if (e.isDirectory()) walk(path.join(d, e.name));
    else if (/\.asp$/i.test(e.name)) files.push(path.join(d, e.name));
  }
})(root);
let ok = 0;
const vbscript = [];
const fails = [];
for (const f of files) {
  try {
    const r = compileAsp(f, root);
    if (r.vbscript) vbscript.push(path.relative(root, f));
    else ok++;
  } catch (e) {
    const loc = e.aspLocation ? `${path.relative(root, e.aspLocation.file)}:${e.aspLocation.line}` : "?";
    fails.push(`${path.relative(root, f)}  ->  ${loc}  ${e.message.split("\n")[0]}`);
  }
}
console.log(`${files.length} páginas (incluye server/ports): ${ok} OK, ${vbscript.length} en VBScript sin portar, ${fails.length} con errores`);
for (const l of vbscript) console.log("  VBScript sin portar: " + l);
for (const l of fails) console.log("  " + l);

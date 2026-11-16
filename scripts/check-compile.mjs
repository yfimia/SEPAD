// Compila todas las páginas .asp y reporta las que fallan.
import fs from "node:fs";
import path from "node:path";
import { compileAsp } from "../server/asp/compiler.js";

const root = path.resolve(path.dirname(new URL(import.meta.url).pathname), "..");
const SKIP = /^(Beta|SEPAD_06_01_09_Agenda|node_modules|\.git|data)(\/|$)|pages_esp\/Administracion\/pages_esp/;
const files = [];
(function walk(d) {
  for (const e of fs.readdirSync(d, { withFileTypes: true })) {
    const rel = path.relative(root, path.join(d, e.name));
    if (SKIP.test(rel)) continue;
    if (e.isDirectory()) walk(path.join(d, e.name));
    else if (/\.asp$/i.test(e.name)) files.push(path.join(d, e.name));
  }
})(root);
let ok = 0, vb = 0, ported = 0;
const pending = [];
const fails = [];
for (const f of files) {
  try {
    const r = compileAsp(f, root);
    if (r.vbscript) {
      if (fs.existsSync(path.join(root, "server/ports", path.relative(root, f)))) ported++;
      else if (/upload\.asp$/i.test(f)) vb++; // include VBScript de las páginas de subida (reemplazado por el port)
      else pending.push(path.relative(root, f));
    } else ok++;
  } catch (e) {
    const loc = e.aspLocation ? `${path.relative(root, e.aspLocation.file)}:${e.aspLocation.line}` : "?";
    fails.push(`${path.relative(root, f)}  ->  ${loc}  ${e.message.split("\n")[0]}`);
  }
}
console.log(`${files.length} páginas: ${ok} OK, ${ported} VBScript portadas a JScript (server/ports), ${vb} includes VBScript sin uso, ${pending.length} VBScript pendientes, ${fails.length} con errores`);
for (const l of pending) console.log("  pendiente: " + l);
for (const l of fails) console.log("  " + l);

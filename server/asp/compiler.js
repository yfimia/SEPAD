// Compilador de páginas ASP clásicas (JScript) a vm.Script de Node.
import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";
import { resolveCI } from "./fsci.js";
import { rewriteCallAssign } from "./jsrewrite.js";

const INCLUDE_RE = /<!--\s*#include\s+(file|virtual)\s*=\s*["']?([^"'\s>]+)["']?\s*-->/gi;
const SCRIPT_RE = /<script\b([^>]*\brunat\s*=\s*["']?server["']?[^>]*)>([\s\S]*?)<\/script\s*>/gi;

export function readLatin1(file) {
  return fs.readFileSync(file).toString("latin1");
}

/** Expande los #include y devuelve el texto con un mapa de rangos -> (archivo, línea). */
function expandIncludes(file, root, stack = []) {
  if (stack.length > 20) throw new Error("Includes anidados en exceso: " + stack.join(" > "));
  const text = readLatin1(file);
  const ranges = []; // {pos, file, line}
  let out = "";
  let last = 0;
  let lineAt = 1;
  const countLines = (s) => (s.match(/\n/g) || []).length;
  INCLUDE_RE.lastIndex = 0;
  let m;
  const pieces = [];
  while ((m = INCLUDE_RE.exec(text))) {
    pieces.push({ kind: "text", s: text.slice(last, m.index), line: lineAt });
    lineAt += countLines(text.slice(last, m.index));
    const target = m[1].toLowerCase() === "virtual" ? path.join(root, m[2]) : path.join(path.dirname(file), m[2]);
    const real = resolveCI(target);
    if (!real) {
      pieces.push({ kind: "missing", s: m[2], line: lineAt });
    } else {
      pieces.push({ kind: "inc", file: real });
    }
    lineAt += countLines(m[0]);
    last = m.index + m[0].length;
  }
  pieces.push({ kind: "text", s: text.slice(last), line: lineAt });
  for (const p of pieces) {
    if (p.kind === "text") {
      ranges.push({ pos: out.length, file, line: p.line });
      out += p.s;
    } else if (p.kind === "missing") {
      ranges.push({ pos: out.length, file, line: p.line });
      out += `<% throw __aspError("Active Server Pages", "0126", "No se encuentra el archivo de inclusión '${p.s.replace(/["\\]/g, "")}'."); %>`;
    } else {
      const sub = expandIncludes(p.file, root, [...stack, file]);
      for (const r of sub.ranges) ranges.push({ ...r, pos: r.pos + out.length });
      out += sub.text;
    }
  }
  return { text: out, ranges };
}

function locate(ranges, text, pos) {
  let r = ranges[0];
  for (const x of ranges) {
    if (x.pos <= pos) r = x;
    else break;
  }
  const extra = (text.slice(r.pos, pos).match(/\n/g) || []).length;
  return { file: r.file, line: r.line + extra };
}

/** Traductor mínimo para los bloques VBScript de SEPAD (funciones del tipo X = Now()). */
function translateVBScript(code) {
  const out = [];
  const fnRe = /\b(function|sub)\s+(\w+)\s*(\(([^)]*)\))?([\s\S]*?)\bend\s+(function|sub)\b/gi;
  let m;
  while ((m = fnRe.exec(code))) {
    const name = m[2];
    const params = (m[4] || "").split(",").map((s) => s.replace(/\b(byval|byref)\b/gi, "").trim()).filter(Boolean);
    const body = m[5].trim();
    const simple = new RegExp(`^${name}\\s*=\\s*(.+)$`, "i").exec(body);
    if (simple) {
      const expr = simple[1]
        .trim()
        .replace(/\bDate\s*\(\s*\)/gi, "__vbToday()")
        .replace(/\b(Now|Time)\b(\s*\(\s*\))?/gi, "__vbNow()")
        .replace(/&/g, "+");
      out.push(`function ${name}(${params.join(",")}) { return ${expr}; }`);
    } else {
      out.push(`function ${name}() { throw __aspError("Microsoft VBScript", "800a01f4", "VBScript no soportado en el runtime moderno: ${name}"); }`);
    }
  }
  return out.join("\n");
}

function parseDirective(s) {
  const d = {};
  s.replace(/(\w+)\s*=\s*"?([^"\s%]+)"?/g, (_, k, v) => {
    d[k.toLowerCase()] = v;
  });
  return d;
}

/**
 * Compila un archivo .asp. Devuelve { script, literals, lineMap, language }.
 */
export function compileAsp(file, root) {
  const { text, ranges } = expandIncludes(file, root);
  const literals = [];
  let gen = "";
  const lineMap = [];
  let language = "jscript";
  let sawDirective = false;

  let genLines = 0;
  const emitFast = (code, pos) => {
    const parts = code.split("\n");
    let p = pos;
    for (let i = 0; i < parts.length; i++) {
      if (i > 0) {
        gen += "\n";
        genLines++;
      }
      if (lineMap.length <= genLines) lineMap.push(locate(ranges, text, p));
      gen += parts[i];
      p += parts[i].length + 1;
    }
  };

  // segmentación
  const segs = [];
  let i = 0;
  while (i < text.length) {
    const a = text.indexOf("<%", i);
    SCRIPT_RE.lastIndex = i;
    const sm = SCRIPT_RE.exec(text);
    const s = sm ? sm.index : -1;
    if (a < 0 && s < 0) {
      segs.push({ kind: "lit", s: text.slice(i), pos: i });
      break;
    }
    if (s >= 0 && (a < 0 || s < a)) {
      if (s > i) segs.push({ kind: "lit", s: text.slice(i, s), pos: i });
      const lang = /language\s*=\s*["']?(\w+)/i.exec(sm[1]);
      const isVb = lang && /^vb/i.test(lang[1]);
      const bodyPos = s + sm[0].indexOf(">") + 1;
      segs.push({ kind: "code", s: isVb ? translateVBScript(sm[2]) : sm[2], pos: bodyPos });
      i = s + sm[0].length;
      continue;
    }
    if (a > i) segs.push({ kind: "lit", s: text.slice(i, a), pos: i });
    let e = text.indexOf("%>", a + 2);
    if (e < 0) e = text.length;
    const inner = text.slice(a + 2, e);
    if (inner.startsWith("@")) {
      const d = parseDirective(inner.slice(1));
      if (d.language) language = d.language.toLowerCase();
      sawDirective = true;
    } else if (inner.startsWith("=")) {
      segs.push({ kind: "expr", s: inner.slice(1), pos: a + 3 });
    } else {
      segs.push({ kind: "code", s: inner, pos: a + 2 });
    }
    i = e + 2;
  }

  // Sin directiva, IIS asumía VBScript: detectarlo por la sintaxis.
  if (!sawDirective && segs.some((s) => s.kind === "code" && /^\s*'|^\s*Dim\s+\w|\bEnd\s+(If|Sub|Function)\b/im.test(s.s))) {
    language = "vbscript";
  }
  if (/^vb/.test(language)) {
    return { vbscript: true, language, file };
  }

  segs.forEach((seg, k) => {
    if (seg.kind === "lit") {
      if (!seg.s) return;
      // Bloques de solo espacios entre dos bloques de código (p.ej. entre "switch {" y "case")
      const prev = segs[k - 1];
      const next = segs[k + 1];
      if (/^\s*$/.test(seg.s) && seg.s.includes("\n") && prev && prev.kind === "code" && next && next.kind === "code") return;
      literals.push(seg.s);
      emitFast(`__w(__L[${literals.length - 1}]);`, seg.pos);
      emitFast("\n", seg.pos + seg.s.length);
    } else if (seg.kind === "expr") {
      const expr = seg.s.trim().replace(/;+\s*$/, "");
      emitFast(`__w(${expr});`, seg.pos);
      emitFast("\n", seg.pos + seg.s.length);
    } else {
      emitFast(seg.s, seg.pos);
      emitFast("\n", seg.pos + seg.s.length);
    }
  });

  const code = rewriteCallAssign(gen);
  let script;
  try {
    script = new vm.Script(code, { filename: file });
  } catch (err) {
    const m = /:(\d+)\s*$/m.exec(err.stack.split("\n")[0]);
    const lineNo = m ? Number(m[1]) : 1;
    err.aspLocation = lineMap[lineNo - 1] || lineMap[0];
    err.generated = code;
    throw err;
  }
  return { script, literals, lineMap, language, file, code };
}

const cache = new Map();

/** Compilación con caché invalidada por mtime (de la página y sus includes no; reiniciar en dev). */
export function getCompiled(file, root) {
  const st = fs.statSync(file);
  const hit = cache.get(file);
  if (hit && hit.mtime === st.mtimeMs) return hit.result;
  const result = compileAsp(file, root);
  cache.set(file, { mtime: st.mtimeMs, result });
  return result;
}

// Reescritura de las construcciones de JScript que JavaScript moderno rechaza:
//   Session("x") = v;        ->  __cs(Session, ["x"], v);
//   Application("n") += 1;   ->  __cs(Application, ["n"], Application("n") + (1));

const KEYWORDS_REGEX_BEFORE = new Set(["return", "typeof", "case", "do", "else", "in", "instanceof", "new", "delete", "void", "throw"]);
const NOT_CALLEE = new Set(["if", "while", "for", "switch", "catch", "with", "function", "return", "typeof", "new", "delete", "void", "throw", "in", "case"]);
const PUNCT3 = [">>>=", "===", "!==", ">>>", "<<=", ">>="];
const PUNCT2 = ["==", "!=", "<=", ">=", "&&", "||", "++", "--", "+=", "-=", "*=", "/=", "%=", "&=", "|=", "^=", "<<", ">>", "=>"];

export function tokenize(src) {
  const toks = [];
  const n = src.length;
  let i = 0;
  let prevSig = null;
  const regexAllowed = () => {
    if (!prevSig) return true;
    if (prevSig.type === "ident") return KEYWORDS_REGEX_BEFORE.has(prevSig.value);
    if (prevSig.type === "num" || prevSig.type === "str" || prevSig.type === "regex") return false;
    if (prevSig.type === "punct") return ![")", "]"].includes(prevSig.value);
    return true;
  };
  while (i < n) {
    const c = src[i];
    const start = i;
    let type;
    if (c === " " || c === "\t" || c === "\r" || c === "\n" || c === "\f" || c === "\v" || c === " ") {
      while (i < n && /[\s ]/.test(src[i])) i++;
      type = "ws";
    } else if (c === "/" && src[i + 1] === "/") {
      while (i < n && src[i] !== "\n") i++;
      type = "comment";
    } else if (c === "/" && src[i + 1] === "*") {
      const j = src.indexOf("*/", i + 2);
      i = j < 0 ? n : j + 2;
      type = "comment";
    } else if (c === '"' || c === "'" || c === "`") {
      i++;
      while (i < n && src[i] !== c) {
        if (src[i] === "\\") i++;
        else if (src[i] === "\n" && c !== "`") break;
        i++;
      }
      i++;
      type = "str";
    } else if (c === "/" && regexAllowed()) {
      i++;
      let inClass = false;
      while (i < n) {
        const d = src[i];
        if (d === "\\") i += 2;
        else if (d === "[") (inClass = true), i++;
        else if (d === "]") (inClass = false), i++;
        else if (d === "/" && !inClass) break;
        else if (d === "\n") break;
        else i++;
      }
      i++;
      while (i < n && /[a-z]/i.test(src[i])) i++;
      type = "regex";
    } else if (/[A-Za-z_$À-￿]/.test(c)) {
      while (i < n && /[A-Za-z0-9_$À-￿]/.test(src[i])) i++;
      type = "ident";
    } else if (/[0-9]/.test(c) || (c === "." && /[0-9]/.test(src[i + 1] || ""))) {
      while (i < n && /[0-9a-fA-FxX.]/.test(src[i])) i++;
      if ((src[i - 1] === "e" || src[i - 1] === "E") && (src[i] === "+" || src[i] === "-")) {
        i++;
        while (i < n && /[0-9]/.test(src[i])) i++;
      }
      type = "num";
    } else {
      const p3 = PUNCT3.find((p) => src.startsWith(p, i));
      const p2 = !p3 && PUNCT2.find((p) => src.startsWith(p, i));
      i += p3 ? p3.length : p2 ? p2.length : 1;
      type = "punct";
    }
    const tok = { type, start, end: i, value: src.slice(start, i) };
    toks.push(tok);
    if (type !== "ws" && type !== "comment") prevSig = tok;
  }
  return toks;
}

const isSig = (t) => t.type !== "ws" && t.type !== "comment";
const CONT_AFTER = new Set(["+", "-", "*", "/", "%", "=", "==", "!=", "===", "!==", "<", ">", "<=", ">=", "&&", "||", "&", "|", "^", "?", ":", ",", "(", "[", "{", "!", "~", ".", "+=", "-="]);
const CONT_BEFORE = new Set([".", "+", "-", "*", "/", "%", "?", ":", "&&", "||", "&", "|", "^", "==", "!=", "===", "!==", "<", ">", "<=", ">=", ",", "(", "["]);

export function rewriteCallAssign(src) {
  if (!/\)\s*[+-]?=(?![=>])/.test(src)) return src;
  const toks = tokenize(src);
  // emparejar paréntesis
  const match = new Array(toks.length).fill(-1);
  const stack = [];
  toks.forEach((t, k) => {
    if (t.type !== "punct") return;
    if (t.value === "(" || t.value === "[" || t.value === "{") stack.push(k);
    else if (t.value === ")" || t.value === "]" || t.value === "}") {
      const o = stack.pop();
      if (o !== undefined) {
        match[k] = o;
        match[o] = k;
      }
    }
  });
  const nextSig = (k) => {
    for (let j = k + 1; j < toks.length; j++) if (isSig(toks[j])) return j;
    return -1;
  };
  const prevSig = (k) => {
    for (let j = k - 1; j >= 0; j--) if (isSig(toks[j])) return j;
    return -1;
  };

  const edits = [];
  for (let k = 0; k < toks.length; k++) {
    const t = toks[k];
    if (!(t.type === "punct" && t.value === ")")) continue;
    const eq = nextSig(k);
    if (eq < 0 || toks[eq].type !== "punct" || !["=", "+=", "-="].includes(toks[eq].value)) continue;
    const o = match[k];
    if (o < 0) continue;
    // callee: ident(.ident)*
    let p = prevSig(o);
    if (p < 0 || toks[p].type !== "ident") continue;
    let calleeStart = p;
    while (true) {
      const d = prevSig(calleeStart);
      if (d >= 0 && toks[d].value === ".") {
        const id = prevSig(d);
        if (id >= 0 && toks[id].type === "ident") {
          calleeStart = id;
          continue;
        }
      }
      break;
    }
    if (NOT_CALLEE.has(toks[calleeStart].value) || NOT_CALLEE.has(toks[p].value)) continue;
    // fin de la expresión de la derecha
    let depth = 0;
    let end = toks.length;
    let lastSig = eq;
    for (let j = eq + 1; j < toks.length; j++) {
      const u = toks[j];
      if (u.type === "punct") {
        if (u.value === "(" || u.value === "[" || u.value === "{") depth++;
        else if (u.value === ")" || u.value === "]" || u.value === "}") {
          if (depth === 0) {
            end = j;
            break;
          }
          depth--;
        } else if (depth === 0 && (u.value === ";" || u.value === ",")) {
          end = j;
          break;
        }
      } else if (u.type === "ws" && depth === 0 && u.value.includes("\n") && j > eq + 1 && lastSig !== eq) {
        const ls = toks[lastSig];
        const ns = nextSig(j);
        const contAfter = (ls.type === "punct" && CONT_AFTER.has(ls.value)) || (ls.type === "ident" && KEYWORDS_REGEX_BEFORE.has(ls.value));
        const contBefore = ns >= 0 && toks[ns].type === "punct" && CONT_BEFORE.has(toks[ns].value);
        if (!contAfter && !contBefore) {
          end = j;
          break;
        }
      }
      if (isSig(u)) lastSig = j;
    }
    const startPos = toks[calleeStart].start;
    const endPos = end < toks.length ? toks[end].start : src.length;
    const callee = src.slice(startPos, toks[o].start).trim();
    const args = src.slice(toks[o].end, t.start);
    const rhs = src.slice(toks[eq].end, endPos);
    edits.push({ startPos, endPos, callee, args, rhs, op: toks[eq].value });
    k = end - 1;
  }
  let out = src;
  for (let e = edits.length - 1; e >= 0; e--) {
    const { startPos, endPos, callee, args, op } = edits[e];
    const a = rewriteCallAssign(args);
    let rhs = rewriteCallAssign(edits[e].rhs);
    // mantener los saltos de línea al final fuera de la llamada
    const trail = /\s*$/.exec(rhs)[0];
    rhs = rhs.slice(0, rhs.length - trail.length);
    let value = rhs;
    if (op === "+=") value = `${callee}(${a}) + (${rhs})`;
    else if (op === "-=") value = `${callee}(${a}) - (${rhs})`;
    out = out.slice(0, startPos) + `__cs(${callee}, [${a}], ${value})${trail}` + out.slice(endPos);
  }
  return out;
}

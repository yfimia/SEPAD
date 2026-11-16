// Traducción del SQL de Access / SQL Server (como lo escribía SEPAD) a SQLite.

const pad = (n) => String(n).padStart(2, "0");

// Válido también para Date creados en otros contextos vm.
export const isDate = (v) => Object.prototype.toString.call(v) === "[object Date]";

/** Fecha "VT_DATE": como la devolvía ADO; al pasarla a texto usa el formato de Windows. */
export class VtDate extends Date {
  toString() {
    return formatVtDate(this);
  }
}

export function formatVtDate(d) {
  const h = d.getHours();
  const date = `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`;
  if (h === 0 && d.getMinutes() === 0 && d.getSeconds() === 0) return date;
  return `${date} ${h % 12 || 12}:${pad(d.getMinutes())}:${pad(d.getSeconds())} ${h < 12 ? "AM" : "PM"}`;
}

export function toSqlDate(d) {
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}

// m/d/yyyy [h:mm[:ss] [AM|PM]]  o  yyyy-mm-dd[ hh:mm[:ss]]
const US_DATE = /^\s*(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AaPp]\.?[Mm]\.?)?)?\s*$/;
const ISO_DATE = /^\s*(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?\s*$/;

export function parseLooseDate(s) {
  if (isDate(s)) return s;
  if (s == null) return null;
  s = String(s);
  let m = US_DATE.exec(s);
  if (m) {
    let [, mo, d, y, h = 0, mi = 0, se = 0, ap] = m;
    y = +y;
    if (y < 100) y += y < 30 ? 2000 : 1900;
    h = +h;
    if (ap) {
      const pm = /p/i.test(ap);
      if (pm && h < 12) h += 12;
      if (!pm && h === 12) h = 0;
    }
    return new Date(y, +mo - 1, +d, h, +mi, +se);
  }
  m = ISO_DATE.exec(s);
  if (m) {
    const [, y, mo, d, h = 0, mi = 0, se = 0] = m;
    return new Date(+y, +mo - 1, +d, +h, +mi, +se);
  }
  const t = Date.parse(s);
  return isNaN(t) ? null : new Date(t);
}

function dateLiteral(body) {
  const d = parseLooseDate(body);
  return d ? `'${toSqlDate(d)}'` : `'${body.replace(/'/g, "''")}'`;
}

const DATEPARTS = {
  year: "yyyy", yy: "yyyy", yyyy: "yyyy",
  quarter: "q", qq: "q", q: "q",
  month: "m", mm: "m", m: "m",
  dayofyear: "y", dy: "y", y: "y",
  day: "d", dd: "d", d: "d",
  week: "ww", wk: "ww", ww: "ww",
  weekday: "w", dw: "w",
  hour: "h", hh: "h",
  minute: "n", mi: "n", n: "n",
  second: "s", ss: "s", s: "s",
};

/** Tokenizador mínimo para SQL. */
function tokenize(sql) {
  const toks = [];
  let i = 0;
  const n = sql.length;
  while (i < n) {
    const c = sql[i];
    if (c === "'" || c === '"') {
      let j = i + 1;
      let s = "";
      while (j < n) {
        if (sql[j] === c) {
          if (sql[j + 1] === c) {
            s += c;
            j += 2;
            continue;
          }
          break;
        }
        s += sql[j++];
      }
      toks.push({ t: "str", v: s });
      i = j + 1;
    } else if (c === "[") {
      const j = sql.indexOf("]", i);
      const end = j < 0 ? n : j;
      toks.push({ t: "id", v: `"${sql.slice(i + 1, end).replace(/"/g, '""')}"`, raw: sql.slice(i + 1, end) });
      i = end + 1;
    } else if (c === "#") {
      const j = sql.indexOf("#", i + 1);
      if (j > 0 && j - i < 40) {
        toks.push({ t: "date", v: sql.slice(i + 1, j) });
        i = j + 1;
      } else {
        toks.push({ t: "op", v: c });
        i++;
      }
    } else if (/\s/.test(c)) {
      let j = i;
      while (j < n && /\s/.test(sql[j])) j++;
      toks.push({ t: "ws", v: " " });
      i = j;
    } else if (/[A-Za-z_À-ÿ]/.test(c)) {
      let j = i;
      while (j < n && /[A-Za-z0-9_À-ÿ$]/.test(sql[j])) j++;
      toks.push({ t: "word", v: sql.slice(i, j) });
      i = j;
    } else if (/[0-9]/.test(c)) {
      let j = i;
      while (j < n && /[0-9.]/.test(sql[j])) j++;
      toks.push({ t: "num", v: sql.slice(i, j) });
      i = j;
    } else {
      toks.push({ t: "op", v: c });
      i++;
    }
  }
  return toks;
}

const sig = (toks, i, dir) => {
  for (let j = i + dir; j >= 0 && j < toks.length; j += dir) if (toks[j].t !== "ws") return j;
  return -1;
};

/**
 * Devuelve { sql, kind, table, alias } donde kind es select|insert|update|delete|other
 * y table/alias se rellenan cuando la consulta es un SELECT editable de una sola tabla.
 */
export function translate(input) {
  let sql = String(input).trim().replace(/;\s*$/, "");
  const toks = tokenize(sql);
  let limit = null;
  const out = [];
  // (SELECT ...) UNION (SELECT ...): SQLite no admite paréntesis en los miembros de una unión.
  const skip = new Set();
  const match = new Map();
  const stack = [];
  toks.forEach((t, k) => {
    if (t.v === "(" && t.t === "op") stack.push(k);
    else if (t.v === ")" && t.t === "op" && stack.length) {
      const o = stack.pop();
      match.set(o, k);
      match.set(k, o);
    }
  });
  toks.forEach((t, k) => {
    if (!(t.t === "op" && t.v === "(" && match.has(k))) return;
    const n = sig(toks, k, 1);
    if (n < 0 || !/^select$/i.test(toks[n].v)) return;
    const p = sig(toks, k, -1);
    const close = match.get(k);
    const after = sig(toks, close, 1);
    const prevIsUnion = p >= 0 && /^(union|all)$/i.test(toks[p].v);
    const nextIsUnion = after >= 0 && /^union$/i.test(toks[after].v);
    if (prevIsUnion || (nextIsUnion && p < 0)) {
      skip.add(k);
      skip.add(close);
    }
  });
  // JOIN anidados de SQL Server: A JOIN B JOIN C ON bc ON ab  ->  A JOIN (B JOIN C ON bc) ON ab
  const openBefore = new Map();
  const closeBefore = new Map();
  {
    const stacks = [[]];
    toks.forEach((t, k) => {
      if (t.t === "op" && t.v === "(") stacks.push([]);
      else if (t.t === "op" && t.v === ")") stacks.length > 1 && stacks.pop();
      else if (t.t === "word" && /^join$/i.test(t.v)) {
        stacks[stacks.length - 1].push({ start: sig(toks, k, 1), nested: false });
      } else if (t.t === "word" && /^on$/i.test(t.v)) {
        const st = stacks[stacks.length - 1];
        const j = st.pop();
        if (!j) return;
        if (j.nested) {
          openBefore.set(j.start, (openBefore.get(j.start) || 0) + 1);
          closeBefore.set(k, (closeBefore.get(k) || 0) + 1);
        }
        if (st.length) st[st.length - 1].nested = true;
      }
    });
  }
  for (let i = 0; i < toks.length; i++) {
    const tk = toks[i];
    if (closeBefore.has(i)) out.push(")".repeat(closeBefore.get(i)) + " ");
    if (openBefore.has(i)) out.push("(".repeat(openBefore.get(i)));
    if (skip.has(i)) {
      out.push(" ");
      continue;
    }
    if (tk.t === "str") {
      // Cadenas con forma de fecha (m/d/yyyy) -> ISO para que comparen bien.
      const p = sig(toks, i, -1);
      const prev = p >= 0 ? toks[p] : null;
      const isDateish = US_DATE.test(tk.v) && prev && prev.t === "op" && /[=<>]/.test(prev.v);
      const isBetween = US_DATE.test(tk.v) && prev && prev.t === "word" && /^(between|and)$/i.test(prev.v);
      if (isDateish || isBetween) out.push(dateLiteral(tk.v));
      else out.push(`'${tk.v.replace(/'/g, "''")}'`);
    } else if (tk.t === "date") {
      out.push(dateLiteral(tk.v));
    } else if (tk.t === "word" && /^top$/i.test(tk.v) && limit === null) {
      const j = sig(toks, i, 1);
      const p = sig(toks, i, -1);
      if (j >= 0 && toks[j].t === "num" && p >= 0 && /^(select|distinct|all)$/i.test(toks[p].v)) {
        limit = toks[j].v;
        i = j;
        // también "TOP n PERCENT"
        const k = sig(toks, i, 1);
        if (k >= 0 && /^percent$/i.test(toks[k].v)) i = k;
        continue;
      }
      out.push(tk.v);
    } else if (tk.t === "word" && /^delete$/i.test(tk.v)) {
      // DELETE * FROM t   /   DELETE t.* FROM t
      out.push("DELETE");
      let j = sig(toks, i, 1);
      if (j >= 0 && toks[j].v === "*") i = j;
      else if (j >= 0 && toks[j].t === "word" && !/^from$/i.test(toks[j].v)) {
        const d = sig(toks, j, 1);
        const s = sig(toks, d, 1);
        if (d >= 0 && toks[d].v === "." && s >= 0 && toks[s].v === "*") i = s;
      }
    } else if (tk.t === "word" && /^n$/i.test(tk.v) && toks[i + 1] && toks[i + 1].t === "str") {
      // N'texto' de SQL Server
      continue;
    } else if (tk.t === "word" && /^(datediff|dateadd|datepart)$/i.test(tk.v)) {
      // DATEDIFF(minute, a, b) -> DATEDIFF('n', a, b)
      out.push(tk.v);
      const open = sig(toks, i, 1);
      const part = open >= 0 && toks[open].v === "(" ? sig(toks, open, 1) : -1;
      if (part >= 0 && toks[part].t === "word" && DATEPARTS[toks[part].v.toLowerCase()]) {
        out.push("(", `'${DATEPARTS[toks[part].v.toLowerCase()]}'`);
        i = part;
      }
    } else if (tk.t === "word" && /^convert$/i.test(tk.v)) {
      // CONVERT(char(8), x, 112) -> sqlconvert('char', 8, x, 112)
      const open = sig(toks, i, 1);
      const typ = open >= 0 && toks[open].v === "(" ? sig(toks, open, 1) : -1;
      if (typ >= 0 && toks[typ].t === "word") {
        let j = sig(toks, typ, 1);
        let len = "NULL";
        if (j >= 0 && toks[j].v === "(") {
          const num = sig(toks, j, 1);
          const close = sig(toks, num, 1);
          if (num >= 0 && close >= 0 && toks[close].v === ")") {
            len = toks[num].v;
            j = sig(toks, close, 1);
          }
        }
        if (j >= 0 && toks[j].v === ",") {
          out.push(`sqlconvert('${toks[typ].v.toLowerCase()}', ${len},`);
          i = j;
          continue;
        }
      }
      out.push(tk.v);
    } else if (tk.t === "op" && tk.v === "&") {
      out.push("||");
    } else {
      out.push(tk.v);
    }
  }
  sql = out.join("");
  if (limit !== null) sql += ` LIMIT ${parseInt(limit, 10)}`;

  const first = (toks.find((t) => t.t === "word") || { v: "" }).v.toLowerCase();
  const kind = ["select", "insert", "update", "delete"].includes(first) ? first : "other";
  const res = { sql, kind, table: null, alias: null };
  if (kind === "select") {
    const m = /\bfrom\s+("(?:[^"]|"")+"|[A-Za-z_À-ÿ][\wÀ-ÿ]*)(?:\s+(?:as\s+)?(?!where\b|order\b|group\b|having\b|limit\b|union\b|inner\b|left\b|right\b|join\b)([A-Za-z_][\w]*))?\s*(where\b|order\b|limit\b|$)/i.exec(sql);
    const complex = /\b(join|group\s+by|distinct|union|having)\b|\bfrom\s*\(|,\s*\w+\s*(\w+\s*)?(where|$)/i.test(sql.replace(/\(.*?\)/gs, "()"));
    const agg = /\b(count|sum|avg|min|max)\s*\(/i.test(sql.slice(0, sql.search(/\bfrom\b/i)));
    if (m && !complex && !agg) {
      res.table = m[1].replace(/^"|"$/g, "").replace(/""/g, '"');
      res.alias = m[2] || null;
    }
  }
  return res;
}

// ------------------------------------------------------------------ funciones Access
function asDate(v) {
  if (v == null) return null;
  if (typeof v === "number") return new Date(v);
  return parseLooseDate(v);
}

const INTERVAL_MS = { s: 1e3, n: 6e4, h: 36e5, d: 864e5, y: 864e5, w: 864e5, ww: 6048e5 };

function dateDiff(interval, a, b) {
  const d1 = asDate(a);
  const d2 = asDate(b);
  if (!d1 || !d2) return null;
  interval = String(interval).toLowerCase();
  if (interval === "yyyy") return d2.getFullYear() - d1.getFullYear();
  if (interval === "q") return Math.floor(d2.getMonth() / 3) - Math.floor(d1.getMonth() / 3) + 4 * (d2.getFullYear() - d1.getFullYear());
  if (interval === "m") return d2.getMonth() - d1.getMonth() + 12 * (d2.getFullYear() - d1.getFullYear());
  if (interval === "d" || interval === "y" || interval === "w" || interval === "ww") {
    const u1 = Date.UTC(d1.getFullYear(), d1.getMonth(), d1.getDate());
    const u2 = Date.UTC(d2.getFullYear(), d2.getMonth(), d2.getDate());
    return Math.round((u2 - u1) / INTERVAL_MS[interval]);
  }
  const ms = INTERVAL_MS[interval] || 864e5;
  return Math.trunc((d2 - d1) / ms);
}

function dateAdd(interval, n, v) {
  const d = asDate(v);
  if (!d) return null;
  n = Number(n);
  interval = String(interval).toLowerCase();
  const r = new Date(d);
  if (interval === "yyyy") r.setFullYear(r.getFullYear() + n);
  else if (interval === "q") r.setMonth(r.getMonth() + 3 * n);
  else if (interval === "m") r.setMonth(r.getMonth() + n);
  else r.setTime(r.getTime() + n * (INTERVAL_MS[interval] || 864e5));
  return toSqlDate(r);
}

export function registerFunctions(db) {
  const today = () => {
    const d = new Date();
    d.setHours(0, 0, 0, 0);
    return toSqlDate(d);
  };
  const f = (name, fn) => db.function(name, { varargs: true, deterministic: false }, fn);
  f("now", () => toSqlDate(new Date()));
  f("getdate", () => toSqlDate(new Date()));
  f("date", (...a) => {
    if (!a.length) return today();
    const d = asDate(a[0]);
    return d ? toSqlDate(d).slice(0, 10) : null;
  });
  f("time", () => toSqlDate(new Date()).slice(11));
  f("year", (v) => asDate(v)?.getFullYear() ?? null);
  f("month", (v) => (asDate(v) ? asDate(v).getMonth() + 1 : null));
  f("day", (v) => asDate(v)?.getDate() ?? null);
  f("hour", (v) => asDate(v)?.getHours() ?? null);
  f("minute", (v) => asDate(v)?.getMinutes() ?? null);
  f("weekday", (v) => (asDate(v) ? asDate(v).getDay() + 1 : null));
  f("datediff", dateDiff);
  f("dateadd", dateAdd);
  f("datevalue", (v) => (asDate(v) ? toSqlDate(asDate(v)).slice(0, 10) + " 00:00:00" : null));
  f("cdate", (v) => (asDate(v) ? toSqlDate(asDate(v)) : null));
  f("lcase", (s) => (s == null ? null : String(s).toLowerCase()));
  f("ucase", (s) => (s == null ? null : String(s).toUpperCase()));
  f("len", (s) => (s == null ? null : String(s).length));
  f("mid", (s, start, len) => (s == null ? null : String(s).substr(Number(start) - 1, len == null ? undefined : Number(len))));
  f("left", (s, n) => (s == null ? null : String(s).slice(0, Number(n))));
  f("right", (s, n) => (s == null ? null : Number(n) <= 0 ? "" : String(s).slice(-Number(n))));
  f("iif", (c, a, b) => (c ? a : b));
  f("nz", (v, d) => (v == null ? (d === undefined ? "" : d) : v));
  f("isnull", (v, d) => (d === undefined ? (v == null ? 1 : 0) : v == null ? d : v));
  f("cstr", (v) => (v == null ? null : String(v)));
  f("cint", (v) => Math.round(Number(v)));
  f("clng", (v) => Math.round(Number(v)));
  f("val", (v) => parseFloat(v) || 0);
  f("format", (v) => (v == null ? null : String(v)));
  f("datepart", (part, v) => {
    const d = asDate(v);
    if (!d) return null;
    return { yyyy: d.getFullYear(), q: Math.floor(d.getMonth() / 3) + 1, m: d.getMonth() + 1, d: d.getDate(), h: d.getHours(), n: d.getMinutes(), s: d.getSeconds(), w: d.getDay() + 1 }[String(part).toLowerCase()] ?? null;
  });
  // CONVERT(tipo(n), valor, estilo) de SQL Server
  f("sqlconvert", (type, len, v, style) => {
    if (v == null) return null;
    let out;
    if (/^(int|bigint|smallint|tinyint|bit)$/.test(type)) return Math.trunc(Number(v));
    if (/^(float|real|decimal|numeric|money)$/.test(type)) return Number(v);
    const d = style != null ? asDate(v) : null;
    if (d) {
      const y = d.getFullYear(), mo = pad(d.getMonth() + 1), da = pad(d.getDate());
      const t = `${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
      out = { 101: `${mo}/${da}/${y}`, 103: `${da}/${mo}/${y}`, 102: `${y}.${mo}.${da}`, 104: `${da}.${mo}.${y}`, 105: `${da}-${mo}-${y}`, 108: t, 112: `${y}${mo}${da}`, 120: `${y}-${mo}-${da} ${t}`, 121: `${y}-${mo}-${da} ${t}.000`, 23: `${y}-${mo}-${da}` }[Number(style)];
    }
    if (out === undefined) out = /^datetime/.test(type) ? (asDate(v) ? toSqlDate(asDate(v)) : null) : String(v);
    return len != null && out != null ? String(out).slice(0, Number(len)) : out;
  });
}

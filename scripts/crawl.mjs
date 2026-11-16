// Rastrea el sitio en ejecución siguiendo frames y enlaces .asp, y reporta errores ASP.
// Uso: node scripts/crawl.mjs [baseUrl] [maxPaginas] [usuario] [contraseña]
import crypto from "node:crypto";

const base = process.argv[2] || "http://localhost:8080/";
const max = Number(process.argv[3]) || 300;
const user = process.argv[4];
const pass = process.argv[5] ?? "";
let cookie = "";
const seen = new Set();
const queue = [];
const errors = [];
let ok = 0;

async function get(url, opts = {}) {
  const r = await fetch(url, { redirect: "manual", ...opts, headers: { cookie, ...(opts.headers || {}) } });
  const sc = r.headers.getSetCookie?.() || [];
  for (const c of sc) {
    const kv = c.split(";")[0];
    const name = kv.split("=")[0];
    cookie = cookie
      .split("; ")
      .filter((x) => x && !x.startsWith(name + "="))
      .concat(kv)
      .join("; ");
  }
  const buf = Buffer.from(await r.arrayBuffer());
  return { status: r.status, location: r.headers.get("location"), text: buf.toString("latin1") };
}

function enqueue(href, from) {
  if (!href || /^(javascript|mailto|#)/i.test(href)) return;
  let u;
  try {
    u = new URL(href.replace(/&amp;/g, "&"), from);
  } catch {
    return;
  }
  if (u.origin !== new URL(base).origin || !/\.asp$/i.test(u.pathname)) return;
  // evitar acciones destructivas al rastrear
  const baseName = u.pathname.split("/").pop();
  if (/delete|borra|elimina|logout|salir|abandon|confirm|acepta|salvar|actualiza|proceso|agrega|update|modif|send|upload|^del[A-Z]|matricular/i.test(baseName)) return;
  // volver a default.asp cierra la sesión del usuario (entra como invitado)
  if (user && /^\/default\.asp$/i.test(u.pathname)) return;
  const key = u.pathname.toLowerCase();
  const count = [...seen].filter((s) => s.startsWith(key + "?") || s === key).length;
  if (count >= 3) return;
  const full = u.pathname + u.search;
  if (seen.has(full.toLowerCase())) return;
  seen.add(full.toLowerCase());
  queue.push({ url: u.href, from });
}

if (user) {
  await get(base);
  const md5 = crypto.createHash("md5").update(pass).digest("hex").toUpperCase();
  const r = await get(new URL("default.asp", base).href, {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body: `userName=${encodeURIComponent(user)}&userPassword=${md5}`,
  });
  console.log("login:", r.status, r.location || "");
  for (const m of r.text.matchAll(/\bsrc\s*=\s*["']([^"']+)["']/gi)) enqueue(m[1], new URL("default.asp", base).href);
} else {
  enqueue(base + "default.asp", base);
}

while (queue.length && ok + errors.length < max) {
  const { url, from } = queue.shift();
  let r;
  try {
    r = await get(url);
  } catch (e) {
    errors.push({ url, from, msg: e.message });
    continue;
  }
  if (r.status >= 500 || /JScript (runtime|compilation) error|OLE DB Provider|Server object error|Active Server Pages error/.test(r.text)) {
    const msg = (r.text.match(/<p>([^<]*)<\/p>\s*<p>([^<]*)<\/p>\s*<p>([^<]*)<\/p>/) || []).slice(1).join(" | ");
    errors.push({ url, from, msg: msg || `HTTP ${r.status}` });
  } else ok++;
  if (r.location) enqueue(r.location, url);
  for (const m of r.text.matchAll(/\b(?:src|href|action)\s*=\s*["']?([^"'\s>]+)/gi)) enqueue(m[1], url);
  for (const m of r.text.matchAll(/["']([\w/.]+\.asp(?:\?[^"']*)?)["']/gi)) enqueue(m[1], url);
}
console.log(`\n${ok} páginas OK, ${errors.length} con error (cola restante: ${queue.length})`);
const byMsg = new Map();
for (const e of errors) {
  const k = e.msg.replace(/\d{6,}/g, "N");
  if (!byMsg.has(k)) byMsg.set(k, []);
  byMsg.get(k).push(e);
}
for (const [msg, list] of byMsg) {
  console.log(`\n[${list.length}] ${msg}`);
  for (const e of list.slice(0, 3)) console.log(`   ${e.url.replace(base, "/")}`);
}

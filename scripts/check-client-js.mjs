// Descarga páginas ya servidas (con sesión) y compila sus <script> en línea y
// manejadores on*="..." para detectar errores de sintaxis del lado del cliente.
// Uso: node scripts/check-client-js.mjs <baseUrl> <usuario> <contraseña> <ruta...>
import crypto from "node:crypto";
import vm from "node:vm";

const [base, user, pass, ...paths] = process.argv.slice(2);
let cookie = "";
const get = async (url, opts = {}) => {
  const r = await fetch(new URL(url, base), { redirect: "manual", ...opts, headers: { cookie, ...(opts.headers || {}) } });
  for (const c of r.headers.getSetCookie()) cookie = [cookie, c.split(";")[0]].filter(Boolean).join("; ");
  return Buffer.from(await r.arrayBuffer()).toString("latin1");
};
await get("/");
const home = await get("/default.asp", {
  method: "POST",
  headers: { "content-type": "application/x-www-form-urlencoded" },
  body: `userName=${user}&userPassword=${crypto.createHash("md5").update(pass).digest("hex").toUpperCase()}`,
});
const uid = /uid=(\d+)/.exec(home)?.[1];
for (const p of paths) {
  const html = await get(p.replace("{uid}", uid));
  let k = 0;
  for (const m of html.matchAll(/<script\b([^>]*)>([\s\S]*?)<\/script>/gi)) {
    k++;
    const attrs = m[1];
    if (/\bsrc\s*=/i.test(attrs) || /vbscript|text\/x-/i.test(attrs)) continue;
    try {
      new vm.Script(m[2]);
    } catch (e) {
      const line = m[2].split("\n")[Number((/:(\d+)/.exec(e.stack) || [])[1] || 1) - 1] || "";
      console.log(`${p} script#${k}: ${e.message}\n    ${line.trim().slice(0, 160)}`);
    }
  }
  for (const m of html.matchAll(/\bon\w+\s*=\s*"([^"]*)"/gi)) {
    try {
      new vm.Script(`(function(){${m[1].replace(/&quot;/g, '"').replace(/&amp;/g, "&")}})`);
    } catch (e) {
      console.log(`${p} handler: ${e.message}\n    ${m[0].slice(0, 160)}`);
    }
  }
}
console.log("listo");

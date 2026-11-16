// Abre, con sesión iniciada, cada ventana emergente (abreVentana/openwin/window.open)
// enlazada desde los frames principales y el Aula Virtual; registra errores y capturas.
// Uso: node scripts/popups.mjs <baseUrl> <dirSalida> <usuario> <contraseña>
import puppeteer from "puppeteer-core";
import fs from "node:fs";
import path from "node:path";

const [base = "http://localhost:8080/", outDir = "popups", user = "alumno", pass = "alumno"] = process.argv.slice(2);
const CHROME = process.env.CHROME || "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
fs.mkdirSync(outDir, { recursive: true });
const browser = await puppeteer.launch({ executablePath: CHROME, headless: "new", args: ["--no-sandbox"] });
const page = await browser.newPage();
await page.setViewport({ width: 1024, height: 700 });
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const frame = (name) => page.frames().find((f) => f.name() === name);

await page.goto(base, { waitUntil: "networkidle0" });
const top = frame("topFrame");
await top.type("#userName", user);
await top.type("#userPassword", pass);
await top.focus("#userPassword");
await Promise.all([page.waitForNavigation({ waitUntil: "networkidle0" }), page.keyboard.press("Enter")]);

const urls = new Map();
const collect = async (f) => {
  const found = await f.evaluate(() =>
    Array.from(document.querySelectorAll("a[href^='javascript:']")).map((a) => ({ text: a.textContent.trim() || a.querySelector("img")?.title || "", href: a.getAttribute("href") }))
  );
  for (const { text, href } of found) {
    const m = /(?:abreVentana|openwin|window\.open|abrir\w*)\s*\(([^)]*)\)/i.exec(href);
    if (!m) continue;
    const u = /'([^']+\.(?:asp|htm)[^']*)'|"([^"]+\.(?:asp|htm)[^"]*)"/i.exec(m[1]);
    if (!u) continue;
    const abs = new URL(u[1] || u[2], f.url()).href;
    const key = new URL(abs).pathname.toLowerCase();
    if (!urls.has(key)) urls.set(key, { text, url: abs });
  }
};
for (const f of page.frames()) await collect(f).catch(() => {});
// Aula Virtual del módulo 1 (barra de herramientas inferior)
const main = frame("mainFrame");
const uid = /uid=(\d+)/.exec(frame("topFrame").url())[1];
await main.goto(new URL(`/pages_esp/modulo.asp?uid=${uid}&ID=1&Name=Diplomado&state=1&cordinador=2000&grupo=100&claustro=101`, base).href, { waitUntil: "networkidle0" });
await main.goto(new URL(`/pages_esp/Course.asp?uid=${uid}&ID=1&Name=Introducción a Internet`, base).href, { waitUntil: "networkidle0" });
for (const f of page.frames()) await collect(f).catch(() => {});
await main.goto(new URL(`/pages_esp/lessonBrowser.asp?uid=${uid}&init=0&ID=1&Name=Introducción a Internet`, base).href, { waitUntil: "networkidle0" });
await sleep(1500);
for (const f of page.frames()) await collect(f).catch(() => {});

console.log(`${urls.size} ventanas encontradas\n`);
let n = 0;
for (const { text, url } of urls.values()) {
  const p = await browser.newPage();
  await p.setViewport({ width: 760, height: 520 });
  const errs = [];
  p.on("pageerror", (e) => errs.push("js: " + e.message.split("\n")[0]));
  p.on("response", (r) => r.status() >= 400 && !/favicon/.test(r.url()) && errs.push(`http ${r.status()} ${r.url().replace(base, "/")}`));
  try {
    await p.goto(url, { waitUntil: "networkidle0", timeout: 10000 });
  } catch (e) {
    errs.push("nav: " + e.message.split("\n")[0]);
  }
  await sleep(400);
  const body = await p.evaluate(() => document.body ? document.body.innerText.slice(0, 2000) : "").catch(() => "");
  const aspErr = /(JScript|VBScript) (runtime|compilation) error|OLE DB Provider|Server object error|todavía no se ha portado/.exec(body);
  if (aspErr) errs.push("asp: " + body.split("\n").filter(Boolean).slice(0, 3).join(" | ").slice(0, 200));
  const file = path.join(outDir, `${String(++n).padStart(2, "0")}-${new URL(url).pathname.split("/").pop().replace(/\W+/g, "_")}.png`);
  await p.screenshot({ path: file }).catch(() => {});
  console.log(`${errs.length ? "✗" : "✓"} ${text.slice(0, 28).padEnd(28)} ${new URL(url).pathname}${errs.length ? "\n     " + errs.join("\n     ") : ""}`);
  await p.close();
}
await browser.close();

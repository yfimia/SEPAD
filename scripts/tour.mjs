// Recorrido guiado con Chrome: inicia sesión y navega como lo haría un usuario,
// guardando capturas y errores de JavaScript de todos los frames y ventanas.
// Uso: node scripts/tour.mjs <baseUrl> <dirSalida> <usuario> <contraseña>
import puppeteer from "puppeteer-core";
import fs from "node:fs";
import path from "node:path";

const [base = "http://localhost:8080/", outDir = "tour", user = "alumno", pass = "alumno"] = process.argv.slice(2);
const CHROME = process.env.CHROME || "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
fs.mkdirSync(outDir, { recursive: true });
const browser = await puppeteer.launch({ executablePath: CHROME, headless: "new", args: ["--no-sandbox"] });
const logs = [];
const watch = (p, label) => {
  p.on("pageerror", (e) => logs.push(`[${label}] pageerror: ${e.message.split("\n")[0]}`));
  p.on("console", (m) => m.type() === "error" && !/favicon/.test(m.text()) && logs.push(`[${label}] console: ${m.text()}`));
  p.on("response", (r) => r.status() >= 400 && !/favicon/.test(r.url()) && logs.push(`[${label}] http ${r.status()} ${r.url()}`));
};
browser.on("targetcreated", async (t) => {
  const p = await t.page();
  if (p) {
    watch(p, "popup");
    p.setViewport({ width: 800, height: 600 }).catch(() => {});
  }
});
const page = await browser.newPage();
watch(page, "main");
await page.setViewport({ width: 1024, height: 700 });
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const frame = (name) => page.frames().find((f) => f.name() === name);
let n = 0;
const shot = async (label, p = page) => {
  await sleep(700);
  const file = path.join(outDir, `${String(++n).padStart(2, "0")}-${label}.png`);
  await p.screenshot({ path: file });
  console.log("captura:", file);
};
// Los clics sintéticos sobre enlaces javascript: dentro de frames no siempre se
// ejecutan en Chrome headless: se evalúa el código del enlace en su frame.
const follow = async (f, selector) => {
  const el = await f.waitForSelector(selector, { timeout: 4000 });
  const href = await el.evaluate((a) => a.getAttribute("href") || "");
  if (/^javascript:/i.test(href)) await f.evaluate((code) => window.eval(decodeURIComponent(code)), href.slice(11));
  else await el.evaluate((a) => (location.href = a.href));
  await sleep(1800);
};
const step = async (label, fn) => {
  try {
    await fn();
  } catch (e) {
    logs.push(`[paso ${label}] ${e.message.split("\n")[0]}`);
  }
};

await page.goto(base, { waitUntil: "networkidle0" });
await step("login", async () => {
  const top = frame("topFrame");
  await top.type("#userName", user);
  await top.type("#userPassword", pass);
  // Enter en la contraseña: el manejador original calcula el MD5 y envía el formulario
  await Promise.all([page.waitForNavigation({ waitUntil: "networkidle0" }), top.focus("#userPassword").then(() => page.keyboard.press("Enter"))]);
});
await shot("portada");

await step("modalidad", async () => {
  await follow(frame("leftFrame"), "a[href*='goGenMod'], a[href*='modulo.asp']");
});
await shot("modalidad");

await step("modulo", async () => {
  const main = frame("mainFrame");
  const links = await main.$$eval("a", (as) => as.map((a) => a.getAttribute("href") || "").filter((h) => /course/i.test(h)));
  logs.push(`[info] enlaces a módulos: ${links.slice(0, 2).join(" | ")}`);
  await follow(main, "a[href*='ourse']");
});
await shot("modulo-prologo");

await step("lecciones", async () => {
  const main = frame("mainFrame");
  const acts = await main.$$eval("a", (as) => as.map((a) => (a.textContent.trim() + " -> " + (a.getAttribute("href") || "")).slice(0, 120)));
  logs.push(`[info] acciones del módulo: ${acts.slice(0, 6).join(" | ")}`);
  await follow(main, "a[href*='contClick'], a[href*='inicioClick']");
});
await shot("leccion");

console.log("\n" + (logs.join("\n") || "(sin errores)"));
console.log("\nframes:", page.frames().map((f) => `${f.name() || "-"}=${f.url().replace(base, "/")}`).join("\n  "));
await browser.close();

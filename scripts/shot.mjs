// Captura de pantalla con Chrome y registro de errores de JavaScript del cliente.
// Uso: node scripts/shot.mjs <url> <salida.png> [usuario contraseña]
import puppeteer from "puppeteer-core";
import crypto from "node:crypto";

const [url, out, user, pass = ""] = process.argv.slice(2);
const CHROME = process.env.CHROME || "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const browser = await puppeteer.launch({ executablePath: CHROME, headless: "new", args: ["--no-sandbox"] });
const page = await browser.newPage();
await page.setViewport({ width: 1024, height: 700 });
const logs = [];
page.on("console", (m) => (m.type() === "error" || m.type() === "warning") && logs.push(`[console.${m.type()}] ${m.text()}`));
page.on("pageerror", (e) => logs.push(`[pageerror] ${e.message.split("\n")[0]}`));
page.on("frameattached", (f) => f.page() && 0);
page.on("requestfailed", (r) => logs.push(`[requestfailed] ${r.url()}`));
page.on("response", (r) => r.status() >= 400 && logs.push(`[http ${r.status()}] ${r.url()}`));
if (user) {
  await page.goto(new URL("/", url).href, { waitUntil: "networkidle0" });
  const md5 = crypto.createHash("md5").update(pass).digest("hex").toUpperCase();
  await page.evaluate((u, p) => {
    const f = document.createElement("form");
    f.method = "POST"; f.action = "/default.asp";
    f.innerHTML = `<input name="userName" value="${u}"><input name="userPassword" value="${p}">`;
    document.body ? document.body.appendChild(f) : document.documentElement.appendChild(f);
    f.submit();
  }, user, md5);
  await page.waitForNavigation({ waitUntil: "networkidle0" }).catch(() => {});
  if (url !== new URL("/", url).href) await page.goto(url, { waitUntil: "networkidle0" });
} else {
  await page.goto(url, { waitUntil: "networkidle0" });
}
await new Promise((r) => setTimeout(r, 800));
await page.screenshot({ path: out });
console.log(logs.join("\n") || "(sin errores)");
await browser.close();

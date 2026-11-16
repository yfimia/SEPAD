// SEPAD — servidor moderno para la aplicación ASP clásica original.
import http from "node:http";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { openDatabase, closeDatabase } from "./asp/ado.js";
import { AspApplication } from "./asp/runtime.js";
import { resolveCI } from "./asp/fsci.js";
import { COMPAT_URL, rewriteHtml, rewriteJs, rewriteCss } from "./compat/rewrite.js";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const PORT = Number(process.env.PORT) || 8080;
const DB_FILE = process.env.SEPAD_DB || path.join(ROOT, "data/sepad.sqlite");
const MAX_BODY = 8 * 1024 * 1024;

const MIME = {
  ".htm": "text/html; charset=iso-8859-1",
  ".html": "text/html; charset=iso-8859-1",
  ".css": "text/css; charset=iso-8859-1",
  ".js": "application/javascript; charset=iso-8859-1",
  ".gif": "image/gif",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".png": "image/png",
  ".bmp": "image/bmp",
  ".ico": "image/x-icon",
  ".swf": "application/x-shockwave-flash",
  ".xml": "text/xml; charset=iso-8859-1",
  ".xsl": "text/xml; charset=iso-8859-1",
  ".txt": "text/plain; charset=iso-8859-1",
  ".pdf": "application/pdf",
  ".zip": "application/zip",
  ".rar": "application/x-rar-compressed",
  ".doc": "application/msword",
  ".ppt": "application/vnd.ms-powerpoint",
  ".mid": "audio/midi",
  ".wav": "audio/wav",
};

// Lo que IIS nunca debió servir (o el código del runtime moderno).
const BLOCKED = [
  /^\/(server|scripts|node_modules|data|database|\.git)(\/|$)/i,
  /\.(inc|asa|mdb|ldb|sqlite|scc|rar|bat|db|md|json|mjs)$/i,
  /^\/xml\/config\.xml$/i,
  /\/thumbs\.db$/i,
];
const COMPAT_JS = fs.readFileSync(path.join(ROOT, "server/compat/ie-compat.js"));
// Librerías de terceros que ya eran multinavegador: no se reescriben.
const NO_REWRITE = /\/(tinymce|wysiwyg)\//i;

function compatBody(file, buf, contentType) {
  if (NO_REWRITE.test(file)) return buf;
  if (/text\/html/.test(contentType)) return Buffer.from(rewriteHtml(buf.toString("latin1")), "latin1");
  if (/javascript/.test(contentType)) return Buffer.from(rewriteJs(buf.toString("latin1")), "latin1");
  if (/text\/css/.test(contentType)) return Buffer.from(rewriteCss(buf.toString("latin1")), "latin1");
  return buf;
}

const DEFAULT_DOCS = ["default.asp", "index.asp", "default.htm", "index.htm", "index.html"];

const SEED_DB = path.join(ROOT, "data/sepad.seed.sqlite");

/**
 * Restaura la base de datos y los cursos de demostración. El código de 1998 construye
 * el SQL concatenando cadenas, así que en una demo pública conviene reiniciar a menudo.
 */
function resetDemo() {
  if (!fs.existsSync(SEED_DB)) {
    console.error("No existe data/sepad.seed.sqlite. Ejecuta: npm run setup");
    process.exit(1);
  }
  closeDatabase();
  for (const f of [DB_FILE + "-wal", DB_FILE + "-shm"]) fs.rmSync(f, { force: true });
  fs.copyFileSync(SEED_DB, DB_FILE);
  fs.rmSync(path.join(ROOT, "Courses"), { recursive: true, force: true });
  fs.cpSync(path.join(ROOT, "demo/Courses"), path.join(ROOT, "Courses"), { recursive: true });
  openDatabase(DB_FILE);
  console.log(`[demo] datos restaurados (${new Date().toLocaleString()})`);
}

if (process.env.SEPAD_RESET_ON_START === "1" || !fs.existsSync(DB_FILE)) {
  resetDemo();
} else {
  openDatabase(DB_FILE);
}
const resetHours = Number(process.env.SEPAD_RESET_HOURS) || 0;
if (resetHours > 0) setInterval(resetDemo, resetHours * 3600 * 1000).unref();
const app = new AspApplication({ root: ROOT });
app.start();

function send(res, status, headers, body) {
  res.writeHead(status, { "X-Powered-By": "ASP (SEPAD retro runtime)", ...headers });
  res.end(body);
}

function notFound(res) {
  send(
    res,
    404,
    { "Content-Type": "text/html; charset=iso-8859-1" },
    `<html><head><title>No se encuentra la página</title></head><body style="font-family:Verdana;font-size:12px"><h3>No se encuentra la página</h3><p>HTTP 404 - Archivo no encontrado<br>Internet Information Services</p></body></html>`
  );
}

const server = http.createServer((req, res) => {
  const t0 = Date.now();
  let urlPath;
  let rawQuery;
  try {
    const q = req.url.indexOf("?");
    urlPath = decodeURIComponent((q < 0 ? req.url : req.url.slice(0, q)).replace(/\+/g, "%20"));
    rawQuery = q < 0 ? "" : req.url.slice(q + 1);
  } catch {
    urlPath = req.url.split("?")[0];
    rawQuery = "";
  }
  urlPath = urlPath.replace(/\\/g, "/");
  if (urlPath === COMPAT_URL) {
    return send(res, 200, { "Content-Type": "application/javascript; charset=utf-8", "Cache-Control": "no-cache" }, COMPAT_JS);
  }
  if (urlPath.includes("\0") || BLOCKED.some((re) => re.test(urlPath))) return notFound(res);

  let file = resolveCI(path.join(ROOT, urlPath));
  if (!file || !(file === ROOT || file.startsWith(ROOT + path.sep))) return notFound(res);
  if (fs.statSync(file).isDirectory()) {
    if (!urlPath.endsWith("/")) {
      return send(res, 301, { Location: urlPath + "/" + (rawQuery ? "?" + rawQuery : "") }, "");
    }
    const doc = DEFAULT_DOCS.map((d) => resolveCI(path.join(file, d))).find(Boolean);
    if (!doc) return notFound(res);
    file = doc;
  }

  if (/\.asp$/i.test(file)) {
    const chunks = [];
    let size = 0;
    req.on("data", (c) => {
      size += c.length;
      if (size > MAX_BODY) {
        req.destroy();
        return;
      }
      chunks.push(c);
    });
    req.on("end", () => {
      const result = app.execute({
        root: ROOT,
        file,
        urlPath,
        rawQuery,
        method: req.method,
        headers: req.headers,
        contentType: req.headers["content-type"] || "",
        body: Buffer.concat(chunks),
        remoteAddr: req.socket.remoteAddress,
        secure: req.headers["x-forwarded-proto"] === "https",
      });
      send(res, result.status, result.headers, compatBody(file, result.body, result.headers["Content-Type"] || ""));
      if (process.env.SEPAD_LOG !== "0") console.log(`${req.method} ${req.url} ${result.status} ${Date.now() - t0}ms`);
    });
    return;
  }

  const ext = path.extname(file).toLowerCase();
  const st = fs.statSync(file);
  const etag = `"${st.size.toString(16)}-${Math.floor(st.mtimeMs).toString(16)}"`;
  if (req.headers["if-none-match"] === etag) return send(res, 304, { ETag: etag }, "");
  const type = MIME[ext] || "application/octet-stream";
  const body = compatBody(file, fs.readFileSync(file), type);
  send(res, 200, { "Content-Type": type, "Content-Length": body.length, ETag: etag, "Cache-Control": "public, max-age=300" }, body);
});

server.listen(PORT, () => {
  console.log(`SEPAD escuchando en http://localhost:${PORT}/`);
});

// Componentes COM usados por SEPAD: Scripting.FileSystemObject, Microsoft.XMLDOM,
// aspSmartUpload.SmartUpload y RWMS.RWMailSender.
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { execFileSync } from "node:child_process";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import xpath from "xpath";
import { com, prim, makeError } from "./com.js";
import { resolveCI, mapWindowsPath } from "./fsci.js";
import { VtDate } from "./sql.js";

// ------------------------------------------------------------------ FileSystemObject
function within(root, p) {
  const r = path.resolve(p);
  if (r !== root && !r.startsWith(root + path.sep)) throw makeError(0x800a0046, "Permiso denegado: " + p);
  return r;
}

class TextStreamImpl {
  constructor(file, mode) {
    this.__file = file;
    this.__mode = mode; // 1 lectura, 2 escritura, 8 añadir
    this.__buf = mode === 1 ? fs.readFileSync(file).toString("latin1") : "";
    this.__pos = 0;
  }
  get AtEndOfStream() {
    return this.__pos >= this.__buf.length;
  }
  ReadAll() {
    const s = this.__buf.slice(this.__pos);
    this.__pos = this.__buf.length;
    return s;
  }
  ReadLine() {
    let e = this.__buf.indexOf("\n", this.__pos);
    if (e < 0) e = this.__buf.length;
    const line = this.__buf.slice(this.__pos, e).replace(/\r$/, "");
    this.__pos = e + 1;
    return line;
  }
  Read(n) {
    const s = this.__buf.substr(this.__pos, n);
    this.__pos += s.length;
    return s;
  }
  Write(s) {
    this.__buf += String(prim(s));
  }
  WriteLine(s = "") {
    this.__buf += String(prim(s)) + "\r\n";
  }
  WriteBlankLines(n) {
    this.__buf += "\r\n".repeat(n);
  }
  Close() {
    if (this.__mode === 2) fs.writeFileSync(this.__file, Buffer.from(this.__buf, "latin1"));
    else if (this.__mode === 8) fs.appendFileSync(this.__file, Buffer.from(this.__buf, "latin1"));
    this.__buf = "";
  }
}

class FileImpl {
  constructor(p) {
    this.__p = p;
  }
  get Name() {
    return path.basename(this.__p);
  }
  get Path() {
    return this.__p;
  }
  get Size() {
    return fs.statSync(this.__p).size;
  }
  get DateLastModified() {
    return new VtDate(fs.statSync(this.__p).mtimeMs);
  }
  get DateCreated() {
    return new VtDate(fs.statSync(this.__p).birthtimeMs);
  }
  get Type() {
    return path.extname(this.__p).slice(1).toUpperCase() + " File";
  }
  Delete() {
    fs.rmSync(this.__p, { force: true });
  }
  toString() {
    return this.__p;
  }
}

class CollectionImpl {
  constructor(items, key) {
    this.__list = items;
    this.__key = key;
  }
  get Count() {
    return this.__list.length;
  }
  Item(k) {
    if (typeof k === "number") return this.__list[k];
    return this.__list.find((x) => this.__key(x).toLowerCase() === String(k).toLowerCase());
  }
  __call(k) {
    return this.Item(k);
  }
  __items() {
    return this.__list;
  }
}

class FolderImpl extends FileImpl {
  get Size() {
    const sum = (d) =>
      fs.readdirSync(d, { withFileTypes: true }).reduce((a, e) => a + (e.isDirectory() ? sum(path.join(d, e.name)) : fs.statSync(path.join(d, e.name)).size), 0);
    return sum(this.__p);
  }
  get Files() {
    const list = fs
      .readdirSync(this.__p, { withFileTypes: true })
      .filter((e) => e.isFile())
      .map((e) => com(new FileImpl(path.join(this.__p, e.name))));
    return com(new CollectionImpl(list, (f) => f.Name));
  }
  get SubFolders() {
    const list = fs
      .readdirSync(this.__p, { withFileTypes: true })
      .filter((e) => e.isDirectory())
      .map((e) => com(new FolderImpl(path.join(this.__p, e.name))));
    return com(new CollectionImpl(list, (f) => f.Name));
  }
  Delete() {
    fs.rmSync(this.__p, { recursive: true, force: true });
  }
}

class FileSystemObjectImpl {
  constructor(root) {
    this.__root = root;
  }
  __p(p) {
    const s = mapWindowsPath(String(prim(p)));
    return within(this.__root, resolveCI(s) || s);
  }
  FileExists(p) {
    try {
      const r = this.__p(p);
      return fs.existsSync(r) && fs.statSync(r).isFile();
    } catch {
      return false;
    }
  }
  FolderExists(p) {
    try {
      const r = this.__p(p);
      return fs.existsSync(r) && fs.statSync(r).isDirectory();
    } catch {
      return false;
    }
  }
  CreateFolder(p) {
    const r = this.__p(p);
    fs.mkdirSync(r, { recursive: true });
    return com(new FolderImpl(r));
  }
  DeleteFile(p) {
    fs.rmSync(this.__p(p), { force: true });
  }
  DeleteFolder(p) {
    fs.rmSync(this.__p(p), { recursive: true, force: true });
  }
  CopyFile(a, b) {
    fs.copyFileSync(this.__p(a), this.__p(b));
  }
  MoveFile(a, b) {
    fs.renameSync(this.__p(a), this.__p(b));
  }
  GetFile(p) {
    const r = this.__p(p);
    if (!fs.existsSync(r)) throw makeError(0x800a0035, "No se encontró el archivo");
    return com(new FileImpl(r));
  }
  GetFolder(p) {
    const r = this.__p(p);
    if (!fs.existsSync(r)) throw makeError(0x800a004c, "No se encontró la ruta de acceso");
    return com(new FolderImpl(r));
  }
  OpenTextFile(p, mode = 1, create = false) {
    const r = this.__p(p);
    if (mode === 1 && !fs.existsSync(r)) throw makeError(0x800a0035, "No se encontró el archivo");
    if (mode !== 1 && !fs.existsSync(r) && !create && mode !== 8) throw makeError(0x800a0035, "No se encontró el archivo");
    return com(new TextStreamImpl(r, mode));
  }
  CreateTextFile(p) {
    return com(new TextStreamImpl(this.__p(p), 2));
  }
  GetExtensionName(p) {
    return path.extname(String(prim(p))).slice(1);
  }
  GetFileName(p) {
    return path.basename(String(prim(p)));
  }
  GetBaseName(p) {
    const s = String(prim(p));
    return path.basename(s, path.extname(s));
  }
  GetParentFolderName(p) {
    return path.dirname(String(prim(p)));
  }
  BuildPath(a, b) {
    return path.join(String(prim(a)), String(prim(b)));
  }
  GetAbsolutePathName(p) {
    return path.resolve(String(prim(p)));
  }
}

export const createFSO = (root) => com(new FileSystemObjectImpl(root));

// ------------------------------------------------------------------ Microsoft.XMLDOM
// Extensiones MSXML sobre los nodos de xmldom: .text, .xml, selectSingleNode, selectNodes.
(function patchXmldom() {
  const doc = new DOMParser().parseFromString("<a b='c'>t</a>", "text/xml");
  const protos = new Set();
  for (let n of [doc, doc.documentElement, doc.documentElement.firstChild, doc.documentElement.getAttributeNode("b")]) {
    let p = Object.getPrototypeOf(n);
    while (p && p !== Object.prototype) {
      protos.add(p);
      p = Object.getPrototypeOf(p);
    }
  }
  const nodeProto = [...protos].find((p) => Object.prototype.hasOwnProperty.call(p, "appendChild") && Object.prototype.hasOwnProperty.call(p, "insertBefore"));
  const target = nodeProto || Object.getPrototypeOf(Object.getPrototypeOf(doc.documentElement));
  const ser = new XMLSerializer();
  const nodeList = (arr) => {
    arr.item = (i) => arr[i] ?? null;
    arr.nextNode = function () {
      this.__i = (this.__i ?? -1) + 1;
      return this[this.__i] ?? null;
    };
    arr.reset = function () {
      this.__i = -1;
    };
    return arr;
  };
  Object.defineProperties(target, {
    text: {
      configurable: true,
      get() {
        return this.textContent ?? "";
      },
      set(v) {
        this.textContent = String(prim(v));
      },
    },
    xml: {
      configurable: true,
      get() {
        return ser.serializeToString(this);
      },
    },
    nodeTypedValue: {
      configurable: true,
      get() {
        return this.textContent;
      },
    },
    selectSingleNode: {
      configurable: true,
      value(expr) {
        try {
          return xpath.select1(String(prim(expr)), this) || null;
        } catch {
          return null;
        }
      },
    },
    selectNodes: {
      configurable: true,
      value(expr) {
        try {
          return nodeList(xpath.select(String(prim(expr)), this));
        } catch {
          return nodeList([]);
        }
      },
    },
  });
})();

const XSLT_NS = "http://www.w3.org/1999/XSL/Transform";

/**
 * XSL de IE5 (http://www.w3.org/TR/WD-xsl, 1998) -> XSLT 1.0:
 *  - sin reglas implícitas (lo que no tiene plantilla no se emite);
 *  - <xsl:template> anidados dentro de <xsl:apply-templates> pasan a ser
 *    plantillas con un modo propio;
 *  - <xsl:value-of/> sin select equivale a select=".".
 */
export function wdxslToXslt(xsl) {
  if (!/WD-xsl/.test(xsl)) return xsl;
  const doc = new DOMParser().parseFromString(xsl.replace(/http:\/\/www\.w3\.org\/TR\/WD-xsl/g, XSLT_NS), "text/xml");
  const sheet = doc.documentElement;
  if (!sheet.getAttribute("version")) sheet.setAttribute("version", "1.0");
  const els = (name) => Array.from(doc.getElementsByTagNameNS(XSLT_NS, name));
  let modeN = 0;
  for (;;) {
    const at = els("apply-templates").find((a) => Array.from(a.childNodes).some((c) => c.localName === "template"));
    if (!at) break;
    const mode = "wd" + ++modeN;
    at.setAttribute("mode", mode);
    for (const t of Array.from(at.childNodes).filter((c) => c.localName === "template")) {
      at.removeChild(t);
      t.setAttribute("mode", mode);
      sheet.appendChild(t);
    }
    // Plantillas por defecto del modo: elementos sin plantilla no emiten nada,
    // salvo texto (igual que la plantilla global text() de la hoja).
    const noop = doc.createElementNS(XSLT_NS, "xsl:template");
    noop.setAttribute("match", "*|@*");
    noop.setAttribute("mode", mode);
    noop.setAttribute("priority", "-10");
    sheet.appendChild(noop);
  }
  for (const v of els("value-of")) if (!v.getAttribute("select")) v.setAttribute("select", ".");
  const def = doc.createElementNS(XSLT_NS, "xsl:template");
  def.setAttribute("match", "*|text()|@*");
  def.setAttribute("priority", "-10");
  sheet.insertBefore(def, sheet.firstChild);
  return new XMLSerializer().serializeToString(doc);
}

class XmlDocumentImpl {
  constructor(ctx) {
    this.__ctx = ctx;
    this.async = true;
    this.validateOnParse = true;
    this.preserveWhiteSpace = false;
    this.resolveExternals = true;
    this.__doc = null;
    this.__src = null;
    this.parseError = { errorCode: 0, reason: "", line: 0, linepos: 0, srcText: "", url: "" };
  }
  __parse(text) {
    const errors = [];
    const parser = new DOMParser({
      onError: (level, msg) => {
        if (level !== "warning") errors.push(msg);
      },
    });
    let doc = null;
    try {
      doc = parser.parseFromString(text, "text/xml");
    } catch (e) {
      errors.push(e.message);
    }
    if (errors.length || !doc || !doc.documentElement) {
      this.__doc = null;
      this.parseError = { errorCode: -1072896760, reason: errors.join("; ") || "Documento XML no válido", line: 0, linepos: 0, srcText: "", url: "" };
      return false;
    }
    this.__doc = doc;
    this.parseError = { errorCode: 0, reason: "", line: 0, linepos: 0, srcText: "", url: "" };
    return true;
  }
  load(p) {
    const s = mapWindowsPath(String(prim(p)));
    const real = resolveCI(s);
    if (!real || !fs.existsSync(real)) {
      this.__doc = null;
      this.parseError = { errorCode: -2146697210, reason: "El sistema no puede encontrar el archivo especificado.", url: s };
      return false;
    }
    this.__src = real;
    const buf = fs.readFileSync(real);
    const head = buf.subarray(0, 200).toString("latin1");
    const enc = /encoding=["']([^"']+)/i.exec(head);
    const text = enc && /utf-?8/i.test(enc[1]) ? buf.toString("utf8") : buf.toString("latin1");
    return this.__parse(text);
  }
  loadXML(s) {
    return this.__parse(String(prim(s)));
  }
  save(p) {
    if (!this.__doc) return;
    let s = mapWindowsPath(String(prim(p)));
    s = resolveCI(s) || s;
    fs.writeFileSync(s, Buffer.from(new XMLSerializer().serializeToString(this.__doc), "latin1"));
  }
  setProperty() {}
  get documentElement() {
    return this.__doc ? this.__doc.documentElement : null;
  }
  get xml() {
    return this.__doc ? new XMLSerializer().serializeToString(this.__doc) : "";
  }
  get text() {
    return this.__doc ? this.__doc.documentElement.textContent : "";
  }
  get childNodes() {
    return this.__doc ? this.__doc.childNodes : [];
  }
  get firstChild() {
    return this.__doc ? this.__doc.firstChild : null;
  }
  __ensure() {
    if (!this.__doc) this.__doc = new DOMParser().parseFromString("<_/>", "text/xml"), this.__doc.removeChild(this.__doc.documentElement);
    return this.__doc;
  }
  createElement(n) {
    return this.__ensure().createElement(String(prim(n)));
  }
  createTextNode(t) {
    return this.__ensure().createTextNode(String(prim(t)));
  }
  createAttribute(n) {
    return this.__ensure().createAttribute(String(prim(n)));
  }
  createProcessingInstruction(a, b) {
    return this.__ensure().createProcessingInstruction(String(a), String(b));
  }
  appendChild(n) {
    return this.__ensure().appendChild(n);
  }
  getElementsByTagName(n) {
    return this.__doc ? this.__doc.getElementsByTagName(String(prim(n))) : [];
  }
  selectSingleNode(x) {
    return this.__doc ? this.__doc.selectSingleNode(x) : null;
  }
  selectNodes(x) {
    return this.__doc ? this.__doc.selectNodes(x) : [];
  }
  transformNode(xslDoc) {
    const xslText = xslDoc ? String(xslDoc.xml) : "";
    if (!this.__doc || !xslText) return "";
    const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "sepad-xsl-"));
    try {
      let sheet = wdxslToXslt(xslText);
      // Salida en ISO-8859-1, como el resto de las páginas
      if (!/<xsl:output\b/.test(sheet)) sheet = sheet.replace(/(<xsl:stylesheet\b[^>]*>)/, '$1<xsl:output method="html" encoding="ISO-8859-1"/>');
      fs.writeFileSync(path.join(tmp, "s.xsl"), Buffer.from(sheet, "latin1"));
      fs.writeFileSync(path.join(tmp, "d.xml"), Buffer.from(this.xml, "latin1"));
      const out = execFileSync("xsltproc", ["--novalid", path.join(tmp, "s.xsl"), path.join(tmp, "d.xml")]);
      return out.toString("latin1").replace(/^<\?xml[^>]*\?>\s*/, "");
    } catch (e) {
      throw makeError(0x80004005, "Error en la transformación XSL: " + (e.stderr ? e.stderr.toString() : e.message));
    } finally {
      fs.rmSync(tmp, { recursive: true, force: true });
    }
  }
}

export const createXmlDom = (ctx) => com(new XmlDocumentImpl(ctx));

// ------------------------------------------------------------------ aspSmartUpload
export function parseMultipart(body, contentType) {
  const m = /boundary=(?:"([^"]+)"|([^;]+))/i.exec(contentType || "");
  if (!m) return { fields: [], files: [] };
  const boundary = Buffer.from("--" + (m[1] || m[2]).trim(), "latin1");
  const fields = [];
  const files = [];
  let pos = body.indexOf(boundary);
  while (pos >= 0) {
    pos += boundary.length;
    if (body[pos] === 0x2d && body[pos + 1] === 0x2d) break; // --
    if (body[pos] === 0x0d) pos += 2;
    const headEnd = body.indexOf("\r\n\r\n", pos, "latin1");
    if (headEnd < 0) break;
    const head = body.subarray(pos, headEnd).toString("latin1");
    const next = body.indexOf(boundary, headEnd + 4);
    if (next < 0) break;
    const data = body.subarray(headEnd + 4, next - 2);
    const name = /name="([^"]*)"/i.exec(head)?.[1] ?? "";
    const filename = /filename="([^"]*)"/i.exec(head);
    const ctype = /content-type:\s*([^\r\n]+)/i.exec(head)?.[1] ?? "application/octet-stream";
    if (filename) files.push({ name, filename: filename[1].split(/[\\/]/).pop(), filePathName: filename[1], contentType: ctype, data });
    else fields.push({ name, value: data.toString("latin1") });
    pos = next;
  }
  return { fields, files };
}

class UploadFileImpl {
  constructor(f, up) {
    this.__f = f;
    this.__up = up;
  }
  get Name() {
    return this.__f ? this.__f.name : "";
  }
  get FileName() {
    return this.__f ? this.__f.filename : "";
  }
  get FilePathName() {
    return this.__f ? this.__f.filePathName : "";
  }
  get FileExt() {
    return this.__f ? path.extname(this.__f.filename).slice(1) : "";
  }
  get Size() {
    return this.__f ? this.__f.data.length : 0;
  }
  get ContentType() {
    return this.__f ? this.__f.contentType : "";
  }
  get IsMissing() {
    return !this.__f || !this.__f.filename;
  }
  SaveAs(dest) {
    if (!this.__f) return;
    const target = this.__up.__ctx.mapPath(String(prim(dest)));
    const real = within(this.__up.__ctx.root, target);
    if (/\.(asp|asa|inc|php|pl|dll|exe|js|mjs|sh)$/i.test(real)) throw makeError(0x80004005, "Tipo de archivo no permitido");
    fs.mkdirSync(path.dirname(real), { recursive: true });
    fs.writeFileSync(real, this.__f.data);
  }
}

class SmartUploadImpl {
  constructor(ctx) {
    this.__ctx = ctx;
    this.MaxFileSize = 0;
    this.TotalMaxFileSize = 0;
    this.AllowedFilesList = "";
    this.DeniedFilesList = "";
    this.__parsed = { fields: [], files: [] };
  }
  Upload() {
    this.__parsed = parseMultipart(this.__ctx.body, this.__ctx.contentType);
    const denied = String(this.DeniedFilesList || "").toLowerCase().split(",").map((s) => s.trim()).filter(Boolean);
    const allowed = String(this.AllowedFilesList || "").toLowerCase().split(",").map((s) => s.trim()).filter(Boolean);
    for (const f of this.__parsed.files) {
      if (!f.filename) continue;
      const ext = path.extname(f.filename).slice(1).toLowerCase();
      if (denied.includes(ext) || (allowed.length && !allowed.includes(ext))) throw makeError(0x80040460, `La extensión del archivo (${ext}) no está permitida.`);
      if (this.MaxFileSize > 0 && f.data.length > this.MaxFileSize) throw makeError(0x80040456, "El archivo excede el tamaño máximo permitido.");
    }
  }
  get Files() {
    const up = this;
    return com({
      get Count() {
        return up.__parsed.files.length;
      },
      Item(k) {
        const f = typeof k === "number" ? up.__parsed.files[k] : up.__parsed.files.find((x) => x.name.toLowerCase() === String(prim(k)).toLowerCase());
        return com(new UploadFileImpl(f, up));
      },
      __call(k) {
        return this.Item(k);
      },
      __items() {
        return up.__parsed.files.map((f) => com(new UploadFileImpl(f, up)));
      },
      get TotalBytes() {
        return up.__parsed.files.reduce((a, f) => a + f.data.length, 0);
      },
    });
  }
  get Form() {
    const up = this;
    const item = (name) => {
      const vals = up.__parsed.fields.filter((f) => f.name.toLowerCase() === String(prim(name)).toLowerCase()).map((f) => f.value);
      return com({
        get Values() {
          return vals.join(",");
        },
        get Count() {
          return vals.length;
        },
        get Name() {
          return String(name);
        },
        __value() {
          return vals.length ? vals.join(",") : undefined;
        },
        valueOf() {
          return this.__value();
        },
        toString() {
          return String(this.__value());
        },
      });
    };
    return com({
      Item: item,
      __call: item,
      get Count() {
        return up.__parsed.fields.length;
      },
    });
  }
  Save(dir) {
    let n = 0;
    for (const f of this.__parsed.files) {
      if (!f.filename) continue;
      com(new UploadFileImpl(f, this)).SaveAs(path.posix.join(String(prim(dir)), f.filename));
      n++;
    }
    return n;
  }
  DownloadFile(p) {
    this.__ctx.download(String(prim(p)));
  }
}

export const createSmartUpload = (ctx) => com(new SmartUploadImpl(ctx));

// ------------------------------------------------------------------ RWMS.RWMailSender
// No se envía correo real: se registra en consola (servicio de 1998 apagado).
export function createMailSender() {
  const props = {};
  return new Proxy(
    {},
    {
      get(_t, k) {
        if (typeof k !== "string") return undefined;
        const l = k.toLowerCase();
        if (l in props) return props[l];
        return (...args) => {
          console.log(`[correo simulado] ${k}(${args.map((a) => JSON.stringify(String(prim(a))).slice(0, 60)).join(", ")}) to=${props.to ?? props.recipient ?? "?"}`);
          return 0;
        };
      },
      set(_t, k, v) {
        props[String(k).toLowerCase()] = prim(v);
        return true;
      },
    }
  );
}


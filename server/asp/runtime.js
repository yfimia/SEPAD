// Objetos intrínsecos de ASP (Request, Response, Server, Session, Application)
// y ejecución de páginas compiladas.
import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";
import crypto from "node:crypto";
import { com, callSet, prim, makeError, implOf, Dictionary, VBArray, makeEnumerator } from "./com.js";
import { createConnection, createRecordset, createCommand } from "./ado.js";
import { createFSO, createXmlDom, createSmartUpload, createMailSender } from "./objects.js";
import { getCompiled } from "./compiler.js";
import { resolveCI, resolveCIPartial, setSiteRoot } from "./fsci.js";
import { isDate, formatVtDate, VtDate } from "./sql.js";

const END = Symbol("Response.End");

// ------------------------------------------------------------------ codificación Latin-1
export function decodeLatin1Component(s) {
  s = s.replace(/\+/g, " ");
  if (!s.includes("%")) return s;
  const bytes = [];
  for (let i = 0; i < s.length; i++) {
    const c = s[i];
    if (c === "%" && /^[0-9a-fA-F]{2}$/.test(s.substr(i + 1, 2))) {
      bytes.push(parseInt(s.substr(i + 1, 2), 16));
      i += 2;
    } else if (c === "%" && s[i + 1] === "u" && /^[0-9a-fA-F]{4}$/.test(s.substr(i + 2, 4))) {
      // %uXXXX (escape() de IE)
      const code = parseInt(s.substr(i + 2, 4), 16);
      bytes.push(code > 255 ? 63 : code);
      i += 5;
    } else {
      const code = c.charCodeAt(0);
      if (code < 256) bytes.push(code);
      else bytes.push(...Buffer.from(c, "latin1"));
    }
  }
  const buf = Buffer.from(bytes);
  // Los navegadores modernos envían UTF-8 en URLs escritas a mano; si es UTF-8 válido, úsalo.
  const asUtf8 = buf.toString("utf8");
  if (!asUtf8.includes("�") && /[\u0080-￿]/.test(asUtf8) && Buffer.from(asUtf8, "utf8").equals(buf)) {
    return Buffer.from(asUtf8, "utf8").toString("utf8").replace(/[^\u0000-ÿ]/g, "?");
  }
  return buf.toString("latin1");
}

export function encodeLatin1Url(s, spacePlus = true) {
  let out = "";
  for (const ch of String(s)) {
    const c = ch.charCodeAt(0);
    if (/[A-Za-z0-9\-_.*]/.test(ch)) out += ch;
    else if (ch === " " && spacePlus) out += "+";
    else if (c < 256) out += "%" + c.toString(16).toUpperCase().padStart(2, "0");
    else out += "%3F";
  }
  return out;
}

/** Codifica los caracteres no ASCII de una URL (para Location) sin tocar su estructura. */
function safeLocation(url) {
  let out = "";
  for (const ch of String(url)) {
    const c = ch.charCodeAt(0);
    if (c > 126 || c < 33 || ch === '"' || ch === "<" || ch === ">") out += c < 256 ? "%" + c.toString(16).toUpperCase().padStart(2, "0") : "%3F";
    else out += ch;
  }
  return out;
}

export function parseQuery(qs) {
  const map = new Map(); // lower -> {name, values[]}
  const order = [];
  if (!qs) return { map, order };
  for (const part of qs.split("&")) {
    if (!part) continue;
    const eq = part.indexOf("=");
    const k = decodeLatin1Component(eq < 0 ? part : part.slice(0, eq));
    const v = eq < 0 ? "" : decodeLatin1Component(part.slice(eq + 1));
    const l = k.toLowerCase();
    if (!map.has(l)) {
      map.set(l, { name: k, values: [] });
      order.push(l);
    }
    map.get(l).values.push(v);
  }
  return { map, order };
}

// ------------------------------------------------------------------ colecciones
class StringListImpl {
  constructor(values, name) {
    this.__values = values;
    this.__name = name;
  }
  __value() {
    return this.__values.length ? this.__values.join(", ") : undefined;
  }
  [Symbol.toPrimitive]() {
    return this.__value();
  }
  valueOf() {
    return this.__value();
  }
  toString() {
    return String(this.__value());
  }
  get Count() {
    return this.__values.length;
  }
  Item(i) {
    if (i === undefined) return this.__value();
    return this.__values[Number(i) - 1];
  }
  __call(i) {
    return this.Item(i);
  }
  __items() {
    return this.__values;
  }
}

class RequestCollectionImpl {
  constructor(parsed, raw) {
    this.__parsed = parsed;
    this.__raw = raw;
  }
  __lookup(k) {
    k = prim(k);
    if (typeof k === "number") {
      const key = this.__parsed.order[k - 1];
      return key ? this.__parsed.map.get(key).values : [];
    }
    const e = this.__parsed.map.get(String(k).toLowerCase());
    return e ? e.values : [];
  }
  Item(k) {
    if (k === undefined) return this.__raw;
    return com(new StringListImpl(this.__lookup(k), k));
  }
  __call(k) {
    return this.Item(k);
  }
  Key(i) {
    const key = this.__parsed.order[Number(i) - 1];
    return key ? this.__parsed.map.get(key).name : undefined;
  }
  get Count() {
    return this.__parsed.order.length;
  }
  __items() {
    return this.__parsed.order.map((l) => this.__parsed.map.get(l).name);
  }
  __value() {
    return this.__raw;
  }
  [Symbol.toPrimitive]() {
    return this.__raw;
  }
  toString() {
    return this.__raw;
  }
}

class RequestImpl {
  constructor(ctx) {
    this.__ctx = ctx;
    const q = parseQuery(ctx.rawQuery);
    this.__qs = com(new RequestCollectionImpl(q, ctx.rawQuery));
    let form = { map: new Map(), order: [] };
    let rawForm = "";
    if (/application\/x-www-form-urlencoded/i.test(ctx.contentType || "") && ctx.body.length) {
      rawForm = ctx.body.toString("latin1");
      form = parseQuery(rawForm);
    }
    this.__form = com(new RequestCollectionImpl(form, rawForm));
    const cookies = { map: new Map(), order: [] };
    for (const part of (ctx.headers.cookie || "").split(/;\s*/)) {
      const eq = part.indexOf("=");
      if (eq < 0) continue;
      const k = decodeLatin1Component(part.slice(0, eq).trim());
      const l = k.toLowerCase();
      if (!cookies.map.has(l)) {
        cookies.map.set(l, { name: k, values: [] });
        cookies.order.push(l);
      }
      cookies.map.get(l).values.push(decodeLatin1Component(part.slice(eq + 1)));
    }
    this.__cookies = com(new RequestCollectionImpl(cookies, ctx.headers.cookie || ""));
    const sv = { map: new Map(), order: [] };
    const add = (k, v) => {
      sv.map.set(k.toLowerCase(), { name: k, values: [v == null ? "" : String(v)] });
      sv.order.push(k.toLowerCase());
    };
    const h = ctx.headers;
    add("REQUEST_METHOD", ctx.method);
    add("QUERY_STRING", ctx.rawQuery);
    add("SCRIPT_NAME", ctx.urlPath);
    add("URL", ctx.urlPath);
    add("PATH_INFO", ctx.urlPath);
    add("PATH_TRANSLATED", ctx.file);
    add("APPL_PHYSICAL_PATH", ctx.root + path.sep);
    add("APPL_MD_PATH", "/LM/W3SVC/1/Root");
    add("SERVER_NAME", (h.host || "localhost").split(":")[0]);
    add("SERVER_PORT", (h.host || "").split(":")[1] || "80");
    add("SERVER_PROTOCOL", "HTTP/1.1");
    add("SERVER_SOFTWARE", "Microsoft-IIS/4.0");
    add("HTTPS", ctx.secure ? "on" : "off");
    add("REMOTE_ADDR", ctx.remoteAddr);
    add("REMOTE_HOST", ctx.remoteAddr);
    add("LOCAL_ADDR", "127.0.0.1");
    add("CONTENT_TYPE", ctx.contentType || "");
    add("CONTENT_LENGTH", ctx.body.length);
    add("AUTH_USER", "");
    add("LOGON_USER", "");
    for (const [k, v] of Object.entries(h)) add("HTTP_" + k.toUpperCase().replace(/-/g, "_"), v);
    add("ALL_HTTP", Object.entries(h).map(([k, v]) => `HTTP_${k.toUpperCase().replace(/-/g, "_")}:${v}`).join("\n"));
    this.__sv = com(new RequestCollectionImpl(sv, ""));
  }
  get QueryString() {
    return this.__qs;
  }
  get Form() {
    return this.__form;
  }
  get Cookies() {
    return this.__cookies;
  }
  get ServerVariables() {
    return this.__sv;
  }
  get ClientCertificate() {
    return com(new RequestCollectionImpl({ map: new Map(), order: [] }, ""));
  }
  get TotalBytes() {
    return this.__ctx.body.length;
  }
  BinaryRead(n) {
    return this.__ctx.body.subarray(0, n);
  }
  __call(k) {
    for (const c of [this.__qs, this.__form, this.__cookies, this.__sv]) {
      const item = c.Item(k);
      if (item.Count > 0) return item;
    }
    return com(new StringListImpl([], k));
  }
}

class ResponseImpl {
  constructor(ctx) {
    this.__ctx = ctx;
    this.__out = [];
    this.__headers = {};
    this.__cookies = new Map();
    this.Buffer = true;
    this.Expires = undefined;
    this.ExpiresAbsolute = undefined;
    this.ContentType = "text/html";
    this.Charset = "";
    this.Status = "200 OK";
    this.CacheControl = "private";
    this.__binary = null;
  }
  __str(v) {
    v = prim(v);
    if (v === undefined || v === null) return "";
    if (typeof v === "boolean") return v ? "True" : "False";
    if (v instanceof VtDate) return formatVtDate(v);
    if (isDate(v)) return jscriptDateString(v);
    return String(v);
  }
  Write(v) {
    this.__out.push(this.__str(v));
  }
  BinaryWrite(b) {
    this.__binary = Buffer.concat([this.__binary || Buffer.alloc(0), Buffer.from(b)]);
  }
  Redirect(url) {
    this.Status = "302 Object moved";
    this.__headers["Location"] = safeLocation(prim(url));
    this.__out = [`<head><title>Object moved</title></head><body><h1>Object Moved</h1>This object may be found <a HREF="${prim(url)}">here</a>.</body>`];
    throw END;
  }
  End() {
    throw END;
  }
  Clear() {
    this.__out = [];
  }
  Flush() {}
  AddHeader(k, v) {
    this.__headers[String(prim(k))] = String(prim(v));
  }
  AppendToLog() {}
  get IsClientConnected() {
    return true;
  }
  get Cookies() {
    const self = this;
    return com({
      __call(k) {
        return this.Item(k);
      },
      __set(k, v) {
        self.__cookies.set(String(prim(k)), { value: String(prim(v)) });
      },
      Item(k) {
        const name = String(prim(k));
        if (!self.__cookies.has(name)) self.__cookies.set(name, { value: "" });
        const c = self.__cookies.get(name);
        return com({
          get Expires() {
            return c.expires;
          },
          set Expires(v) {
            c.expires = v;
          },
          get Path() {
            return c.path;
          },
          set Path(v) {
            c.path = v;
          },
          set Value(v) {
            c.value = String(prim(v));
          },
          get Value() {
            return c.value;
          },
        });
      },
    });
  }
}

function jscriptDateString(d) {
  // Formato de Date.toString() en JScript 5: "Mon Nov 16 06:48:00 UTC-0500 2026"
  const days = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const p = (n) => String(n).padStart(2, "0");
  const off = -d.getTimezoneOffset();
  const tz = off === 0 ? "UTC" : `UTC${off > 0 ? "+" : "-"}${p(Math.floor(Math.abs(off) / 60))}${p(Math.abs(off) % 60)}`;
  return `${days[d.getDay()]} ${months[d.getMonth()]} ${d.getDate()} ${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())} ${tz} ${d.getFullYear()}`;
}

// ------------------------------------------------------------------ Session / Application
export class Store {
  constructor() {
    this.__map = new Map(); // lower -> {key, value}
  }
  get(k) {
    const e = this.__map.get(String(prim(k)).toLowerCase());
    return e ? e.value : undefined;
  }
  set(k, v) {
    const key = String(prim(k));
    this.__map.set(key.toLowerCase(), { key, value: prim(v) });
  }
  remove(k) {
    this.__map.delete(String(prim(k)).toLowerCase());
  }
  clear() {
    this.__map.clear();
  }
  keys() {
    return [...this.__map.values()].map((e) => e.key);
  }
}

function contentsOf(store) {
  return com({
    __call: (k) => store.get(k),
    __set: (k, v) => store.set(k, v),
    Item: (k) => store.get(k),
    Remove: (k) => store.remove(k),
    RemoveAll: () => store.clear(),
    get Count() {
      return store.keys().length;
    },
    Key: (i) => store.keys()[Number(i) - 1],
    __items: () => store.keys(),
  });
}

class SessionImpl {
  constructor(session) {
    this.__s = session;
  }
  __call(k) {
    return this.__s.store.get(k);
  }
  __set(k, v) {
    this.__s.store.set(k, v);
  }
  get Contents() {
    return contentsOf(this.__s.store);
  }
  get SessionID() {
    return this.__s.numericId;
  }
  get Timeout() {
    return this.__s.timeout;
  }
  set Timeout(v) {
    this.__s.timeout = Number(prim(v)) || 20;
  }
  get LCID() {
    return 1034;
  }
  set LCID(_v) {}
  get CodePage() {
    return 1252;
  }
  set CodePage(_v) {}
  Abandon() {
    this.__s.abandoned = true;
  }
}

class ApplicationImpl {
  constructor(store) {
    this.__store = store;
  }
  __call(k) {
    return this.__store.get(k);
  }
  __set(k, v) {
    this.__store.set(k, v);
  }
  get Contents() {
    return contentsOf(this.__store);
  }
  Lock() {}
  Unlock() {}
  UnLock() {}
}

// ------------------------------------------------------------------ Server
class ServerImpl {
  constructor(ctx) {
    this.__ctx = ctx;
    this.ScriptTimeout = 90;
  }
  CreateObject(progId) {
    const id = String(prim(progId)).toLowerCase().trim();
    const f = PROGIDS[id];
    if (!f) throw makeError(0x800401f3, `Server.CreateObject ha fallado: '${prim(progId)}' (clase no registrada en el runtime moderno)`);
    return f(this.__ctx);
  }
  MapPath(p) {
    return this.__ctx.mapPath(String(prim(p)));
  }
  URLEncode(s) {
    return encodeLatin1Url(prim(s) ?? "");
  }
  URLPathEncode(s) {
    return encodeLatin1Url(prim(s) ?? "", false);
  }
  HTMLEncode(s) {
    return String(prim(s) ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }
  GetLastError() {
    return com({ Number: 0, Description: "", ASPCode: "", ASPDescription: "", Line: 0, File: "", Source: "", Category: "" });
  }
  Execute(p) {
    this.__ctx.app.executeChild(this.__ctx, this.MapPath(p));
  }
  Transfer(p) {
    this.__ctx.app.executeChild(this.__ctx, this.MapPath(p));
    throw END;
  }
}

const PROGIDS = {
  "adodb.connection": () => createConnection(),
  "adodb.recordset": () => createRecordset(),
  "adodb.command": () => createCommand(),
  "scripting.dictionary": () => com(new Dictionary()),
  "scripting.filesystemobject": (ctx) => createFSO(ctx.root),
  "microsoft.xmldom": (ctx) => createXmlDom(ctx),
  "msxml.domdocument": (ctx) => createXmlDom(ctx),
  "msxml2.domdocument": (ctx) => createXmlDom(ctx),
  "msxml2.domdocument.3.0": (ctx) => createXmlDom(ctx),
  "msxml2.freethreadeddomdocument": (ctx) => createXmlDom(ctx),
  "aspsmartupload.smartupload": (ctx) => createSmartUpload(ctx),
  "rwms.rwmailsender": () => createMailSender(),
  "cdonts.newmail": () => createMailSender(),
};

// ------------------------------------------------------------------ preludio de JScript
const PRELUDE = new vm.Script(`
  Date.prototype.getYear = function () { var y = this.getFullYear(); return y >= 1900 && y < 2000 ? y - 1900 : y; };
  Date.prototype.getVarDate = function () { return this; };
  String.prototype.anchor = String.prototype.anchor || function (n) { return '<A NAME="' + n + '">' + this + '</A>'; };
`);

// ------------------------------------------------------------------ errores al estilo IIS
function aspErrorInfo(err, compiled, root) {
  let number = "800a01b6";
  let source = "Microsoft JScript runtime error";
  let message = err && err.message ? err.message : String(err);
  if (err && err.aspSource) {
    source = err.aspSource;
    number = err.aspNumber;
  } else if (err && typeof err.number === "number" && err.number) {
    number = (err.number >>> 0).toString(16).padStart(8, "0");
    if (/^80040e/.test(number)) source = "Microsoft OLE DB Provider for ODBC Drivers error";
    if (/^800401f3/.test(number)) source = "Server object error 'ASP 0177 : 800401f3'";
  } else if (err instanceof SyntaxError || (err && err.name === "SyntaxError")) {
    source = "Microsoft JScript compilation error";
    number = "800a03ea";
  } else if (err && err.name === "ReferenceError") {
    number = "800a1391";
    const m = /^(\S+) is not defined/.exec(message);
    if (m) message = `'${m[1]}' no está definido`;
  } else if (err && err.name === "TypeError") {
    number = "800a138f";
    const m = /Cannot read propert(?:y|ies) of (null|undefined)(?: \(reading '([^']+)'\))?/.exec(message);
    if (m) message = `'${m[2] || "objeto"}' es nulo o no es un objeto`;
    else if (/is not a function/.test(message)) {
      number = "800a01b6";
      message = `El objeto no acepta esta propiedad o método (${message})`;
    }
  }
  let where = null;
  if (err && err.aspLocation) where = err.aspLocation;
  else if (compiled && err && err.stack) {
    const re = new RegExp(compiled.file.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + ":(\\d+)");
    const m = re.exec(err.stack);
    if (m) where = compiled.lineMap[Number(m[1]) - 1];
  }
  const loc = where ? `/${path.relative(root, where.file).split(path.sep).join("/")}, line ${where.line}` : "";
  return { source, number, message, loc };
}

function errorHtml(info) {
  const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;");
  return `<font face="Arial" size=2>
<p>${esc(info.source)} '${esc(info.number)}'</p>
<p>${esc(info.message).replace(/\n/g, "<br>")}</p>
<p>${esc(info.loc)}</p>
</font>`;
}

// ------------------------------------------------------------------ ejecución
export class AspApplication {
  constructor({ root }) {
    this.root = root;
    setSiteRoot(root);
    this.appStore = new Store();
    this.sessions = new Map();
    this.globalAsa = null;
  }

  start() {
    const asa = resolveCI(path.join(this.root, "global.asa"));
    if (asa) {
      this.globalAsa = getCompiled(asa, this.root);
      const ctx = this.__fakeCtx();
      const sandbox = this.__sandbox(ctx, null);
      this.__runGlobalAsa(sandbox);
      if (typeof sandbox.Application_OnStart === "function") sandbox.Application_OnStart();
    }
    // Rutas relativas a la aplicación
    this.appStore.set("filePath", this.appStore.get("filePath") || "sepad");
  }

  // global.asa solo define funciones: su texto literal no se envía al cliente.
  __runGlobalAsa(sandbox) {
    const w = sandbox.__w;
    sandbox.__w = () => {};
    sandbox.__L = this.globalAsa.literals;
    try {
      this.globalAsa.script.runInContext(sandbox);
    } finally {
      sandbox.__w = w;
    }
  }

  __fakeCtx() {
    return {
      app: this,
      root: this.root,
      file: path.join(this.root, "global.asa"),
      urlPath: "/global.asa",
      rawQuery: "",
      method: "GET",
      headers: {},
      body: Buffer.alloc(0),
      remoteAddr: "127.0.0.1",
      mapPath: (p) => this.mapPath(p, path.join(this.root, "global.asa")),
    };
  }

  /** Páginas portadas a JScript (originales en VBScript) en server/ports/<misma ruta>. */
  portFor(file) {
    const rel = path.relative(this.root, file);
    const port = resolveCI(path.join(this.root, "server/ports", rel));
    return port || file;
  }

  mapPath(p, file) {
    p = String(p).replace(/\\/g, "/");
    const base = p.startsWith("/") ? this.root : path.dirname(file);
    const abs = path.resolve(base, "." + (p.startsWith("/") ? p : "/" + p));
    return resolveCIPartial(abs);
  }

  getSession(cookieHeader) {
    const now = Date.now();
    for (const [id, s] of this.sessions) if (now - s.lastSeen > s.timeout * 60000 || s.abandoned) this.sessions.delete(id);
    const m = /ASPSESSIONID[A-Z]{8}=([A-Z]{24})/.exec(cookieHeader || "");
    if (m && this.sessions.has(m[1])) {
      const s = this.sessions.get(m[1]);
      s.lastSeen = now;
      return { session: s, isNew: false };
    }
    const id = Array.from(crypto.randomBytes(24), (b) => String.fromCharCode(65 + (b % 26))).join("");
    const s = {
      id,
      numericId: String(crypto.randomInt(100000000, 999999999)),
      store: new Store(),
      timeout: 20,
      lastSeen: now,
      abandoned: false,
    };
    this.sessions.set(id, s);
    return { session: s, isNew: true };
  }

  __sandbox(ctx, session) {
    const response = ctx.responseObj || com(new ResponseImpl(ctx));
    ctx.responseObj = response;
    const respImpl = implOf(response);
    const sandbox = {
      Request: ctx.requestObj || (ctx.requestObj = com(new RequestImpl(ctx))),
      Response: response,
      Server: com(new ServerImpl(ctx)),
      Session: session ? com(new SessionImpl(session)) : undefined,
      Application: com(new ApplicationImpl(this.appStore)),
      __w: (v) => {
        respImpl.__out.push(typeof v === "string" ? v : respImpl.__str(v));
      },
      __cs: callSet,
      __aspError: (src, num, msg) => {
        const e = new Error(msg);
        e.aspSource = src;
        e.aspNumber = num;
        return e;
      },
      // Now() de VBScript devuelve un VT_DATE
      __vbNow: () => new VtDate(),
      __vbToday: () => {
        const d = new VtDate();
        d.setHours(0, 0, 0, 0);
        return d;
      },
      Enumerator: function (c) {
        return makeEnumerator(c);
      },
      VBArray,
      ActiveXObject: function (id) {
        return sandbox.Server.CreateObject(id);
      },
      GetObject: function () {
        throw makeError(0x800a01ad, "GetObject no está disponible");
      },
      ScriptEngine: () => "JScript",
      ScriptEngineMajorVersion: () => 5,
      ScriptEngineMinorVersion: () => 6,
      ScriptEngineBuildVersion: () => 8835,
      console,
    };
    vm.createContext(sandbox);
    PRELUDE.runInContext(sandbox);
    return sandbox;
  }

  /** Ejecuta una página ASP y devuelve {status, headers, body(Buffer)} */
  execute(ctx) {
    const { session, isNew } = this.getSession(ctx.headers.cookie);
    ctx.app = this;
    ctx.mapPath = (p) => this.mapPath(p, ctx.file);
    ctx.download = (p) => {
      const real = this.mapPath(p, ctx.file);
      ctx.responseObj && implOf(ctx.responseObj).BinaryWrite(fs.readFileSync(real));
      implOf(ctx.responseObj).ContentType = "application/octet-stream";
      implOf(ctx.responseObj).AddHeader("Content-Disposition", `attachment; filename="${path.basename(real)}"`);
    };
    const sandbox = this.__sandbox(ctx, session);
    const resp = implOf(sandbox.Response);
    let compiled = null;
    try {
      if (isNew && this.globalAsa) {
        this.__runGlobalAsa(sandbox);
        if (typeof sandbox.Session_OnStart === "function") sandbox.Session_OnStart();
      }
      compiled = getCompiled(this.portFor(ctx.file), this.root);
      if (compiled.vbscript) {
        resp.Write(vbscriptNotice(ctx));
      } else {
        sandbox.__L = compiled.literals;
        compiled.script.runInContext(sandbox, { timeout: 15000 });
      }
    } catch (err) {
      if (err !== END) {
        const info = aspErrorInfo(err, compiled, this.root);
        console.error(`[ASP] ${info.source} '${info.number}' ${info.message} ${info.loc}`);
        if (process.env.SEPAD_DEBUG && err && err.stack) console.error(err.stack);
        resp.Status = "500 Internal Server Error";
        resp.__out.push(errorHtml(info));
      }
    }
    const headers = { ...resp.__headers };
    const status = parseInt(resp.Status, 10) || 200;
    let body;
    if (resp.__binary) {
      body = resp.__binary;
      headers["Content-Type"] = resp.ContentType;
    } else {
      body = Buffer.from(resp.__out.join(""), "latin1");
      const ct = String(resp.ContentType || "text/html");
      headers["Content-Type"] = /charset/i.test(ct) ? ct : `${ct}; charset=${resp.Charset || "iso-8859-1"}`;
    }
    const expires = prim(resp.Expires);
    if (expires !== undefined && Number(expires) <= 0) headers["Cache-Control"] = "no-cache";
    else headers["Cache-Control"] = String(resp.CacheControl || "private");
    const cookies = [];
    if (isNew) cookies.push(`ASPSESSIONID${sessionCookieSuffix(this)}=${session.id}; path=/; HttpOnly; SameSite=Lax`);
    for (const [k, c] of resp.__cookies) {
      let s = `${encodeLatin1Url(k)}=${encodeLatin1Url(c.value)}; path=${c.path || "/"}`;
      if (c.expires) {
        const d = new Date(prim(c.expires));
        if (!isNaN(d)) s += `; expires=${d.toUTCString()}`;
      }
      cookies.push(s);
    }
    if (cookies.length) headers["Set-Cookie"] = cookies;
    return { status, headers, body };
  }

  executeChild(ctx, file) {
    const compiled = getCompiled(file, this.root);
    const sandbox = this.__sandbox({ ...ctx, file }, this.getSession(ctx.headers.cookie).session);
    sandbox.__L = compiled.literals;
    compiled.script.runInContext(sandbox, { timeout: 15000 });
  }
}

function sessionCookieSuffix(app) {
  if (!app.__cookieSuffix) app.__cookieSuffix = Array.from(crypto.randomBytes(8), (b) => String.fromCharCode(65 + (b % 26))).join("");
  return app.__cookieSuffix;
}

function vbscriptNotice(ctx) {
  return `<html><head><title>SEPAD</title></head><body style="font-family:Verdana;font-size:12px">
<p><b>Página VBScript</b></p>
<p>Esta página (${path.basename(ctx.file)}) estaba escrita en VBScript y todavía no se ha portado al runtime moderno.</p>
<p><a href="javascript:history.back()">Regresar</a></p></body></html>`;
}

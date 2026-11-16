// ADODB (Connection, Recordset, Command, Field) sobre better-sqlite3.
// Todo es síncrono, igual que en ASP clásico.
import Database from "better-sqlite3";
import { com, prim, makeError, implOf } from "./com.js";
import { translate, registerFunctions, toSqlDate, parseLooseDate, isDate, VtDate } from "./sql.js";

let DB = null;
const tableInfoCache = new Map();

export function openDatabase(file) {
  DB = new Database(file);
  DB.pragma("journal_mode = WAL");
  DB.pragma("case_sensitive_like = OFF");
  registerFunctions(DB);
  tableInfoCache.clear();
  return DB;
}
export const db = () => DB;

export function closeDatabase() {
  if (DB) DB.close();
  DB = null;
  tableInfoCache.clear();
}

function tableInfo(table) {
  const key = table.toLowerCase();
  let info = tableInfoCache.get(key);
  if (!info) {
    const cols = DB.prepare(`PRAGMA table_info("${table.replace(/"/g, '""')}")`).all();
    if (!cols.length) return null;
    info = {
      name: table,
      cols: cols.map((c) => ({ name: c.name, type: (c.type || "").toUpperCase(), pk: !!c.pk })),
    };
    info.byLower = new Map(info.cols.map((c) => [c.name.toLowerCase(), c]));
    tableInfoCache.set(key, info);
  }
  return info;
}

// ------------------------------------------------------------------ conversiones
function fromDb(v, type) {
  if (v == null) return null;
  if (type === "BOOLEAN") return v !== 0 && v !== "0" && v !== false;
  if (type === "DATETIME") {
    const d = parseLooseDate(v);
    return d ? new VtDate(d.getTime()) : v;
  }
  if (typeof v === "bigint") return Number(v);
  return v;
}

function toDb(v, type) {
  v = prim(v);
  if (v === undefined) return null;
  if (v === null) return null;
  if (type === "BOOLEAN") {
    if (typeof v === "string") return /^(true|-1|1|on|yes|s[ií])$/i.test(v.trim()) ? 1 : 0;
    return v ? 1 : 0;
  }
  if (isDate(v)) return toSqlDate(v);
  if (type === "DATETIME" && typeof v === "string") {
    const d = parseLooseDate(v);
    return d ? toSqlDate(d) : v;
  }
  if (typeof v === "boolean") return v ? 1 : 0;
  if ((type === "INTEGER" || type === "REAL") && typeof v === "string") {
    const t = v.trim();
    if (t === "") return null;
    const n = Number(t);
    return isNaN(n) ? v : n;
  }
  if (typeof v === "object") return String(v);
  return v;
}

function bindParam(v) {
  v = prim(v);
  if (isDate(v)) return toSqlDate(v);
  if (typeof v === "boolean") return v ? 1 : 0;
  if (v === undefined) return null;
  return v;
}

function runSql(sqlText) {
  let t;
  try {
    t = translate(prim(sqlText));
    const stmt = DB.prepare(t.sql);
    return { t, stmt };
  } catch (e) {
    throw makeError(0x80040e14, `[Microsoft][ODBC Driver] ${e.message}\nSQL: ${t ? t.sql : sqlText}`);
  }
}

// ------------------------------------------------------------------ Field
class FieldImpl {
  constructor(rs, index) {
    this.__rs = rs;
    this.__i = index;
  }
  __value() {
    return this.Value;
  }
  valueOf() {
    const v = this.Value;
    return isDate(v) ? v.valueOf() : v;
  }
  toString() {
    const v = this.Value;
    return v == null ? (v === null ? "null" : "undefined") : String(v);
  }
  [Symbol.toPrimitive](hint) {
    const v = this.Value;
    if (isDate(v)) return hint === "number" ? v.valueOf() : String(v);
    return v;
  }
  get Name() {
    return this.__rs.__cols[this.__i].name;
  }
  get Type() {
    const t = this.__rs.__cols[this.__i].type;
    return { BOOLEAN: 11, DATETIME: 7, INTEGER: 3, REAL: 5, TEXT: 202 }[t] || 202;
  }
  get DefinedSize() {
    return 255;
  }
  get ActualSize() {
    const v = this.Value;
    return v == null ? 0 : String(v).length;
  }
  get Value() {
    return this.__rs.__get(this.__i);
  }
  set Value(v) {
    this.__rs.__put(this.__i, v);
  }
  get OriginalValue() {
    return this.Value;
  }
}

class FieldsImpl {
  constructor(rs) {
    this.__rs = rs;
  }
  __call(k) {
    return this.Item(k);
  }
  __set(k, v) {
    this.Item(k).Value = v;
  }
  __set_Item(k, v) {
    this.__set(k, v);
  }
  Item(k) {
    return this.__rs.__field(k);
  }
  get Count() {
    return this.__rs.__cols.length;
  }
  __items() {
    return this.__rs.__cols.map((_, i) => this.__rs.__field(i));
  }
  Refresh() {}
}

// ------------------------------------------------------------------ Recordset
class RecordsetImpl {
  constructor() {
    this.CursorLocation = 2;
    this.CursorType = 0;
    this.LockType = 1;
    this.CacheSize = 1;
    this.MaxRecords = 0;
    this.PageSize = 10;
    this.ActiveConnection = null;
    this.Source = "";
    this.__reset();
    this.__fieldsObj = com(new FieldsImpl(this));
  }
  __reset() {
    this.__open = false;
    this.__rows = [];
    this.__cols = [];
    this.__pos = 0;
    this.__table = null;
    this.__pending = null; // {mode:'add'|'edit', values:{}}
    this.__fieldCache = new Map();
  }
  __call(k) {
    return this.Fields.Item(k);
  }
  __set(k, v) {
    this.Fields.Item(k).Value = v;
  }

  get State() {
    return this.__open ? 1 : 0;
  }
  get Fields() {
    return this.__fieldsObj;
  }
  get EOF() {
    return !this.__pending?.mode?.startsWith("add") && this.__pos >= this.__rows.length;
  }
  get BOF() {
    return this.__rows.length === 0 || this.__pos < 0;
  }
  get RecordCount() {
    return this.__open ? this.__rows.length : -1;
  }
  get AbsolutePosition() {
    return this.__pos + 1;
  }
  set AbsolutePosition(v) {
    this.__commit();
    this.__pos = Number(v) - 1;
  }
  get PageCount() {
    return Math.ceil(this.__rows.length / (this.PageSize || 10));
  }
  get AbsolutePage() {
    return Math.floor(this.__pos / (this.PageSize || 10)) + 1;
  }
  set AbsolutePage(p) {
    this.__commit();
    this.__pos = (Number(p) - 1) * (this.PageSize || 10);
  }
  get Bookmark() {
    return this.__pos;
  }
  set Bookmark(v) {
    this.__pos = Number(v);
  }

  __colIndex(k) {
    k = prim(k);
    if (typeof k === "number") return k;
    if (/^\d+$/.test(String(k))) return Number(k);
    const l = String(k).toLowerCase();
    let i = this.__cols.findIndex((c) => c.name.toLowerCase() === l);
    if (i < 0) {
      // A.Campo / Tabla.Campo
      const short = l.includes(".") ? l.split(".").pop() : null;
      if (short) i = this.__cols.findIndex((c) => c.name.toLowerCase() === short);
    }
    return i;
  }
  __field(k) {
    const i = this.__colIndex(k);
    if (i < 0 || i >= this.__cols.length) {
      throw makeError(0x800a0cc1, `No se encontró el elemento en la colección correspondiente al nombre o al ordinal solicitado: ${prim(k)}`);
    }
    let f = this.__fieldCache.get(i);
    if (!f) {
      f = com(new FieldImpl(this, i));
      this.__fieldCache.set(i, f);
    }
    return f;
  }
  __current() {
    if (this.__pending && this.__pending.mode === "add") return this.__pending.row;
    if (this.__pos < 0 || this.__pos >= this.__rows.length) {
      throw makeError(0x800a0bcd, "BOF o EOF es True, o el registro actual se eliminó. La operación solicitada requiere un registro actual.");
    }
    return this.__rows[this.__pos];
  }
  __get(i) {
    const row = this.__current();
    const col = this.__cols[i];
    if (this.__pending && this.__pending.values && col.name in this.__pending.values) {
      return fromDb(this.__pending.values[col.name], col.type);
    }
    return fromDb(row[i], col.type);
  }
  __put(i, v) {
    if (!this.__table) {
      throw makeError(0x800a0cb3, "El recordset no admite actualizaciones (consulta de varias tablas). SQL: " + this.Source);
    }
    const col = this.__cols[i];
    if (!this.__pending) {
      this.__current();
      this.__pending = { mode: "edit", values: {} };
    }
    this.__pending.values[col.name] = toDb(v, col.type);
  }

  Open(source, conn, cursorType, lockType) {
    if (this.__open) this.Close();
    source = prim(source);
    if (conn !== undefined && conn !== null) this.ActiveConnection = conn;
    if (cursorType !== undefined) this.CursorType = cursorType;
    if (lockType !== undefined) this.LockType = lockType;
    let text = String(source).trim();
    if (/^[\wÀ-ÿ]+$/.test(text) || /^\[[^\]]+\]$/.test(text)) text = `SELECT * FROM ${text}`;
    this.Source = text;
    this.__execute(text);
  }
  Execute(source, conn) {
    this.Open(source, conn);
  }
  __execute(text) {
    const { t, stmt } = runSql(text);
    this.__reset();
    if (!stmt.reader) {
      stmt.run();
      return;
    }
    let sql = t.sql;
    let s = stmt;
    let ridIdx = -1;
    if (t.table && tableInfo(t.table)) {
      const q = t.alias ? `${t.alias}.rowid` : "rowid";
      const withRid = sql.replace(/^\s*select\s+/i, `SELECT ${q} AS __rid__, `);
      try {
        s = DB.prepare(withRid);
        ridIdx = 0;
        this.__table = tableInfo(t.table);
      } catch {
        s = stmt;
      }
    }
    s.raw(true);
    const meta = s.columns();
    const rows = s.all();
    const cols = meta.map((c) => {
      let type = (c.type || "").toUpperCase();
      if (!type && c.table) {
        const ti = tableInfo(c.table);
        const tc = ti && ti.byLower.get(String(c.column).toLowerCase());
        if (tc) type = tc.type;
      }
      // Con joins entre paréntesis SQLite desambigua como "ID:1"; ADO devolvía "ID".
      return { name: c.name.replace(/:\d+$/, ""), type, column: c.column, table: c.table };
    });
    if (ridIdx === 0) {
      this.__rids = rows.map((r) => r[0]);
      this.__rows = rows.map((r) => r.slice(1));
      this.__cols = cols.slice(1);
    } else {
      this.__rids = null;
      this.__rows = rows;
      this.__cols = cols;
    }
    this.__pos = 0;
    this.__open = true;
  }

  __commit() {
    if (this.__pending) this.Update();
  }
  MoveNext() {
    this.__commit();
    if (this.__pos < this.__rows.length) this.__pos++;
  }
  MovePrevious() {
    this.__commit();
    if (this.__pos >= 0) this.__pos--;
  }
  MoveFirst() {
    this.__commit();
    this.__pos = 0;
  }
  MoveLast() {
    this.__commit();
    this.__pos = Math.max(0, this.__rows.length - 1);
  }
  Move(n, start) {
    this.__commit();
    const base = start === 1 ? 0 : start === 2 ? this.__rows.length - 1 : this.__pos;
    this.__pos = Math.min(this.__rows.length, Math.max(-1, base + Number(n)));
  }

  AddNew(fields, values) {
    if (!this.__table) throw makeError(0x800a0cb3, "El recordset no admite AddNew. SQL: " + this.Source);
    this.__commit();
    this.__pending = { mode: "add", row: this.__cols.map(() => null), values: {} };
    if (fields !== undefined) {
      const fs = Array.isArray(fields) ? fields : [fields];
      const vs = Array.isArray(values) ? values : [values];
      fs.forEach((f, k) => this.__put(this.__colIndex(f), vs[k]));
    }
  }
  Update(fields, values) {
    if (fields !== undefined) {
      const fs = Array.isArray(fields) ? fields : [fields];
      const vs = Array.isArray(values) ? values : [values];
      fs.forEach((f, k) => this.__put(this.__colIndex(f), vs[k]));
    }
    const p = this.__pending;
    if (!p) return;
    this.__pending = null;
    const tbl = this.__table;
    const names = Object.keys(p.values);
    const q = (n) => `"${n.replace(/"/g, '""')}"`;
    if (p.mode === "add") {
      const sql = names.length
        ? `INSERT INTO ${q(tbl.name)} (${names.map(q).join(",")}) VALUES (${names.map(() => "?").join(",")})`
        : `INSERT INTO ${q(tbl.name)} DEFAULT VALUES`;
      const info = DB.prepare(sql).run(names.map((n) => p.values[n]));
      const rid = Number(info.lastInsertRowid);
      const fresh = DB.prepare(`SELECT * FROM ${q(tbl.name)} WHERE rowid = ?`).raw(true).get(rid);
      const freshCols = DB.prepare(`SELECT * FROM ${q(tbl.name)} LIMIT 0`).columns().map((c) => c.name.toLowerCase());
      const row = this.__cols.map((c) => {
        const k = freshCols.indexOf((c.column || c.name).toLowerCase());
        return k >= 0 ? fresh[k] : null;
      });
      this.__rows.push(row);
      if (this.__rids) this.__rids.push(rid);
      this.__pos = this.__rows.length - 1;
    } else if (names.length) {
      const rid = this.__rids[this.__pos];
      DB.prepare(`UPDATE ${q(tbl.name)} SET ${names.map((n) => `${q(n)} = ?`).join(", ")} WHERE rowid = ?`).run(
        ...names.map((n) => p.values[n]),
        rid
      );
      const row = this.__rows[this.__pos];
      names.forEach((n) => {
        const i = this.__cols.findIndex((c) => (c.column || c.name) === n);
        if (i >= 0) row[i] = p.values[n];
      });
    }
  }
  CancelUpdate() {
    this.__pending = null;
  }
  Delete() {
    if (!this.__table) throw makeError(0x800a0cb3, "El recordset no admite Delete. SQL: " + this.Source);
    const rid = this.__rids[this.__pos];
    DB.prepare(`DELETE FROM "${this.__table.name}" WHERE rowid = ?`).run(rid);
    this.__rows.splice(this.__pos, 1);
    this.__rids.splice(this.__pos, 1);
  }
  Requery() {
    this.__execute(this.Source);
  }
  Close() {
    if (this.__pending) {
      try {
        this.Update();
      } catch {
        /* ADO descartaba cambios inválidos al cerrar */
      }
    }
    this.__reset();
  }
  GetRows(n) {
    const out = [];
    const end = n > 0 ? Math.min(this.__rows.length, this.__pos + n) : this.__rows.length;
    for (; this.__pos < end; this.__pos++) out.push(this.__cols.map((_, i) => this.__get(i)));
    return out;
  }
  GetString() {
    return this.__rows.map((r) => r.join("\t")).join("\r");
  }
  Supports() {
    return true;
  }
  Clone() {
    const c = new RecordsetImpl();
    Object.assign(c, { __open: true, __rows: this.__rows, __rids: this.__rids, __cols: this.__cols, __table: this.__table, Source: this.Source });
    return com(c);
  }
}

// ------------------------------------------------------------------ Connection
class ErrorsImpl {
  constructor() {
    this.__list = [];
  }
  get Count() {
    return this.__list.length;
  }
  Item(i) {
    return this.__list[i];
  }
  __call(i) {
    return this.__list[i];
  }
  Clear() {
    this.__list.length = 0;
  }
  __items() {
    return this.__list;
  }
}

class ConnectionImpl {
  constructor() {
    this.ConnectionString = "";
    this.ConnectionTimeout = 15;
    this.CommandTimeout = 30;
    this.CursorLocation = 2;
    this.Mode = 0;
    this.Provider = "MSDASQL";
    this.__open = false;
    this.__errors = com(new ErrorsImpl());
  }
  get State() {
    return this.__open ? 1 : 0;
  }
  get Errors() {
    return this.__errors;
  }
  Open(cs) {
    if (cs !== undefined) this.ConnectionString = prim(cs);
    this.__open = true;
  }
  Close() {
    this.__open = false;
  }
  Execute(sql) {
    const rs = new RecordsetImpl();
    rs.ActiveConnection = this;
    rs.Source = prim(sql);
    rs.__execute(String(prim(sql)));
    return com(rs);
  }
  BeginTrans() {
    if (!DB.inTransaction) DB.exec("BEGIN");
    return 1;
  }
  CommitTrans() {
    if (DB.inTransaction) DB.exec("COMMIT");
  }
  RollbackTrans() {
    if (DB.inTransaction) DB.exec("ROLLBACK");
  }
}

// ------------------------------------------------------------------ Command
class ParametersImpl {
  constructor() {
    this.__list = [];
  }
  get Count() {
    return this.__list.length;
  }
  Append(p) {
    this.__list.push(p);
  }
  Item(i) {
    return typeof i === "number" ? this.__list[i] : this.__list.find((p) => p.Name === i);
  }
  __call(i) {
    return this.Item(i);
  }
  Refresh() {}
}

class CommandImpl {
  constructor() {
    this.ActiveConnection = null;
    this.CommandText = "";
    this.CommandType = 1;
    this.CommandTimeout = 30;
    this.Prepared = false;
    this.__params = com(new ParametersImpl());
  }
  get Parameters() {
    return this.__params;
  }
  CreateParameter(Name, Type, Direction, Size, Value) {
    return { Name, Type, Direction, Size, Value };
  }
  Execute() {
    let text = String(prim(this.CommandText));
    const params = implOf(this.__params).__list;
    if (params.length) {
      let k = 0;
      text = text.replace(/\?/g, () => {
        const v = bindParam(params[k++]?.Value);
        return v == null ? "NULL" : typeof v === "number" ? String(v) : `'${String(v).replace(/'/g, "''")}'`;
      });
    }
    const rs = new RecordsetImpl();
    rs.Source = text;
    rs.__execute(text);
    return com(rs);
  }
}

export const createConnection = () => com(new ConnectionImpl());
export const createRecordset = () => com(new RecordsetImpl());
export const createCommand = () => com(new CommandImpl());

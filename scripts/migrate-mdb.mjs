// Convierte database/Sepad.mdb (Access 97/2000) en una base SQLite.
// Uso: node scripts/migrate-mdb.mjs [origen.mdb] [destino.sqlite]
import fs from "node:fs";
import path from "node:path";
import MDBReader from "mdb-reader";
import Database from "better-sqlite3";

const root = path.resolve(path.dirname(new URL(import.meta.url).pathname), "..");
const src = process.argv[2] || path.join(root, "database/Sepad.mdb");
const dst = process.argv[3] || path.join(root, "data/sepad.sqlite");

const TYPE_MAP = {
  boolean: "BOOLEAN",
  byte: "INTEGER",
  integer: "INTEGER",
  long: "INTEGER",
  currency: "REAL",
  float: "REAL",
  double: "REAL",
  datetime: "DATETIME",
  binary: "BLOB",
  text: "TEXT",
  ole: "BLOB",
  memo: "TEXT",
  repid: "TEXT",
  numeric: "REAL",
  complex: "INTEGER",
  bigint: "INTEGER",
};

const CREATED_AT = /^(fecha|date|fechahora|fecha_publicada|uploaddate|fechaingreso)$/i;
const NOW_DEFAULT = "(strftime('%Y-%m-%d %H:%M:%S', 'now', 'localtime'))";
const q = (id) => `"${id.replace(/"/g, '""')}"`;
const pad = (n) => String(n).padStart(2, "0");
export const toSqlDate = (d) =>
  `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;

const reader = new MDBReader(fs.readFileSync(src));
fs.mkdirSync(path.dirname(dst), { recursive: true });
for (const f of [dst, dst + "-wal", dst + "-shm"]) if (fs.existsSync(f)) fs.unlinkSync(f);
const db = new Database(dst);
db.pragma("journal_mode = WAL");

let total = 0;
db.transaction(() => {
  for (const name of reader.getTableNames()) {
    // Las tablas ~TMP son restos temporales de Access (contienen datos personales antiguos).
    if (name.startsWith("~")) continue;
    const table = reader.getTable(name);
    const cols = table.getColumns();
    const defs = cols.map((c) => {
      if (c.autoLong) return `${q(c.name)} INTEGER PRIMARY KEY AUTOINCREMENT`;
      const type = TYPE_MAP[c.type] || "TEXT";
      // En SQL Server estas columnas tenían DEFAULT getdate(): varias páginas no las asignan.
      if (type === "DATETIME" && CREATED_AT.test(c.name)) return `${q(c.name)} ${type} DEFAULT ${NOW_DEFAULT}`;
      return `${q(c.name)} ${type}`;
    });
    db.exec(`CREATE TABLE ${q(name)} (${defs.join(", ")})`);
    const rows = table.getData();
    if (rows.length) {
      const ins = db.prepare(
        `INSERT INTO ${q(name)} (${cols.map((c) => q(c.name)).join(",")}) VALUES (${cols.map(() => "?").join(",")})`
      );
      for (const row of rows) {
        ins.run(
          cols.map((c) => {
            const v = row[c.name];
            if (v == null) return null;
            if (v instanceof Date) return toSqlDate(v);
            if (typeof v === "boolean") return v ? 1 : 0;
            if (Buffer.isBuffer(v)) return v;
            return v;
          })
        );
      }
    }
    total += rows.length;
    console.log(`  ${name}: ${rows.length} filas`);
  }
})();
// Tablas de la etapa SQL Server que no estaban en el .mdb
db.exec(fs.readFileSync(path.join(root, "scripts/schema-extra.sql"), "utf8"));
db.close();
console.log(`Migración completa: ${total} filas -> ${path.relative(root, dst)}`);

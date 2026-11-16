# SEPAD — Sistema de Enseñanza Personalizada a Distancia

Legacy LMS (distance-learning system) built by the repo owner in 1998 for the Cuban context, using classic **Active Server Pages (ASP)** with server-side **JScript**, well before .NET existed. It later moved from Access to SQL Server (the code uses that dialect: `CONVERT`, `DATEDIFF(minute, …)`, nested joins).

## Goal: modernize while keeping the original interface

The original pages, frames, CSS, GIFs and texts are the product; people revisit it for nostalgia. Modernization happens **underneath**, in a Node.js runtime that executes the original `.asp` files almost unchanged.

- Don't redesign or rewrite the UI. Fix original pages only for genuine bugs (they were broken in IIS too) and keep the diff minimal.
- Pages written in VBScript are ported to JScript in `server/ports/<same path>`; the runtime prefers the port and the original stays as a historical artifact.
- Files are ISO-8859-1. Edit them byte-safely (e.g. `LC_ALL=C sed`), never re-save as UTF-8.

## Layout

- `default.asp`, `global.asa`, `pages_esp/`, `Js/`, `Css/`, `images/`, `Help/`, `XSL/`, `XML/` — the original 1998–2004 application.
- `server/index.js` — HTTP server: routes `.asp` to the runtime, serves static files case-insensitively, blocks `.inc/.asa/.mdb`, `XML/config.xml` and runtime code.
- `server/asp/` — classic ASP runtime: `compiler.js` (`<% %>`, `#include`, `runat=server`, mini VBScript), `jsrewrite.js` (`Session("x") = v` → `__cs(...)`), `runtime.js` (Request/Response/Server/Session/Application, IIS-style error page), `ado.js` (ADODB over better-sqlite3), `sql.js` (Access/SQL Server → SQLite), `objects.js` (FileSystemObject, MSXML + WD-xsl → XSLT via `xsltproc`, aspSmartUpload, simulated mail), `com.js` (case-insensitive, callable COM-like proxies).
- `server/compat/` — client-side IE5/6 compatibility: `ie-compat.js` (injected shim) and `rewrite.js` (on-the-fly rewrites like `frames("x")`, `function document.onclick()`, CSS `10 px`).
- `server/ports/` — JScript ports of the VBScript pages (downloads/uploads).
- `scripts/` — `migrate-mdb.mjs` (+ `schema-extra.sql` for SQL Server-era tables), `seed-demo.mjs`, and test tools (`check-compile`, `crawl`, `tour`, `popups`, `shot`, `check-client-js`).
- `demo/Courses/` — demo course content; copied to `Courses/` (gitignored) on setup/reset.

## Commands

- `npm install && npm run setup` — build `data/sepad.sqlite` from `database/Sepad.mdb` plus demo data (`data/sepad.seed.sqlite` is the pristine copy).
- `npm start` — serve on `PORT` (default 8080). Env: `SEPAD_RESET_ON_START=1`, `SEPAD_RESET_HOURS=n`, `SEPAD_DEBUG=1` (stack traces), `SEPAD_LOG=0`.
- `npm run check` — compile every `.asp`; must report 0 errors.
- `node scripts/crawl.mjs http://localhost:8080/ 600 alumno alumno` — crawl as a role; `scripts/tour.mjs` and `scripts/popups.mjs` drive Chrome via puppeteer-core.
- Demo users: `administrador/admin`, `profesor/profesor`, `alumno/alumno`.

## Gotchas

- The shell is zsh: `$VAR` with spaces isn't word-split, and the `grep` alias is ugrep, which misses matches in Latin-1 files. Use `/usr/bin/grep` with `LC_ALL=C`, or a bash script.
- The original builds SQL by string concatenation (SQL injection everywhere); a public demo must reset its data periodically.
- Crawlers must skip side-effect pages (`Confirm*`, `Acepta*`, `proceso*`…); following them over GET corrupts config and data, just as it would have in IIS.

# Base de datos

## De Access a SQLite

`database/Sepad.mdb` es un respaldo de la base Access. `scripts/migrate-mdb.mjs` (`npm run migrate`) lo lee con mdb-reader y crea `data/sepad.sqlite`:

- Conserva nombres de tablas y columnas; los autonuméricos pasan a `INTEGER PRIMARY KEY AUTOINCREMENT`.
- Tipos: Sí/No → `BOOLEAN` (0/1), Fecha/Hora → `DATETIME` (texto `YYYY-MM-DD HH:MM:SS`), números → `INTEGER`/`REAL`, Texto/Memo → `TEXT`.
- Las columnas de fecha de creación (`Fecha`, `Date`, `fechahora`, `Fecha_Publicada`, `uploadDate`, `fechaIngreso`) reciben la hora actual por defecto. En SQL Server tenían `DEFAULT getdate()` y varias páginas no las asignan.
- Omite las tablas temporales `~TMP…` de Access, que contenían datos personales antiguos.

El respaldo es anterior al paso a SQL Server y le faltan tablas que el código usa. `scripts/schema-extra.sql` las crea:

| Tabla | Origen del esquema |
| --- | --- |
| `AgendaEventos` | `legacy/SEPAD_06_01_09_Agenda/.../Crea_Tabla_AgendaEventos.sql` (DDL original) |
| `Mensajes_Enviados`, `Glosarios`, `faq`, `HistoriaLecciones`, `Modulo_Curso` | Deducido de las páginas que las usan |

## Traducción del SQL (`server/asp/sql.js`)

El código escribe SQL de Access y, sobre todo, de SQL Server (`config.xml` tiene `SQL=1`, de modo que las fechas van entre `'...'` y los booleanos son `1`/`0`). Antes de ejecutarse, cada consulta se traduce así:

| Original | SQLite |
| --- | --- |
| `SELECT TOP 10 ...` | `SELECT ... LIMIT 10` |
| `[Columna]` | `"Columna"` |
| `"texto"` (cadena) | `'texto'` |
| `N'texto'` | `'texto'` |
| `#11/16/2026#`, `'11/16/2026'` en comparaciones | `'2026-11-16 00:00:00'` |
| `DELETE * FROM t` | `DELETE FROM t` |
| `a & b` (concatenación de Access) | `a \|\| b` |
| `DATEDIFF(minute, a, b)` | `DATEDIFF('n', a, b)` |
| `CONVERT(char(8), fecha, 112)` | `sqlconvert('char', 8, fecha, 112)` |
| `A JOIN B JOIN C ON bc ON ab` | `A JOIN (B JOIN C ON bc) ON ab` |
| `x UNION (SELECT ...)` | `x UNION SELECT ...` |

Funciones de Access y SQL Server registradas en SQLite: `Now`, `GetDate`, `Date`, `Time`, `Year`, `Month`, `Day`, `Hour`, `Minute`, `Weekday`, `DateDiff`, `DateAdd`, `DatePart`, `DateValue`, `CDate`, `LCase`, `UCase`, `Len`, `Mid`, `Left`, `Right`, `IIf`, `Nz`, `IsNull`, `CStr`, `CInt`, `CLng`, `Val`, `Format` y `sqlconvert` (estilos 101, 102, 103, 104, 105, 108, 112, 120, 121 y 23).

Si SQLite desambigua un nombre de columna como `ID:1` (ocurre con joins entre paréntesis), el `Recordset` lo expone como `ID`, igual que ADO.

## Datos de demostración

La base original solo trae el esquema, dos usuarios (`guest` y `administrador`), los grupos predefinidos y la configuración. `scripts/seed-demo.mjs` (`npm run seed`) agrega:

- usuarios `profesor` (ID 2000) y `alumno` (ID 2001), y contraseña `admin` para el administrador (en el original estaba vacía);
- grupos 100 (alumnos) y 101 (claustro) del *Diplomado en Tecnologías de Internet* (modalidad 1);
- módulos 1 (*Introducción a Internet*, 3 lecciones) y 2 (*Diseño de páginas Web con HTML*, 2 lecciones), con su contenido en `demo/Courses/`;
- noticias, un anuncio, temas y mensajes de foro, un mensaje interno, un evento de agenda, glosario, preguntas frecuentes y una sala de chat.

Las fechas se calculan respecto al día de la siembra, para que aparezcan como recientes.

`npm run setup` ejecuta la migración y la siembra, copia `demo/Courses/` a `Courses/` y guarda una copia intacta en `data/sepad.seed.sqlite`.

## Reinicio de la demostración

`server/index.js` puede restaurar `data/sepad.sqlite` desde `data/sepad.seed.sqlite` y `Courses/` desde `demo/Courses/`:

- al arrancar, con `SEPAD_RESET_ON_START=1` (o si no existe la base);
- cada *n* horas, con `SEPAD_RESET_HOURS=n`.

Para empezar desde cero (volver a migrar y sembrar): `make reset`.

## Inspeccionar los datos

```sh
sqlite3 data/sepad.sqlite ".tables"
sqlite3 data/sepad.sqlite "select ID, Name, PermissionType from Usuarios"
```

Las contraseñas se guardan como MD5 en hexadecimal mayúscula, que el navegador calcula con `Js/md5.js` antes de enviar el formulario.

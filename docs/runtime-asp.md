# Runtime de ASP clásico

Código en `server/asp/`.

## Compilación de páginas (`compiler.js`)

1. **Includes**: expande `<!-- #include file="..." -->` (relativo al archivo) y `virtual` (relativo a la raíz), recursivamente y sin distinguir mayúsculas. Un include que no existe produce el error de ASP `0126`.
2. **Segmentos**: separa el texto literal, `<% código %>`, `<%= expresión %>`, la directiva `<%@ Language=... %>` y los bloques `<script runat=server>`.
3. **Generación**: el texto literal se emite como `__w(__L[n])`, donde `__L` es la tabla de literales de la página; las expresiones como `__w(expr)` y el código tal cual. Se omiten los bloques que solo tienen espacios entre dos bloques de código, para que sea válido un `switch {` seguido de `%> <%` y de `case`, algo que JScript aceptaba.
4. **Reescritura** (`jsrewrite.js`): JScript permitía asignar a una llamada, por ejemplo `Session("x") = valor` u `oRec.Fields.Item("c") = v`. Un tokenizador detecta `callee(args) = expr` (también `+=` y `-=`) y lo convierte en `__cs(callee, [args], expr)`, conservando los saltos de línea para no alterar la numeración.
5. **Mapa de líneas**: cada línea generada recuerda su archivo y línea de origen, para que los errores apunten al `.asp` o `.inc` correcto.

**VBScript**: los bloques `<script language=vbscript runat=server>` de SEPAD solo definen funciones del tipo `getDate = Now()`, y se traducen a JScript. Una página completa en VBScript, ya sea por su directiva o porque IIS asumía VBScript cuando no la había, muestra un aviso, salvo que tenga un port en `server/ports/`.

`npm run check` compila todas las páginas para detectar errores de sintaxis.

## Objetos de ASP (`runtime.js`)

| Objeto | Detalles |
| --- | --- |
| `Request` | `QueryString`, `Form`, `Cookies`, `ServerVariables`, `TotalBytes`, `BinaryRead` y `Request("x")`. Cada valor se comporta como en IIS: `+ ""` da `"undefined"` si falta, `.Count` cuenta valores e `.Item(i)` es base 1. Decodifica en Latin-1 y acepta UTF-8 en URLs escritas a mano |
| `Response` | `Write`, `Redirect`, `End`, `Clear`, `Buffer`, `Expires`, `ContentType`, `Charset`, `Status`, `AddHeader`, `Cookies`, `BinaryWrite`. `Write(true)` escribe `True` y las fechas de ADO salen como en Windows (`11/16/2026 7:35:00 AM`) |
| `Server` | `CreateObject`, `MapPath` (sin distinguir mayúsculas), `URLEncode` y `HTMLEncode` (Latin-1), `Execute` y `Transfer` |
| `Session` | En memoria, con claves sin distinguir mayúsculas, `Timeout` (20 minutos), `SessionID`, `Abandon` y `Contents`. Guarda objetos de JavaScript tal cual, ya que el código de SEPAD guarda objetos en la sesión |
| `Application` | Igual que `Session`, pero global. `Lock` y `Unlock` no hacen nada porque la ejecución es secuencial |

`global.asa` se ejecuta una vez al arrancar (`Application_OnStart`) y al crear cada sesión (`Session_OnStart`). Su texto literal no se envía al cliente.

Cada página corre en un contexto `vm` nuevo que además incluye `Enumerator`, `VBArray`, `ActiveXObject`, `ScriptEngine*` y `Date.prototype.getYear` con la semántica de JScript.

## Proxies COM (`com.js`)

JScript accedía a los objetos COM por IDispatch: los nombres no distinguían mayúsculas (`Response.write`, `oRec.movenext`, `.value`) y algunos objetos se podían invocar (`oRec("campo")`, `Session("x")`). `com(impl)` envuelve una clase en un `Proxy` que:

- resuelve cualquier miembro sin distinguir mayúsculas;
- hace invocable el objeto si la clase define `__call`;
- permite asignar a una llamada si define `__set` o `__set_<Método>`;
- convierte a primitivo vía `__value()`, para que un `Field` o un valor de `Request` funcionen en concatenaciones y comparaciones.

## ADODB (`ado.js`)

| Objeto | Soporta |
| --- | --- |
| `Connection` | `Open` (ignora la cadena: siempre es la base SQLite), `Execute`, `Close`, transacciones, `Errors` |
| `Recordset` | `Open`, `EOF`/`BOF`, `Move*`, `RecordCount`, `Fields`, `AddNew`, `Update`, `Delete`, `CancelUpdate`, `Requery`, `GetRows`, paginación, `Clone` |
| `Command` | `CommandText`, `ActiveConnection`, `Execute`, parámetros `?` |
| `Field` | `Value`, `Name`, `Type`, `ActualSize`; se lee siempre del registro actual |

Para que un `Recordset` sea editable, la consulta debe ser un `SELECT` sobre una sola tabla. En ese caso se le añade `rowid` y `Update` escribe con `UPDATE ... WHERE rowid = ?`. Como en ADO, cambiar de registro o cerrar con cambios pendientes los guarda. `AddNew` y `Update` insertan y releen la fila, lo que devuelve el `ID` autonumérico y los valores por defecto.

Conversión de tipos según la columna: `BOOLEAN` pasa a `true`/`false`; `DATETIME` a una fecha `VtDate`, que se escribe como en Windows; los textos numéricos a número. Al escribir se aceptan fechas de JavaScript, textos `m/d/yyyy`, `"True"`, `-1` y similares.

La traducción del SQL está en [Base de datos](base-de-datos.md).

## Componentes COM (`objects.js`)

| ProgID | Implementación |
| --- | --- |
| `Scripting.FileSystemObject` | Archivos y carpetas dentro del sitio (no puede salir de él): `FileExists`, `GetFolder().Files`, `OpenTextFile`, `CreateFolder`, etc. |
| `Microsoft.XMLDOM` (y `MSXML2.DOMDocument`) | `load`, `loadXML`, `save`, `selectSingleNode`, `selectNodes`, `.text`, `.xml` y `transformNode` sobre @xmldom/xmldom y xpath |
| `aspSmartUpload.SmartUpload` | Parser multipart propio: `Upload`, `Files.Item(...).SaveAs`, `Form`, listas de extensiones permitidas y prohibidas, tamaño máximo |
| `RWMS.RWMailSender`, `CDONTS.NewMail` | Aceptan cualquier propiedad y método y solo registran el envío en la consola |
| `Scripting.Dictionary` | En `com.js` |

**XSL de IE5**: `XSL/prologo.xsl` usa el espacio de nombres `http://www.w3.org/TR/WD-xsl` de 1998, anterior a XSLT. `wdxslToXslt` lo convierte en XSLT 1.0: quita las reglas implícitas, convierte las plantillas anidadas dentro de `apply-templates` en modos y completa los `value-of` sin `select`. Después ejecuta `xsltproc` con salida ISO-8859-1.

**Rutas de Windows**: código como `load('D:\\sepad\\webapp\\XML\\config.xml')` se resuelve buscando el sufijo más largo que exista dentro del sitio (`XML/config.xml`).

## Errores

Los errores se muestran como lo hacía IIS:

```
Microsoft JScript runtime error '800a138f'
'Fields' es nulo o no es un objeto
/pages_esp/Course.asp, line 23
```

Los errores de SQL aparecen como `Microsoft OLE DB Provider for ODBC Drivers error '80040e14'` y, en la consola del servidor, incluyen el SQL traducido. Con `SEPAD_DEBUG=1` también se imprime la traza de Node.

## Limitaciones conocidas

- No hay intérprete general de VBScript; las páginas VBScript necesitan un port.
- Los `Recordset` de varias tablas son de solo lectura.
- Las sesiones viven en memoria: se pierden al reiniciar y no se comparten entre varias instancias.
- El correo no se envía.

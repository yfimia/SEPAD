# Desarrollo y pruebas

## Convenciones

- **La interfaz no se rediseña.** Las páginas originales solo se tocan para corregir errores que también las rompían en IIS, con el cambio mínimo. Todo lo demás va en `server/`.
- **Codificación ISO-8859-1.** Los `.asp`, `.inc`, `.htm`, `.js` y `.css` originales están en Latin-1. Edítalos sin convertir la codificación (por ejemplo con `LC_ALL=C sed`) y nunca los guardes en UTF-8: los acentos se romperían en la página.
- **Ports de VBScript.** Una página VBScript se reescribe en JScript en `server/ports/<misma ruta>` y el original se mueve a `legacy/<misma ruta>`. El servidor atiende la URL original con el port.
- **`legacy/`** guarda el código original que ya no se usa. No se publica ni se compila.

## Herramientas de prueba

Con el sitio en marcha (`make up PORT=8082`):

| Herramienta | Uso |
| --- | --- |
| `make check` | Compila todas las páginas `.asp`; debe reportar 0 errores |
| `node scripts/crawl.mjs URL 600 usuario clave` | Rastrea el sitio con sesión siguiendo frames y enlaces, y agrupa los errores ASP. Omite las páginas con efectos (`Confirm*`, `Acepta*`, `proceso*`, `public*`…) y no sigue el `action` de los formularios |
| `make crawl PORT=8082` | El rastreo anterior con los tres roles |
| `node scripts/tour.mjs URL dir usuario clave` | Recorrido en Chrome: login, modalidad, módulo y Aula Virtual, con capturas y errores de JavaScript |
| `node scripts/popups.mjs URL dir usuario clave` | Abre cada ventana emergente de los menús y del Aula Virtual, con capturas y errores |
| `node scripts/shot.mjs URL archivo.png [usuario clave]` | Captura de una página |
| `node scripts/check-client-js.mjs URL usuario clave ruta...` | Compila los `<script>` en línea y los manejadores `on*` de páginas concretas |

Las herramientas de Chrome usan puppeteer-core con el Chrome instalado (`/Applications/Google Chrome.app`, o la ruta en la variable `CHROME`).

Para ver por qué falla una página, revisa la consola del servidor (`make logs` o `make debug`): cada error ASP incluye el SQL traducido cuando lo hay.

**Cuidado al probar.** Muchas páginas de SEPAD escriben en la base cuando se abren por GET; por ejemplo, `publicNews.asp` publica una noticia vacía. Después de probar, restaura los datos con `SEPAD_RESET_ON_START=1 make restart PORT=8082`, o desde cero con `make reset`.

## Tareas habituales

### Corregir un error en una página

1. Reproducirlo y leer la página de error: indica el archivo y la línea originales.
2. Si el fallo es del runtime (traducción SQL, objeto COM, compatibilidad), corregirlo en `server/`.
3. Si el fallo está en la página (también fallaba en IIS), corregirla con el cambio mínimo y en Latin-1.
4. Ejecutar `make check` y el rastreo o recorrido correspondiente.

### Portar una página VBScript

1. Crear `server/ports/<ruta>.asp` en JScript (`<%@ Language=JScript %>`). Los includes se resuelven relativos al port; los compartidos están en `server/ports/inc/`. `Server.MapPath` sigue siendo relativo a la URL original.
2. Mover el original a `legacy/<ruta>` con `git mv`.
3. Comprobar que la URL original responde.

### Agregar soporte para un objeto COM

Registrar el ProgID en `PROGIDS` (`server/asp/runtime.js`) e implementarlo como clase envuelta con `com()` (ver [Runtime de ASP](runtime-asp.md#proxies-com-comjs)).

### Agregar una función SQL

Registrarla en `registerFunctions` (`server/asp/sql.js`). Si necesita otra sintaxis, añadir el caso en `translate`.

## Correcciones hechas a las páginas originales

Estas páginas tenían errores de sintaxis que las rompían también en IIS:

| Archivo | Error |
| --- | --- |
| `AdminForumTopics.asp`, `AdminForumMessages.asp` (y sus copias en `ModuloTutor/pages_esp/`) | `<%=...%>` dentro de un string de código de servidor; `<%=enviado><%=` y `<%%=siacept>` |
| `inc/AdminForumMessages.inc` | `<%` sin cerrar |
| `lista_de_anotaciones.asp` | `Fields.Item.("texto")` |
| `news.asp` | `" and where (..."` al paginar |
| `CourseManager.asp` | Faltaba incluir `inc/coursemanager.inc` (`stmp1` no definido) |

`git diff master..modernizacion -- pages_esp` muestra los cambios exactos.

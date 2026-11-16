# Compatibilidad con el navegador

SEPAD se escribió para Internet Explorer 5 y 6. Para que su JavaScript y su CSS funcionen en navegadores actuales sin modificar los archivos, el servidor actúa en dos frentes (código en `server/compat/`).

## Reescrituras en el servidor (`rewrite.js`)

Se aplican al HTML generado por las páginas ASP y a los `.htm`, `.js` y `.css` estáticos. Quedan excluidos TinyMCE y el editor wysiwyg, que ya funcionaban en varios navegadores.

| Escrito para IE | Se sirve como | Motivo |
| --- | --- | --- |
| `parent.frames("mainFrame")` | `parent.__frames("mainFrame")` | `window.frames` no se puede invocar como función |
| `x.children("id")`, `.options(0)`, `.elements("n")`, `.rows(i)`… | `x.children.__c("id")` | Las colecciones del DOM no se pueden invocar |
| `document.all.tags("TD")` | `document.all.__tags("TD")` | `tags()` no existe |
| `function document.onclick() { ... }` | `document.onclick = function () { ... }` | Sintaxis propia de JScript para declarar manejadores |
| `function window::onload() { ... }` | `window.onload = function () { ... }` | Ídem |
| `FONT-SIZE: 10 px` (en CSS, `<style>` y `style=`) | `FONT-SIZE: 10px` | IE aceptaba el espacio; los navegadores actuales descartan la regla y el texto se ve grande |

Además se inyecta `<script src="/__sepad/ie-compat.js">` al principio de cada documento HTML.

## Shim del cliente (`ie-compat.js`)

- `__frames`, `__c` y `__tags`, los destinos de las reescrituras.
- `element.all` para elementos que no son `document`.
- `Date.prototype.getYear` con la semántica de IE: devuelve el año completo desde el 2000 (2026, no 126).
- `attachEvent` y `detachEvent`, `currentStyle`, y `document.createElement("<input name='x'>")`.
- `showModalDialog`, que los navegadores eliminaron y que SEPAD usaba con su valor de retorno síncrono (`user = showModalDialog("findTutor.asp")`). El shim abre el diálogo en una ventana y aborta el manejador. Cuando el diálogo se cierra con `window.returnValue`, guarda el resultado y repite el clic original; esta vez `showModalDialog` devuelve el valor de inmediato. Dentro del diálogo también proporciona `window.dialogArguments`.

## Lo que ya funcionaba

Muchas características de IE siguen existiendo en los navegadores actuales y no requieren nada: `window.event`, `event.srcElement`, `document.all` (incluso invocable), `innerText`, `insertAdjacentHTML`, el acceso a elementos por `id` como variables globales, `<marquee>` y los framesets. Los `<script language="jscript">` se ejecutan como JavaScript.

Los `<script language="vbscript">` del cliente se ignoran (son casi siempre cuadros de mensaje). Las pocas comprobaciones de navegador del código (`Js/Browser.js`) toman el camino "no IE", que usa el DOM estándar.

## Agregar una compatibilidad

1. Reproducir el error con `node scripts/tour.mjs` o `node scripts/popups.mjs`, que muestran los errores de JavaScript de todos los frames y ventanas, o con `node scripts/check-client-js.mjs`, que compila los scripts en línea de páginas concretas.
2. Si es sintaxis que el navegador no puede analizar, agregar una reescritura en `rewrite.js`. Si es una API que falta, agregarla en `ie-compat.js`.
3. No modificar los archivos originales para esto.

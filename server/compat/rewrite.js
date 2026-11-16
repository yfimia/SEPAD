// Reescrituras al vuelo para que el HTML/JS/CSS de 1998 funcione en navegadores modernos.
// Los archivos originales no se modifican.

export const COMPAT_URL = "/__sepad/ie-compat.js";

const COLLECTIONS = /\.(children|childNodes|options|elements|forms|images|links|anchors|rows|cells)\s*\(/g;

/** JavaScript: llamadas de estilo IE a colecciones y frames. */
export function rewriteJs(src) {
  return src
    // JScript: "function document.onclick()" y "function window::onload()"
    .replace(/\bfunction\s+([A-Za-z_$][\w$]*)\s*::\s*([A-Za-z_$][\w$]*)\s*\(/g, "$1.$2 = function (")
    .replace(/\bfunction\s+([A-Za-z_$][\w$]*(?:\.[A-Za-z_$][\w$]*)+)\s*\(/g, "$1 = function (")
    .replace(/\.frames\s*\(/g, ".__frames(")
    .replace(COLLECTIONS, ".$1.__c(")
    .replace(/\.tags\s*\(/g, ".__tags(");
}

/** CSS: IE aceptaba "10 px" (con espacio); los navegadores modernos lo descartan. */
export function rewriteCss(src) {
  return src.replace(/(\d)\s+(px|pt|em|ex|%)\b/gi, "$1$2");
}

/** HTML: JS en línea, estilos y la inyección del shim. */
export function rewriteHtml(html) {
  let out = rewriteJs(html)
    .replace(/(<style\b[^>]*>)([\s\S]*?)(<\/style>)/gi, (_, a, css, b) => a + rewriteCss(css) + b)
    .replace(/\bstyle\s*=\s*(["'])(.*?)\1/gi, (_, q, css) => `style=${q}${rewriteCss(css)}${q}`);
  const tag = `<script src="${COMPAT_URL}"></script>`;
  if (/<head\b[^>]*>/i.test(out)) return out.replace(/<head\b[^>]*>/i, (m) => m + tag);
  if (/<html\b[^>]*>/i.test(out)) return out.replace(/<html\b[^>]*>/i, (m) => m + tag);
  if (/<(body|frameset|table|form|script|p|div)\b/i.test(out)) return tag + out;
  return out;
}

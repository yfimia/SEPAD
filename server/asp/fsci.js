// Resolución de rutas sin distinguir mayúsculas (IIS/Windows lo hacía; Linux no).
import fs from "node:fs";
import path from "node:path";

const dirCache = new Map();

function listDir(dir) {
  let m = dirCache.get(dir);
  if (!m) {
    m = new Map();
    try {
      for (const e of fs.readdirSync(dir)) m.set(e.toLowerCase(), e);
    } catch {
      /* no existe */
    }
    dirCache.set(dir, m);
  }
  return m;
}

export function clearDirCache() {
  dirCache.clear();
}

/** Devuelve la ruta real (con las mayúsculas correctas) o null si no existe. */
export function resolveCI(absPath) {
  if (fs.existsSync(absPath)) return absPath;
  const parts = path.resolve(absPath).split(path.sep);
  let cur = parts[0] === "" ? path.sep : parts[0];
  for (let i = 1; i < parts.length; i++) {
    const p = parts[i];
    if (!p) continue;
    const exact = path.join(cur, p);
    if (fs.existsSync(exact)) {
      cur = exact;
      continue;
    }
    const real = listDir(cur).get(p.toLowerCase());
    if (!real) return null;
    cur = path.join(cur, real);
  }
  return cur;
}

let siteRoot = null;
export function setSiteRoot(root) {
  siteRoot = root;
}

/**
 * Rutas absolutas de Windows escritas a mano en el código (p.ej. 'D:\\sepad\\webapp\\XML\\config.xml'):
 * se busca el sufijo más largo que exista dentro del sitio.
 */
export function mapWindowsPath(p) {
  const s = String(p);
  if (!siteRoot || !/^[A-Za-z]:[\\/]/.test(s)) return s;
  const parts = s.slice(3).split(/[\\/]+/).filter(Boolean);
  for (let i = 0; i < parts.length; i++) {
    const cand = resolveCI(path.join(siteRoot, ...parts.slice(i)));
    if (cand) return cand;
  }
  return path.join(siteRoot, ...parts.slice(-1));
}

/** Como resolveCI, pero si faltan los últimos tramos los añade tal cual al prefijo real. */
export function resolveCIPartial(absPath) {
  const full = resolveCI(absPath);
  if (full) return full;
  const parts = path.resolve(absPath).split(path.sep);
  for (let i = parts.length - 1; i > 0; i--) {
    const prefix = resolveCI(parts.slice(0, i).join(path.sep) || path.sep);
    if (prefix) return path.join(prefix, ...parts.slice(i));
  }
  return absPath;
}

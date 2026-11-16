// Envoltorio "tipo COM": acceso a miembros sin distinguir mayúsculas
// (JScript + IDispatch lo permitía: Response.write, oRec.movenext, .value ...),
// objetos invocables (oRec("campo"), Session("x")) y asignación a llamadas
// (Session("x") = v  ->  __cs(Session, ["x"], v)).

const IMPL = Symbol("impl");

function keyMap(impl) {
  if (impl.__keys) return impl.__keys;
  const map = new Map();
  let o = impl;
  while (o && o !== Object.prototype) {
    for (const k of Object.getOwnPropertyNames(o)) {
      const l = k.toLowerCase();
      if (!map.has(l)) map.set(l, k);
    }
    o = Object.getPrototypeOf(o);
  }
  Object.defineProperty(impl, "__keys", { value: map, enumerable: false });
  return map;
}

function resolveKey(impl, prop) {
  if (typeof prop !== "string") return prop;
  if (prop in impl) return prop;
  const k = keyMap(impl).get(prop.toLowerCase());
  return k === undefined ? prop : k;
}

/**
 * Crea un proxy para `impl`. Si `impl.__call` existe, el objeto es invocable.
 * Los métodos que admiten asignación (p.ej. Item) se declaran como
 * `__set_<Nombre>`; la invocación por defecto usa `__set`.
 */
export function com(impl) {
  const target = typeof impl.__call === "function" ? function () {} : {};
  const proxy = new Proxy(target, {
    get(_t, prop) {
      if (prop === IMPL) return impl;
      if (prop === Symbol.toPrimitive) {
        if (impl[Symbol.toPrimitive]) return impl[Symbol.toPrimitive].bind(impl);
        if (impl.valueOf !== Object.prototype.valueOf) return () => impl.valueOf();
        return undefined;
      }
      if (prop === "__set") return impl.__set ? impl.__set.bind(impl) : undefined;
      const key = resolveKey(impl, prop);
      const v = impl[key];
      if (typeof v === "function" && !v[IMPL] && !(typeof key === "string" && key.startsWith("__"))) {
        const bound = v.bind(impl);
        const setter = impl["__set_" + key];
        if (setter) bound.__set = setter.bind(impl);
        return bound;
      }
      return v;
    },
    set(_t, prop, value) {
      impl[resolveKey(impl, prop)] = value;
      return true;
    },
    has(_t, prop) {
      return resolveKey(impl, prop) in impl;
    },
    apply(_t, _this, args) {
      return impl.__call(...args);
    },
  });
  return proxy;
}

export const implOf = (o) => (o && o[IMPL]) || o;

/** Asignación a una llamada: f(args) = value */
export function callSet(f, args, value) {
  if (f && typeof f.__set === "function") {
    f.__set(...args, value);
    return value;
  }
  throw makeError(0x800a01b6, "El objeto no admite esta propiedad o método (asignación a llamada).");
}

/** Convierte valores "variant" (Field, elemento de Request...) a primitivos. */
export function prim(v) {
  if (v && (typeof v === "object" || typeof v === "function")) {
    const impl = implOf(v);
    if (impl && typeof impl.__value === "function") return impl.__value();
  }
  return v;
}

export function makeError(number, description) {
  const e = new Error(description);
  e.number = number | 0;
  e.description = description;
  return e;
}

// ---------------------------------------------------------------- Scripting.Dictionary
export class Dictionary {
  constructor() {
    this.__map = new Map();
    this.CompareMode = 0;
  }
  __k(k) {
    k = prim(k);
    return this.CompareMode === 1 && typeof k === "string" ? k.toLowerCase() : k;
  }
  __call(k) {
    return this.Item(k);
  }
  __set(k, v) {
    this.__map.set(this.__k(k), prim(v));
  }
  Add(k, v) {
    if (this.__map.has(this.__k(k))) throw makeError(0x800a01c9, "Esta clave ya está asociada a un elemento de esta colección.");
    this.__map.set(this.__k(k), prim(v));
  }
  Exists(k) {
    return this.__map.has(this.__k(k));
  }
  Item(k) {
    return this.__map.get(this.__k(k));
  }
  __set_Item(k, v) {
    this.__set(k, v);
  }
  Remove(k) {
    this.__map.delete(this.__k(k));
  }
  RemoveAll() {
    this.__map.clear();
  }
  Keys() {
    return vbArray([...this.__map.keys()]);
  }
  Items() {
    return vbArray([...this.__map.values()]);
  }
  get Count() {
    return this.__map.size;
  }
}

// Los arrays devueltos por COM se usaban con `new VBArray(x).toArray()`.
export function vbArray(arr) {
  const a = arr.slice();
  Object.defineProperty(a, "__vbarray", { value: true });
  return a;
}

export function VBArray(a) {
  const arr = Array.isArray(a) ? a : [];
  return {
    toArray: () => arr.slice(),
    ubound: () => arr.length - 1,
    lbound: () => 0,
    getItem: (i) => arr[i],
    dimensions: () => 1,
  };
}

/** Enumerator de JScript sobre colecciones COM y arrays. */
export function makeEnumerator(coll) {
  let items;
  const c = implOf(coll);
  if (c && typeof c.__items === "function") items = c.__items();
  else if (Array.isArray(coll)) items = coll;
  else if (c && c.__map instanceof Map) items = [...c.__map.keys()];
  else if (coll && typeof coll.length === "number") items = Array.from(coll);
  else items = [];
  let i = 0;
  return {
    atEnd: () => i >= items.length,
    item: () => items[i],
    moveFirst: () => {
      i = 0;
    },
    moveNext: () => {
      i++;
    },
  };
}

/* SEPAD — capa de compatibilidad con Internet Explorer 5/6 para navegadores modernos.
 * Se inyecta en todas las páginas HTML. El servidor además reescribe:
 *   x.frames("a")    -> x.__frames("a")
 *   x.children("a")  -> x.children.__c("a")   (también options, elements, rows...)
 *   x.tags("TD")     -> x.__tags("TD")
 */
(function () {
  "use strict";
  var w = window;

  // --- window.frames("nombre")
  w.__frames = function (n) {
    return typeof n === "number" ? this.frames[n] : this.frames[n] || (this.document && this.document.getElementsByName(n)[0] && this.document.getElementsByName(n)[0].contentWindow);
  };

  // --- colecciones invocables: coll("nombre") / coll(0)
  function collCall(k) {
    if (typeof k === "number" || /^\d+$/.test(String(k))) return this[Number(k)];
    if (this.namedItem) {
      var r = this.namedItem(k);
      if (r) return r;
    }
    for (var i = 0; i < this.length; i++) if (this[i] && (this[i].id === k || this[i].name === k)) return this[i];
    return null;
  }
  [w.HTMLCollection, w.NodeList, w.HTMLOptionsCollection, w.HTMLFormControlsCollection, w.HTMLAllCollection].forEach(function (C) {
    if (C && C.prototype && !C.prototype.__c) Object.defineProperty(C.prototype, "__c", { value: collCall, configurable: true });
  });

  // --- document.all.tags("TD"), elem.all
  function tags(t) {
    return (this.getElementsByTagName ? this : document).getElementsByTagName(t);
  }
  Object.defineProperty(Document.prototype, "__tags", { value: tags, configurable: true });
  Object.defineProperty(Element.prototype, "__tags", { value: tags, configurable: true });
  if (w.HTMLAllCollection) Object.defineProperty(HTMLAllCollection.prototype, "__tags", { value: function (t) { return document.getElementsByTagName(t); }, configurable: true });
  if (!("all" in Element.prototype)) {
    Object.defineProperty(Element.prototype, "all", {
      configurable: true,
      get: function () {
        var el = this;
        var f = function (k) {
          return typeof k === "number" ? el.getElementsByTagName("*")[k] : el.querySelector('[id="' + k + '"],[name="' + k + '"]');
        };
        var list = el.getElementsByTagName("*");
        f.length = list.length;
        f.tags = function (t) { return el.getElementsByTagName(t); };
        f.__tags = f.tags;
        f.item = f;
        return f;
      },
    });
  }

  // --- Date.getYear() de JScript: año completo a partir de 2000
  Date.prototype.getYear = function () {
    var y = this.getFullYear();
    return y >= 1900 && y < 2000 ? y - 1900 : y;
  };

  // --- attachEvent / detachEvent
  function attach(type, fn) {
    var self = this;
    var wrapped = fn.__ie || (fn.__ie = function (e) { w.event = e; return fn.call(self, e); });
    this.addEventListener(type.replace(/^on/, ""), wrapped, false);
    return true;
  }
  function detach(type, fn) {
    if (fn.__ie) this.removeEventListener(type.replace(/^on/, ""), fn.__ie, false);
  }
  [w, Document.prototype, Element.prototype].forEach(function (o) {
    if (!o.attachEvent) {
      Object.defineProperty(o, "attachEvent", { value: attach, configurable: true, writable: true });
      Object.defineProperty(o, "detachEvent", { value: detach, configurable: true, writable: true });
    }
  });

  // --- currentStyle
  if (!("currentStyle" in Element.prototype)) {
    Object.defineProperty(Element.prototype, "currentStyle", { configurable: true, get: function () { return getComputedStyle(this); } });
  }

  // --- document.createElement("<input name='x'>")
  var origCreate = Document.prototype.createElement;
  Document.prototype.createElement = function (tag) {
    if (typeof tag === "string" && tag.charAt(0) === "<") {
      var d = origCreate.call(this, "div");
      d.innerHTML = tag;
      return d.firstChild;
    }
    return origCreate.apply(this, arguments);
  };

  // --- showModalDialog: abre una ventana y, al cerrarse con returnValue,
  //     repite el clic original con el resultado disponible de forma síncrona.
  var modal = (w.__sepadModal = w.__sepadModal || { results: {}, pending: null });
  if (!w.showModalDialog) {
    w.showModalDialog = function (url, args, features) {
      var key = String(url);
      if (Object.prototype.hasOwnProperty.call(modal.results, key)) {
        var v = modal.results[key];
        delete modal.results[key];
        return v;
      }
      var ev = w.event;
      var f = String(features || "");
      var wd = parseInt((/dialogWidth\s*:\s*(\d+)/i.exec(f) || [])[1] || "520", 10);
      var ht = parseInt((/dialogHeight\s*:\s*(\d+)/i.exec(f) || [])[1] || "420", 10);
      modal.pending = { key: key, args: args, target: ev && ev.target, type: ev && ev.type };
      var win = w.open(url, "sepadModal", "width=" + wd + ",height=" + ht + ",resizable=yes,scrollbars=yes");
      if (win) win.dialogArguments = args;
      throw new Error("__sepad_modal_pending__");
    };
    w.__sepadModalDone = function (value) {
      var p = modal.pending;
      modal.pending = null;
      if (!p || value === undefined) return;
      modal.results[p.key] = value;
      setTimeout(function () {
        var t = p.target;
        if (t && t.nodeType === 3) t = t.parentNode;
        if (t && typeof t.click === "function" && (p.type === "click" || !p.type)) t.click();
        else if (t) t.dispatchEvent(new Event(p.type || "click", { bubbles: true }));
      }, 0);
    };
    w.addEventListener("error", function (e) {
      if (e && e.message && e.message.indexOf("__sepad_modal_pending__") >= 0) e.preventDefault();
    });
  }
  // Lado del diálogo: dialogArguments y returnValue
  try {
    if (w.opener && w.opener.__sepadModal && w.opener.__sepadModal.pending && w.name === "sepadModal") {
      w.dialogArguments = w.opener.__sepadModal.pending.args;
      w.addEventListener("pagehide", function () {
        try {
          w.opener.__sepadModalDone(w.returnValue);
        } catch (e) {}
      });
    }
  } catch (e) {}
})();

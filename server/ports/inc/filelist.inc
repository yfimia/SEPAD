<%
// Port a JScript de las funciones VBScript de Downloads.asp / Upload/default.asp.
// Mismo resultado: icono según la extensión y tamaño en Kb con un decimal si es < 1 Kb.

var EXT_ICONS = {
  txt: "text_plain.gif", bat: "executable.gif", exe: "executable.gif", com: "executable.gif",
  asp: "application_asp.gif", inc: "application_asp.gif", asa: "application_asp.gif",
  css: "application_css.gif", doc: "application_doc.gif",
  html: "application_html.gif", htm: "application_html.gif", shtml: "application_html.gif", phtml: "application_html.gif",
  pdf: "application_pdf.gif", xls: "application_xls.gif",
  bmp: "image_bmp.gif", gif: "image_gif.gif", jpg: "image_jpeg.gif", jpeg: "image_jpeg.gif",
  psd: "image_psd.gif", tif: "image_tiff.gif", tiff: "image_tiff.gif",
  rar: "pack_rar.gif", zip: "pack_zip.gif", gz: "pack_zip.gif", tar: "pack_zip.gif",
  eml: "mail.gif"
};

function extImage(filename) {
  var parts = String(filename).split(".");
  if (parts.length > 1) {
    var icon = EXT_ICONS[parts[parts.length - 1].toLowerCase()];
    if (icon) return icon;
  }
  return "attach.gif";
}

function sizeKb(bytes) {
  var kb = bytes / 1024;
  return kb < 1 ? Math.round(kb * 10) / 10 : Math.round(kb);
}

// Devuelve [{name, size, icon}] de una carpeta física (vacío si no existe).
function listFiles(physPath, createIfMissing) {
  var fso = Server.CreateObject("Scripting.FileSystemObject");
  var out = [];
  if (fso.FolderExists(physPath)) {
    var e = new Enumerator(fso.GetFolder(physPath).Files);
    for (; !e.atEnd(); e.moveNext()) {
      var f = e.item();
      out.push({ name: f.Name, size: sizeKb(f.Size), icon: extImage(f.Name) });
    }
  } else if (createIfMissing) {
    fso.CreateFolder(physPath);
  }
  return out;
}
%>

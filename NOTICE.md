# Avisos de copyright y licencia

Este repositorio se distribuye bajo la licencia **PolyForm Noncommercial 1.0.0** (ver [LICENSE.md](LICENSE.md)), salvo los componentes de terceros indicados más abajo. En resumen: se permite usar, estudiar, modificar y compartir el software **sin fines comerciales**, siempre que se conserven estos avisos. Cualquier uso comercial requiere autorización expresa de los titulares.

Quien redistribuya este software, o una obra derivada, debe incluir las siguientes líneas *Required Notice*.

Required Notice: Copyright (c) 1998-2004 Grupo RedWAY, Universidad Central "Marta Abreu" de Las Villas, y los autores de SEPAD: Alfredo Artiles Larralde, Yoilán Fimia León, Didiosky Benites Erice, Wilder González Días, Ricardo A. Sotolongo Trujillo, Enrique Portela García, Alcides Morales Guedes, Roberto Vicente, Ladys Rodríguez Contreras, Roberto Carlos Rodriguez Hidalgo, Victor Daniel Arteaga y Yenier Cruz.

Required Notice: Copyright (c) 2026 Yoilán Fimia León (https://github.com/yfimia), por la modernización: runtime en Node.js, migración de datos, capa de compatibilidad, ports, scripts, datos de demostración y documentación.

## Qué cubre cada aviso

| Parte | Titulares |
| --- | --- |
| Aplicación original SEPAD: `default.asp`, `global.asa`, `errorpage.asp`, `pages_esp/`, `Js/`, `Css/`, `images/`, `News/`, `Help/`, `XML/`, `XSL/`, `database/`, `Downloads/` y `legacy/` | Grupo RedWAY, UCLV y los autores de SEPAD |
| Modernización: `server/`, `scripts/`, `demo/`, `docs/`, `Makefile`, `Dockerfile`, `package.json` y los archivos de configuración | Yoilán Fimia León |

## Componentes de terceros (excluidos)

Estos componentes **no** están cubiertos por la licencia PolyForm Noncommercial. Se rigen por sus propias licencias, que prevalecen:

| Componente | Ubicación | Licencia |
| --- | --- | --- |
| TinyMCE (Moxiecode Systems) | `pages_esp/tinymce/` | GNU LGPL 2.0 (`pages_esp/tinymce/lgpl.txt`) |
| htmlArea (interactivetools.com) | `pages_esp/wysiwyg/` | Licencia de htmlArea (ver la cabecera de `editor.js`) |
| Implementación de MD5 de Paul Johnston y Greg Holt | `Js/md5.js`, `Js/md5.inc` | Licencia BSD (http://pajhome.org.uk/site/legal.html) |
| aspSmartUpload (Advantys) | `Services/aspSmartUpload/` | Licencia de su fabricante |
| Unzip32.dll (Info-ZIP) | `Services/UnPacker/Unzip32.dll` | Licencia Info-ZIP |
| Dependencias de npm | se instalan en `node_modules/` | La licencia de cada paquete |

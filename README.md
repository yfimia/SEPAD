# SEPAD

**Sistema de Enseñanza Personalizada a Distancia** — LMS created for the Cuban context.

SEPAD nació en 1998, en la Universidad Central "Marta Abreu" de Las Villas, como una plataforma de teleformación pensada para llegar a todos, sin importar la capacidad tecnológica o de conexión. Ofrecía aula virtual, foros, chat, mensajería interna, agenda, noticias, glosario, autoevaluaciones y un acceso alternativo por correo electrónico. Se escribió en **Active Server Pages (ASP) con JScript**, años antes de que existiera .NET.

Este repositorio conserva el código original y le agrega un **runtime moderno** que lo ejecuta en cualquier sistema (macOS, Linux, Windows) con Node.js y SQLite, **manteniendo intacta la interfaz de la época**: frames, GIFs, temas `aqua`/`gold`/`fresh`, la marquesina de noticias y hasta los mensajes de error de IIS.

## Probarlo

Requisitos: Node.js 20 o posterior y `xsltproc` (viene con macOS; en Debian/Ubuntu: `apt install xsltproc`).

```sh
make up           # instala, crea la base de datos y levanta http://localhost:8080/
make status       # ¿está corriendo? ¿en qué puerto?
make stop         # lo detiene
```

Si el puerto 8080 está ocupado, elige otro: `make up PORT=8082`. Con `make help` se ven todos los comandos (`restart`, `logs`, `reset`, `check`…).

Sin `make`:

```sh
npm install
npm run setup     # convierte database/Sepad.mdb a SQLite y carga datos de demostración
npm start         # http://localhost:8080/  (Ctrl+C para detener)
```

### Usuarios de demostración

| Usuario | Contraseña | Rol |
| --- | --- | --- |
| `alumno` | `alumno` | Alumno del *Diplomado en Tecnologías de Internet* |
| `profesor` | `profesor` | Profesora y coordinadora del diplomado |
| `administrador` | `admin` | Administrador del sistema |

Escribe el usuario y la contraseña arriba a la derecha y pulsa **Enter**, como en el original. Sin identificarte entras como invitado.

### Con Docker

```sh
make docker-run PORT=8082     # o: docker build -t sepad . && docker run -p 8082:8080 sepad
```

## Cómo funciona

Las páginas `.asp` originales no se reescribieron. Un servidor en Node.js las compila al vuelo y les da el entorno que esperaban de IIS:

- **ASP clásico**: `Request`, `Response`, `Session`, `Application`, `Server`, `global.asa`, `#include` y las particularidades de JScript, como `Session("x") = valor` o los nombres sin distinguir mayúsculas.
- **Datos**: ADODB sobre SQLite. El SQL de Access y SQL Server (`TOP`, `#fechas#`, `CONVERT`, `DATEDIFF(minute, …)`, joins anidados) se traduce automáticamente.
- **Componentes COM**: `Scripting.FileSystemObject`, `Microsoft.XMLDOM` (incluida la XSL de IE5 de 1998), `aspSmartUpload` y el envío de correo (simulado).
- **Navegador**: una capa de compatibilidad permite que el JavaScript escrito para Internet Explorer 5 funcione en navegadores actuales.
- **VBScript**: las 4 páginas escritas en VBScript se portaron a JScript (`server/ports/`).

> **Antes de publicarlo:** el código de 1998 construye las consultas SQL concatenando texto, como era habitual entonces. En un servidor público, restaura los datos con frecuencia (`SEPAD_RESET_HOURS=6`) y no guardes nada valioso. Más detalles en [Operación](docs/operacion.md#publicar-el-sitio).

## Documentación

| Documento | Contenido |
| --- | --- |
| [Operación](docs/operacion.md) | Comandos, variables de entorno, Docker y publicación |
| [La aplicación original](docs/aplicacion-original.md) | Qué hacía SEPAD, roles, interfaz y estructura de carpetas |
| [Arquitectura](docs/arquitectura.md) | Componentes del runtime y recorrido de una petición |
| [Runtime de ASP clásico](docs/runtime-asp.md) | Compilador, objetos de ASP, ADODB y componentes COM |
| [Base de datos](docs/base-de-datos.md) | Migración de Access a SQLite, traducción del SQL y datos de demostración |
| [Compatibilidad con el navegador](docs/compatibilidad-navegador.md) | Adaptaciones para el código escrito para Internet Explorer |
| [Desarrollo y pruebas](docs/desarrollo.md) | Convenciones, herramientas de prueba y tareas habituales |

## Estructura

| Carpeta | Contenido |
| --- | --- |
| `default.asp`, `global.asa`, `pages_esp/`, `Js/`, `Css/`, `images/`, `Help/` | La aplicación original |
| `database/Sepad.mdb` | La base de datos Access original |
| `Downloads/` | Herramientas cliente de SEPAD 1.5 |
| `server/` | El runtime moderno (ASP, ADO, COM, compatibilidad con IE y ports de VBScript) |
| `scripts/` | Migración, datos de demostración y herramientas de prueba |
| `demo/` | Contenido de los cursos de demostración |
| `docs/` | Documentación del proyecto |
| `legacy/` | Código original que ya no se usa, conservado como historia ([detalle](legacy/README.md)) |

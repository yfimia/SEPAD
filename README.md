# SEPAD

**Sistema de Enseñanza Personalizada a Distancia** — LMS created for the Cuban context.

SEPAD nació en 1998, en la Universidad Central "Marta Abreu" de Las Villas, como una plataforma de teleformación pensada para llegar a todos, sin importar la capacidad tecnológica o de conexión. Ofrecía aula virtual, foros, chat, mensajería interna, agenda, noticias, glosario, autoevaluaciones y un acceso alternativo por correo electrónico. Se escribió en **Active Server Pages (ASP) con JScript**, años antes de que existiera .NET.

Este repositorio conserva el código original y le agrega un **runtime moderno** que lo ejecuta en cualquier sistema (macOS, Linux, Windows) con Node.js, **manteniendo intacta la interfaz de la época**: frames, GIFs, temas `aqua`/`gold`/`fresh`, la marquesina de noticias y hasta los mensajes de error de IIS.

## Probarlo

Requisitos: Node.js 20 o posterior y `xsltproc` (viene con macOS; en Debian/Ubuntu: `apt install xsltproc`).

```sh
npm install
npm run setup     # convierte database/Sepad.mdb a SQLite y carga datos de demostración
npm start         # http://localhost:8080/
```

Usuarios de demostración:

| Usuario | Contraseña | Rol |
| --- | --- | --- |
| `alumno` | `alumno` | Alumno del *Diplomado en Tecnologías de Internet* |
| `profesor` | `profesor` | Profesora y coordinadora del diplomado |
| `administrador` | `admin` | Administrador del sistema |

También se puede entrar como invitado: basta con abrir la página principal.

Con Docker:

```sh
docker build -t sepad .
docker run -p 8080:8080 sepad
```

## Cómo funciona

Las páginas `.asp` originales no se reescribieron. Un servidor en Node.js las compila al vuelo y les da el entorno que esperaban de IIS:

- **ASP clásico**: `Request`, `Response`, `Session`, `Application`, `Server`, `global.asa`, `#include` y las particularidades de JScript, como `Session("x") = valor` o los nombres sin distinguir mayúsculas.
- **Datos**: ADODB sobre SQLite. El SQL de Access y SQL Server (`TOP`, `#fechas#`, `CONVERT`, `DATEDIFF(minute, …)`, joins anidados) se traduce automáticamente.
- **Componentes COM**: `Scripting.FileSystemObject`, `Microsoft.XMLDOM` (incluida la XSL de IE5 de 1998), `aspSmartUpload` y el envío de correo (simulado).
- **Navegador**: una capa de compatibilidad permite que el JavaScript escrito para Internet Explorer 5 funcione en navegadores actuales.

Variables de entorno: `PORT`, `SEPAD_RESET_ON_START=1` (restaura los datos de demostración al arrancar), `SEPAD_RESET_HOURS` (los restaura periódicamente) y `SEPAD_DEBUG=1`.

> Nota: el código de 1998 construye las consultas SQL concatenando texto, como era habitual entonces. Si se publica, conviene restaurar los datos con frecuencia (`SEPAD_RESET_HOURS`) y no guardar nada valioso.

## Estructura

- `default.asp`, `global.asa`, `pages_esp/`, `Js/`, `Css/`, `images/`, `Help/`: la aplicación original.
- `database/Sepad.mdb`: la base de datos Access original.
- `server/`: el runtime moderno (ASP, ADO, COM y compatibilidad con IE).
- `scripts/`: migración, datos de demostración y herramientas de prueba.
- `demo/`: contenido de los cursos de demostración.

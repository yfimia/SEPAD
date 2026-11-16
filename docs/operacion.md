# Operación

## Requisitos

- Node.js 20 o posterior.
- `xsltproc`, para mostrar los prólogos de los módulos. Viene con macOS; en Debian/Ubuntu se instala con `apt install xsltproc`.
- `make` (opcional, para usar los atajos de abajo).

## Comandos

El `Makefile` cubre las operaciones habituales. `make help` las lista todas.

| Comando | Qué hace |
| --- | --- |
| `make up PORT=8082` | Instala dependencias y crea la base de datos si hace falta, y levanta SEPAD en segundo plano |
| `make stop` | Detiene SEPAD. Solo afecta a los procesos de este proyecto, no a otros servidores Node |
| `make restart PORT=8082` | Detiene y vuelve a levantar |
| `make status` | Indica si está corriendo y en qué puerto |
| `make logs` | Sigue el log del servidor en segundo plano (`data/sepad.log`) |
| `make start` | Levanta en primer plano; se detiene con Ctrl+C |
| `make debug` | Como `start`, pero imprime la traza completa de cada error ASP |
| `make reset` | Rehace la base de datos y los cursos de demostración desde cero |
| `make check` | Compila todas las páginas `.asp` y reporta errores de sintaxis |
| `make crawl PORT=8082` | Recorre el sitio en ejecución como alumno, profesor y administrador |
| `make docker-build` / `make docker-run` | Construye y ejecuta la imagen Docker |
| `make clean` | Borra la base de datos generada (`data/`) y la copia de los cursos (`Courses/`) |

Sin `make`, el equivalente mínimo es `npm install`, `npm run setup` y `npm start`.

Si el puerto está ocupado, `make up` lo avisa y no arranca; elige otro con `PORT=`.

## Usuarios de demostración

| Usuario | Contraseña | Rol |
| --- | --- | --- |
| `alumno` | `alumno` | Alumno del *Diplomado en Tecnologías de Internet* |
| `profesor` | `profesor` | Profesora, coordinadora de la modalidad y dueña de los módulos |
| `administrador` | `admin` | Administrador del sistema |

Al abrir la página principal sin identificarse se entra como invitado. El formulario de acceso calcula el MD5 de la contraseña en el navegador al pulsar **Enter**, como en el original.

## Variables de entorno

| Variable | Valor por defecto | Efecto |
| --- | --- | --- |
| `PORT` | `8080` | Puerto HTTP |
| `SEPAD_DB` | `data/sepad.sqlite` | Base de datos en uso |
| `SEPAD_RESET_ON_START` | sin definir | Con `1`, restaura la base y los cursos de demostración al arrancar |
| `SEPAD_RESET_HOURS` | sin definir | Restaura los datos de demostración cada *n* horas |
| `SEPAD_DEBUG` | sin definir | Con `1`, imprime en consola la traza completa de los errores ASP |
| `SEPAD_LOG` | registra | Con `0`, deja de registrar cada petición |

## Docker

```sh
make docker-build
make docker-run PORT=8082
```

La imagen instala `xsltproc`, genera la base de datos al construirse y arranca con `SEPAD_RESET_ON_START=1` y `SEPAD_RESET_HOURS=6`. No incluye `legacy/` ni los datos locales.

## Publicar el sitio

El código de 1998 construye las consultas SQL concatenando texto y no escapa el HTML que escriben los usuarios. Es lo habitual en su época, pero en Internet significa que cualquiera puede alterar los datos. Para una demostración pública:

- Usa `SEPAD_RESET_HOURS` (y `SEPAD_RESET_ON_START=1`) para que los datos vuelvan a su estado inicial.
- No guardes información real ni reutilices contraseñas.
- Ejecútalo aislado (contenedor o máquina dedicada), detrás de un proxy con HTTPS.

Medidas que ya incorpora el runtime:

- No sirve `legacy/`, `server/`, `scripts/`, `data/`, `database/` ni `demo/`, ni archivos `.inc`, `.asa`, `.mdb` o `XML/config.xml`.
- Las subidas de archivos exigen sesión (o rol de administrador para las descargas generales), rechazan extensiones ejecutables y tienen un tope de tamaño.
- Ninguna página puede ejecutarse más de 15 segundos.
- El correo no se envía: queda registrado en la consola.

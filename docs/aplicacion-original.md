# La aplicación original

SEPAD nació en 1998 en la Universidad Central "Marta Abreu" de Las Villas (Santa Clara, Cuba) para llevar la enseñanza a distancia a estudiantes con conexiones lentas o intermitentes. Además del aula web, ofrecía acceso por correo electrónico y un cliente multimedia sin conexión. Los instaladores de esas herramientas (`SepadHP1.5.exe`, editor y publicador de módulos, y `SepadXM1.5.exe`, cliente por correo) están en `Downloads/`.

Corría en IIS con páginas ASP escritas en JScript, primero sobre una base de datos Access y después sobre SQL Server.

## Conceptos

- **Modalidad académica** (tabla `Modulo`): un diplomado o curso de posgrado. Tiene un coordinador, un *claustro* (grupo de profesores), un *grupo* de alumnos y un estado (en inscripción, en curso, finalizada…).
- **Módulo** (tabla `Cursos`): una asignatura dentro de la modalidad. Tiene un profesor principal y un prólogo en XML.
- **Lección** (tabla `Lecciones`): una carpeta dentro del módulo con sus materiales (`Ficheros` y `Ficheros_de_lecciones`).
- **Grupos** (`Grupos`, `Grupos_de_Usuarios`): determinan quién ve qué. Hay grupos predefinidos para invitados (4), todos los usuarios (60), administradores (70) y publicadores (72).

El contenido de cada módulo está en `Courses/courseN/`: el prólogo (`prologo.xml`, que se transforma con `XSL/prologo.xsl`), una carpeta por lección y, opcionalmente, `descargas/` y `Seminarios/`.

## Roles

| Rol | `PermissionType` | Qué puede hacer |
| --- | --- | --- |
| Invitado | (usuario `guest`, ID 580) | Ver la portada, las modalidades y los prólogos; solicitar matrícula |
| Usuario | 0 | Estudiar los módulos de sus modalidades, foros, chat, mensajería, agenda |
| Publicador | 10 | Además, crear y administrar módulos, tutorar alumnos, glosario y preguntas frecuentes |
| Administrador | 20 | Además, gestionar usuarios, grupos, modalidades, matrículas, noticias y estadísticas |

## Interfaz

Tras identificarse, `default.asp` muestra un *frameset* con tres zonas:

- **Superior** (`choicecourse.asp`): logo, usuario activo, marquesina de noticias y menú (Datos personales, Mensajería, Chat, Agenda, Noticias, Ayuda…).
- **Izquierda** (`selection.asp`): búsqueda y árbol de modalidades académicas.
- **Principal**: portada (`P0.asp`), la modalidad (`modulo.asp`), el módulo (`Course.asp`) o el **Aula Virtual** (`lessonBrowser.asp`). El Aula Virtual es otro frameset con el árbol de lecciones, los materiales (`docs.asp`) y una barra de herramientas (`browserframe.asp`) con foros, glosario, anotaciones, ejercicios, descargas y más.

Casi todas las herramientas se abren en ventanas emergentes (`abreVentana(...)` en `Js/Windows.js`). Los paneles de gestión son framesets con un árbol a la izquierda: `Administracion/` (administrador), `ModuloAdm/` (profesor del módulo), `ModuloTutor/` (tutor) y `ModAdm/` (coordinador de la modalidad).

Hay tres temas visuales (`aqua`, `gold`, `fresh`) en `Css/` e `images/`, elegidos por usuario (`Usuarios.skin`).

## Sesión y el parámetro `uid`

`global.asa` inicializa la aplicación (lee `XML/config.xml`) y cada sesión. `Session_OnStart` genera un identificador aleatorio en `Session("uid")`. Casi todas las páginas lo reciben en la URL (`?uid=...`) y, si no coincide con el de la sesión, redirigen a la página de error de "tiempo de espera vencido". Por eso abrir una página copiando una URL de otra sesión lleva al error, y es el comportamiento original.

`default.asp` es a la vez la página de entrada y la de autenticación: sin datos de formulario entra como invitado; con `userName` y `userPassword` (MD5 en mayúsculas) valida al usuario.

## Carpetas

| Carpeta | Contenido |
| --- | --- |
| `default.asp`, `global.asa`, `errorpage.asp` | Entrada, arranque de aplicación y sesión, error general |
| `pages_esp/` | Todas las páginas de la aplicación (en español) |
| `pages_esp/inc/` | Textos e includes de cada página |
| `Js/` | Bibliotecas de servidor (`*.inc`: acceso a datos, usuarios, módulos…) y de cliente (`*.js`) |
| `Css/`, `images/`, `News/` | Temas visuales e imágenes |
| `Help/` | Manuales de usuario, tutor, profesor, coordinador y administrador |
| `XML/`, `XSL/` | Configuración y la hoja de estilo de los prólogos |
| `Services/` | Componentes COM de terceros (subida de archivos, correo, descompresión) |
| `database/Sepad.mdb` | Respaldo de la base de datos Access |
| `Downloads/` | Herramientas cliente de SEPAD 1.5 |
| `legacy/` | Código original que ya no se usa (ver `legacy/README.md`) |

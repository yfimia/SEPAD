# Documentación de SEPAD

SEPAD (Sistema de Enseñanza Personalizada a Distancia) es un LMS escrito en 1998 con Active Server Pages y JScript. Esta rama lo ejecuta sobre un runtime moderno en Node.js y SQLite sin cambiar su interfaz.

| Documento | Para qué sirve |
| --- | --- |
| [Operación](operacion.md) | Levantar, detener y publicar el sitio; variables de entorno; Docker; seguridad |
| [La aplicación original](aplicacion-original.md) | Qué hace SEPAD, sus roles, la estructura de carpetas y cómo fluye una sesión |
| [Arquitectura](arquitectura.md) | Visión general del runtime moderno y el recorrido de una petición |
| [Runtime de ASP clásico](runtime-asp.md) | Compilador de páginas, objetos de ASP, componentes COM y manejo de errores |
| [Base de datos](base-de-datos.md) | Migración de Access a SQLite, traducción del SQL, datos de demostración y reinicio |
| [Compatibilidad con el navegador](compatibilidad-navegador.md) | Cómo funciona en navegadores actuales el JavaScript escrito para Internet Explorer 5 |
| [Desarrollo y pruebas](desarrollo.md) | Convenciones, herramientas de prueba y cómo portar o corregir páginas |

Para empezar rápido: `make up PORT=8082` y abrir `http://localhost:8082/` (usuario `alumno`, contraseña `alumno`).

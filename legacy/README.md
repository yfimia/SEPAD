# legacy/

Código original de SEPAD que ya no se usa en la versión modernizada. Se conserva como
pieza histórica, con las mismas rutas relativas que tenía en la raíz del proyecto.
El servidor no lo publica por HTTP.

- `Beta/`: páginas de prueba (`public.asp`, `uploadCourFile.asp`); ninguna página las enlazaba.
- `SEPAD_06_01_09_Agenda/`: parche de la Agenda (enero de 2006) con el script de SQL Server
  `Crea_Tabla_AgendaEventos.sql`; ya estaba aplicado en `pages_esp/`.
- `pages_esp/Administracion/pages_esp/`: copia duplicada de `pages_esp/` sin referencias.
- `pages_esp/Downloads.asp`, `pages_esp/Upload/` y `pages_esp/Administracion/Upload/`: páginas
  originales en VBScript; ahora se atienden con sus ports a JScript en `server/ports/`.
- `pages_esp/**/NewMsg(old).asp`, `Js/date_funtions.rar`: copias antiguas.
- `vssver.scc` y `Thumbs.db`: metadatos de Visual SourceSafe y de Windows.

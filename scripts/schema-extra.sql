-- Tablas que el código usa pero no estaban en el respaldo Sepad.mdb (el sistema
-- había pasado a SQL Server). AgendaEventos sigue Crea_Tabla_AgendaEventos.sql;
-- el resto se dedujo de las páginas que las usan.

CREATE TABLE IF NOT EXISTS "AgendaEventos" (
  "ID" INTEGER PRIMARY KEY AUTOINCREMENT,
  "OwnerName" TEXT,
  "OwnerId" INTEGER,
  "ToName" TEXT DEFAULT '',
  "ToUsr" INTEGER DEFAULT 0,
  "ToModulo" INTEGER DEFAULT 0,
  "ToCourse" INTEGER DEFAULT 0,
  "ToGroup" INTEGER DEFAULT 0,
  "ToSubGroup" INTEGER DEFAULT 0,
  "Subject" TEXT,
  "Content" TEXT,
  "Priority" INTEGER DEFAULT 0,
  "DateBegin" DATETIME,
  "DateEnd" DATETIME,
  "EventType" INTEGER DEFAULT 1,
  "Readed" BOOLEAN NOT NULL DEFAULT 1
);

CREATE TABLE IF NOT EXISTS "Mensajes_Enviados" (
  "ID" INTEGER PRIMARY KEY AUTOINCREMENT,
  "From" TEXT,
  "To" INTEGER,
  "Subject" TEXT,
  "Content" TEXT,
  "Priority" INTEGER,
  "Date" DATETIME DEFAULT (strftime('%Y-%m-%d %H:%M:%S', 'now', 'localtime')),
  "FromID" INTEGER,
  "Readed" BOOLEAN DEFAULT 0,
  "ToText" TEXT
);

CREATE TABLE IF NOT EXISTS "Glosarios" (
  "ID" INTEGER PRIMARY KEY AUTOINCREMENT,
  "IdCurso" INTEGER,
  "Palabra" TEXT,
  "Signif" TEXT
);

CREATE TABLE IF NOT EXISTS "faq" (
  "ID" INTEGER PRIMARY KEY AUTOINCREMENT,
  "course" INTEGER,
  "question" TEXT,
  "response" TEXT
);

CREATE TABLE IF NOT EXISTS "HistoriaLecciones" (
  "ID" INTEGER PRIMARY KEY AUTOINCREMENT,
  "UserID" INTEGER,
  "Lesson" INTEGER,
  "Course" INTEGER,
  "fullUserName" TEXT,
  "lessonName" TEXT,
  "Cant" INTEGER DEFAULT 0
);

CREATE TABLE IF NOT EXISTS "Modulo_Curso" (
  "ID" INTEGER PRIMARY KEY AUTOINCREMENT,
  "Modulo" INTEGER,
  "Curso" INTEGER
);

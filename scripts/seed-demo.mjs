// Carga datos de demostración en data/sepad.sqlite y copia los cursos de demo/ a Courses/.
// Uso: node scripts/seed-demo.mjs
import fs from "node:fs";
import path from "node:path";
import crypto from "node:crypto";
import Database from "better-sqlite3";

const root = path.resolve(path.dirname(new URL(import.meta.url).pathname), "..");
const dbFile = path.join(root, "data/sepad.sqlite");
const db = new Database(dbFile);
const md5 = (s) => crypto.createHash("md5").update(s, "latin1").digest("hex").toUpperCase();
const pad = (n) => String(n).padStart(2, "0");
const fmt = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}:00`;
const now = new Date();
const daysFromNow = (n, h = 10) => {
  const d = new Date(now);
  d.setDate(d.getDate() + n);
  d.setHours(h, 0, 0, 0);
  return fmt(d);
};

const ins = (table, row) => {
  const cols = Object.keys(row);
  const r = db.prepare(`INSERT INTO "${table}" (${cols.map((c) => `"${c}"`).join(",")}) VALUES (${cols.map(() => "?").join(",")})`).run(...cols.map((c) => row[c]));
  return Number(r.lastInsertRowid);
};

// IDs fijos para que el contenido de demo/Courses/courseN coincida
const PROF = 2000;
const ALUMNO = 2001;
const G_ALUMNOS = 100;
const G_CLAUSTRO = 101;
const MOD = 1;
const ADMIN = 1830;

db.transaction(() => {
  // --- Usuarios de demostración (contraseñas conocidas)
  db.prepare(`UPDATE Usuarios SET Password = ?, Email = 'admin@sepad.demo', skin = 'aqua' WHERE ID = ?`).run(md5("admin"), ADMIN);
  const user = (ID, Name, pass, PermissionType, FullName, extra = {}) =>
    ins("Usuarios", { ID, Name, Password: md5(pass), PermissionType, FullName, Email: `${Name}@sepad.demo`, Logins: 0, fechaIngreso: daysFromNow(-200), sexo: 0, forward: 0, skin: "aqua", Pais: "Cuba", Provincia: "Villa Clara", Institucion: "Universidad (demostración)", ...extra });
  user(PROF, "profesor", "profesor", 10, "Profesora de demostración", { titulo: "Lic. en Ciencias de la Computación", sexo: 0 });
  user(ALUMNO, "alumno", "alumno", 0, "Alumno de demostración", { profesion: "Ingeniero", sexo: 1 });

  // --- Grupos
  ins("Grupos", { ID: G_ALUMNOS, Name: "Diplomado Internet - Alumnos", Description: "Alumnos matriculados en el diplomado de demostración" });
  ins("Grupos", { ID: G_CLAUSTRO, Name: "Diplomado Internet - Claustro", Description: "Profesores y tutores del diplomado de demostración" });
  for (const [u, g] of [
    [PROF, 60],
    [PROF, 72],
    [PROF, G_CLAUSTRO],
    [ALUMNO, 60],
    [ALUMNO, G_ALUMNOS],
  ])
    ins("Grupos_de_Usuarios", { User: u, Group: g });

  // --- Modalidad académica
  ins("Modulo", {
    ID: MOD,
    Name: "Diplomado en Tecnologías de Internet",
    state: 1,
    grupo: G_ALUMNOS,
    cordinador: PROF,
    claustro: G_CLAUSTRO,
    tipo: 1,
    diploma: "Diplomado",
    fechas: "Del 15 de febrero al 30 de abril de 1999",
    org: "Centro de Estudios a Distancia (demostración)",
    objetivos: "Formar a profesionales en el uso de Internet como herramienta de estudio y trabajo.",
    programa: "Módulo 1: Introducción a Internet.<br>Módulo 2: Diseño de páginas Web con HTML.",
    coste: "Gratuito",
    fprevia: "Graduados universitarios.",
    app: "Solicitud de matrícula a través de SEPAD.",
    doc: "Ninguna.",
    inst: "Contenido de demostración",
    Fundament: "La educación a distancia necesita que estudiantes y profesores dominen los servicios básicos de Internet.",
    Sede: "Santa Clara, Villa Clara",
    Bibliog: "Krol, E. <i>The Whole Internet User's Guide &amp; Catalog</i>. O'Reilly, 1994.",
  });

  // --- Módulos (cursos), prólogos y lecciones
  const courses = [
    { ID: 1, Name: "Introducción a Internet", lessons: ["¿Qué es Internet?", "El World Wide Web", "Correo electrónico"] },
    { ID: 2, Name: "Diseño de páginas Web con HTML", lessons: ["Etiquetas básicas de HTML", "Tablas y marcos"] },
  ];
  for (const c of courses) {
    const prolog = ins("Ficheros", { url: "prologo.xml", tama: 1, description: "Prólogo del módulo", title: "Prólogo" });
    ins("Cursos", { ID: c.ID, Name: c.Name, Prologue: prolog, Atmost: 50, Owner: PROF, ctdExAct: 3, ctdExRet: 2, MaxFile: 1024, MaxMem: 10240, state: 1, modulo: MOD });
    ins("Modulo_Curso", { Modulo: MOD, Curso: c.ID });
    c.lessons.forEach((name, i) => {
      const dir = `leccion${i + 1}`;
      const lid = ins("Lecciones", { Name: name, dir, Course: c.ID, Primary: i === 0 ? 1 : 0, Nivel: 0, LOM: 0, state: 1, order: i + 1, orientaciones: "Lea el material y participe en el foro." });
      const size = fs.statSync(path.join(root, "demo/Courses", `course${c.ID}`, dir, "index.htm")).size;
      const fid = ins("Ficheros", { url: "index.htm", tama: Math.ceil(size / 1024), description: `Texto de la lección: ${name}`, title: name });
      ins("Ficheros_de_lecciones", { fileID: fid, lessonID: lid });
    });
  }

  // --- Noticias y anuncios
  const news = [
    ["Comienza el Diplomado en Tecnologías de Internet", "Ya está abierta la matrícula para el diplomado de demostración. Solicite su inscripción desde la página de la modalidad.", -3, "Nacionales.gif"],
    ["SEPAD vuelve a estar en línea", "Esta es una restauración del SEPAD original de 1998, ejecutándose sobre un runtime moderno que conserva la interfaz de la época.", -1, "Others.gif"],
  ];
  for (const [Titulo, Cuerpo, d, Imagen] of news) ins("Noticias", { Titulo, Cuerpo, Fecha_Publicada: daysFromNow(d), Imagen, Url: "", Expira: 0, FechaExpira: null, Publisher: ADMIN });
  const an = ins("Anuncios", { Titulo: "Bienvenidos al módulo", Cuerpo: "Recuerden presentarse en el foro antes del viernes.", Expira: 0, FechaExpira: null, Publisher: PROF, Fecha_Publicada: daysFromNow(-2) });
  ins("AnunciosCurso", { Anuncio: an, curso: 1 });

  // --- Foro
  const topic = ins("ForumTopics", { Curso: 1, User: PROF, Titulo: "Presentaciones", Descripcion: "Preséntese al grupo: nombre, institución y qué espera del módulo.", Fecha: daysFromNow(-5), Activo: 1 });
  const msg = ins("ForumMessages", { User: PROF, Texto: "Bienvenidos. Soy la profesora del módulo; cuéntennos de dónde se conectan.", Fecha: daysFromNow(-5, 11), ReadingCount: 3, AceptaRespuestas: 1, Topic: topic });
  ins("ForumAnswers", { Texto: "Hola a todos, me conecto desde Santa Clara por la red de la universidad.", User: ALUMNO, fecha: daysFromNow(-4, 9), Origen: msg });
  ins("ForumTopics", { Curso: 1, User: PROF, Titulo: "Dudas sobre la lección 2", Descripcion: "Preguntas sobre el World Wide Web.", Fecha: daysFromNow(-2), Activo: 1 });

  // --- Mensajería interna
  ins("Mensajeria", { From: "Profesora de demostración", To: ALUMNO, Subject: "Bienvenido a SEPAD", Content: "Hola:\r\n\r\nYa puede acceder a los módulos del diplomado desde el menú de la izquierda.\r\n\r\nSaludos.", Priority: 0, Date: daysFromNow(-1), FromID: PROF, Readed: 0, ToText: "alumno" });

  // --- Agenda, glosario, preguntas frecuentes, chat
  ins("AgendaEventos", { OwnerName: "Profesora de demostración", OwnerId: PROF, ToName: "SEPAD", ToGroup: 60, Subject: "Tutoría en línea", Content: "Sala de charlas general, de 10:00 a 11:00.", Priority: 1, DateBegin: daysFromNow(2, 10), DateEnd: daysFromNow(2, 11), EventType: 1, Readed: 0 });
  for (const [Palabra, Signif] of [
    ["HTML", "Lenguaje de marcas de hipertexto con el que se escriben las páginas Web."],
    ["HTTP", "Protocolo de transferencia de hipertexto."],
    ["Navegador", "Programa para ver páginas Web, como Netscape Navigator o Internet Explorer."],
    ["URL", "Dirección única de un recurso en Internet."],
  ])
    ins("Glosarios", { IdCurso: 1, Palabra, Signif });
  ins("faq", { course: 1, question: "¿Cómo entrego las tareas?", response: "Use la opción de subir archivos del módulo." });
  ins("faq", { course: 1, question: "¿Puedo estudiar sin conexión permanente?", response: "Sí, SEPAD ofrece acceso por correo electrónico." });
  const sala = ins("Salas", { Nombre: "Sala general", Objetivo: "Charla libre entre alumnos y profesores", Fecha: daysFromNow(-10), Moderador: PROF, Publica: 1 });
  ins("SalasCurso", { curso: 1, sala });

  // --- Configuración (como la dejaba la página de configuración)
  db.prepare(`UPDATE Config SET HomePage = 'P0.asp'`).run();
})();
db.close();

// Copiar el contenido de los cursos
const src = path.join(root, "demo/Courses");
const dst = path.join(root, "Courses");
fs.rmSync(dst, { recursive: true, force: true });
fs.cpSync(src, dst, { recursive: true });
fs.copyFileSync(dbFile, path.join(root, "data/sepad.seed.sqlite"));
console.log("Datos de demostración cargados. Usuarios: administrador/admin, profesor/profesor, alumno/alumno");

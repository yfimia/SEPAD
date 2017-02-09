<%@ Language=JavaScript%>
<%
 Response.Expires = -1;
 
 %>
 <!-- #include file="../js/AdoLibrary.inc" -->
 <!-- #include file="../js/user.inc" -->
 <!-- #include file='../js/course.inc' -->
 <script language="vbscript" RUNAT="Server">
 function getAtualTime
  getAtualTime = Now()
 End function 

</script>

 <%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       


 if (Request.QueryString.Count == 2)   { 


   fullCourse = Session("FullCourse");
   if (Request.QueryString.Item("init") == 1) {
     Session("lastLesson") = fullCourse.firstLessonID;
   }
   else {
     if (fullCourse.lastLessonID != -1)
       Session("lastLesson") = fullCourse.lastLessonID;
     else Session("lastLesson") = fullCourse.firstLessonID;  
   }
 }
 else if (Request.QueryString.Count == 4)   { 
   Session("course")        = Request.QueryString("ID") + "";
   Session("courseName")    = Request.QueryString("Name") + "";
   Session("FullCourse")    = GetFullCourse(Session("course"));
   fullCourse = Session("FullCourse");
   Session("courseOwner")    = fullCourse.owner;

   if (Request.QueryString.Item("init") == 1) {
     Session("lastLesson") = fullCourse.firstLessonID;
   }
   else {
     if (fullCourse.lastLessonID != -1)
       Session("lastLesson") = fullCourse.lastLessonID;
     else Session("lastLesson") = fullCourse.firstLessonID;  
   }
   
   Session("lastLessonIndex") = GetLessonIndex(Session("lastLesson"));
   
//   Response.Write(fullCourse.firstLessonID + " *** "  + fullCourse.lastLessonID);

 }
 else {
   //Error en el pase de parametros.
   Response.Redirect("errorpage.asp?tipo=Error&short=" + BAD_PARAMS_SHORT  + "&desc=" + BAD_PARAMS_TEXT);
 }    
     
   //Innsercion en la base de datos de Conexiones a Cursos...
   insCtocourse(Session("userID"), Session("course"));

  //Estableciendo coneccion para las lecciones..  

    var oConn;
    var oRec;
 
    var filePath = Application('dataPath');
    
    oConn = MakeConnection( oConn, filePath );       
  
    //obtener subgrupo y tutor
    Sql = "SELECT SubGrupo.ID as SId, SubGrupo.Name as SName, Usuarios.ID as UID, Usuarios.Name, Usuarios.Email, Usuarios.FullName " + 
          "FROM Usuarios INNER JOIN (SubGrupo INNER JOIN UsuariosSubGrupo ON SubGrupo.ID = UsuariosSubGrupo.Subgrupo) ON Usuarios.ID = SubGrupo.Tutor " + 
          "WHERE (((UsuariosSubGrupo.Usuario)=" + Session("userId") + ") AND ((SubGrupo.Curso)=" + Session("course") + "))" ;

   //Response.Redirect("errorpage.asp?tipo=Error&short=" + Sql  + "&desc=" + SESSION_TIMEOUT_TEXT);

    oRec = Query( Sql, oRec, oConn  );		

    if (!oRec.EOF)
     { 
      Session("SgID")      = oRec.Fields.Item("SId") + "";
      Session("sgName")      = oRec.Fields.Item("SName") + "";
      Session("sgTutor")    = oRec.Fields.Item("Name") + "";
      Session("tutID")    = oRec.Fields.Item("UID") + "";
      
      Session("sgEmail")    = oRec.Fields.Item("Email") + "";
      Session("sgTutorName")   = oRec.Fields.Item("FullName") + "";
     }
     else {
      Session("sgName")      = "";
      Session("SgId")      = "";
      Session("sgTutor")    = "";
      Session("sgEmail")    = "";
      Session("sgTutorName")   = "";
     }
    
    oRec.Close();

   
      //Consulta para las Lecciones...
   
    Sql ="SELECT A.Id,A.Name,A.dir, A.Nivel,A.[Primary], A.state FROM Lecciones A WHERE (Course = " +  Session("course") + ") ORDER BY [order] asc, ID asc";

    oRec = Query( Sql, oRec, oConn  );		
   
           
   /*  Guardando el recordset(lecciones) en la Session   */    
    
    //Objetos...
    function items(id,name, dir, nivel,primary, state)
     { 
       this.id = id;
       this.name = name;
       this.dir = dir;       
       this.nivel = nivel;
       this.primary = primary; 
       this.state = state;
     }
    
    function record(id,name, dir, nivel,primary, state)
     { 
      this.items = new items(id,name, dir, nivel,primary, state);
     } 
     
    // Variables Globales (campos de la base de datos )para las  leciones... 
    var lessonId = new Array();
    var lessonDir = new Array();
    var lessonName = new Array();
    var lessonNivel = new Array();
    var lessonPrimary = new Array();
    var lessonState = new Array();

       /* Inicializacion de los campos */  
  
      // para el Servidor....
    var i = 0;   
    while (oRec.EOF == false)
     { 
      lessonId[i]      = oRec.Fields.Item("Id") + "";
      lessonName[i]    = oRec.Fields.Item("Name") + "";
      lessonDir[i]    = oRec.Fields.Item("dir") + "";
      lessonNivel[i]   = oRec.Fields.Item("Nivel") + "";
      lessonPrimary[i] = oRec.Fields.Item("Primary") + "";
      lessonState[i] = oRec.Fields.Item("state") + "";
      
      oRec.Move(1);
      i += 1;
     }
     Session("lessonCant") = i - 1;  
     Session("lessons") = new record(lessonId,lessonName,lessonDir, lessonNivel,lessonPrimary, lessonState); 

  DestroyAdoObjects( oConn, oRec );    
 
%>
<html>
<head>
<title>Sistema de Ense&ntilde;anza Personalizada a Distancia...</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<frameset rows="*, 50" frameborder="NO" border="0" framespacing="0" cols="*"> 
<%
  if (Session("lastLesson") == -2) {
%>
   <frame name="mainFrame1" scrolling="YES" noresize src="errorpage.asp?tipo=Aviso&short=El módulo está vacío&desc=El módulo que intenta acceder no tiene lecciones aún." >
<%
  } 
  else {
%>
   <frame name="mainFrame1" scrolling="YES" noresize src="docs.asp?uid=<%=uid%>&lesson=<%=Session("lastLesson")%>&material=" >
<%
  }
%>   
   <frame name="bottomFrame" src="browserframe.asp?uid=<%=uid%>">
   
</frameset>

<noframes> 
<body bgcolor="#FFFFFF" text="#000000">
</body>
</noframes> 
</html>
<%
 }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>

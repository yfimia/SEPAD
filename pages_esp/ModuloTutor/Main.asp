<%@ Language=Jscript%>
<%
Response.Expires = 0;
%>
<!-- #include file='../../js/adolibrary.inc' -->
<!-- #include file='../../js/user.inc' -->
<!-- #include file='../../js/course.inc' -->
<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";    
      var admCourse = GetAdmCourse( Request.QueryString.Item("course") + "" );   
      
/*      Session("tutcurso") = Request.QueryString.Item("course") + "";       
      Session("tutcursoName") = Request.QueryString.Item("courseName") + "";       
      Session("tutcursoOwner") = Request.QueryString.Item("courseOwner") + "";       
      Session("tutcursomodulo")        =      Session("modulo");    
      Session("tutcursomoduloName")    =      Session("moduloName");
      Session("tutcursostate")         =      Session("state");     
      Session("tutcursocordinador")    =      Session("cordinador");
      Session("tutcursogrupo")         =      Session("grupo");     
      Session("tutcursoclaustro")      =      Session("claustro");  
*/      
      
      Session("tutcurso") = admCourse.ID;       
      Session("tutcursoName") = admCourse.Name;       
      Session("tutcursoOwner") = admCourse.Owner;
      Session("tutcursomodulo")        =      admCourse.Modulo;
      Session("tutcursomoduloName")    =      admCourse.ModuloName;
      Session("tutcursostate")         =      admCourse.state;     
      Session("tutcursocordinador")    =      admCourse.Cordinador;
      Session("tutcursogrupo")         =      admCourse.grupo;     
      Session("tutcursoclaustro")      =      admCourse.claustro;  

      Session("tutcursomaxmem")      =      admCourse.MaxMem;  
      Session("tutcursomaxfile")      =      admCourse.MaxFile;  
      Session("tutcursomodstate")      =      admCourse.ModuloState;  
      
      Session("isTutor") = -1;
      var Tutor = isTutorCurso(Session("userID"), Session("tutcurso"));
      Session("isTutor") = Tutor.ID;
      Session("SubGrupo") = Tutor.Name;


//Response.Redirect("../errorpage.asp?tipo=Error&short=" + admCourse.ID + " *** " + Session("tutcursoclaustro")  + "&desc=" + SESSION_TIMEOUT_TEXT);
 if ((Session("isTutor") > -1) || (Session("PermissionType") == ADMINISTRATOR) || (Session("userID") == Session("tutcursocordinador"))  || (Session("userID") == Request.QueryString.Item("courseOwner")))
                          {
%>

<html>
<head>
<% if (Session("SubGrupo") != "") { %>
  <title>Tutoría del SubGrupo <%=Session("SubGrupo")%> en el módulo: <%=Session("tutcursoName")%></title>
<% }
  else {
%>
  <title>Tutoría del módulo: <%=Session("tutcursoName")%></title>
<%  
  }
%>  
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<frameset cols="260,*" frameborder="YES" border="0" framespacing="0" rows="*"> 
    <frame name="leftFrame" scrolling="YES" src="AdminTree.asp?uid=<%=uid%>&courseOwner=<%=Request.QueryString.Item("courseOwner")%>">
    <frame name="mainFrame" src="MainFrame.htm">
</frameset>
<noframes>
<body bgcolor="#FFFFFF" text="#000000">
</body>
</noframes> 
</html>
<%
    }
  else
    {
      Response.Redirect("../errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);     
    }

 }
 else
   Response.Redirect("../errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>

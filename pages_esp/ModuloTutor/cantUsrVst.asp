<%@ Language=JavaScript %>
<!-- #include file="../../js/Adolibrary.inc" -->
<!-- #include file="../../js/library.inc" -->
<!-- #include file='../../js/user.inc' -->
<%
 Response.Expires = -1;
if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       


 var courseID = Session("tutcurso") + "";
 
%>
<%
 if ((Session("isTutor") > -1) || (Session("PermissionType") == ADMINISTRATOR) || (Session("userID") == Session("tutcursocordinador"))  || (Session("userID") == Request.QueryString.Item("courseOwner")))
    {
    first = "-1";
    where = "";
    if (Request.QueryString.Count >= 3) {
      first = Request.QueryString.Item("first") + "";
      
      if (first != "-1") {
        where = " and (Usuarios.Name > '" +  first + "')";
      } 
    } 
    
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Visitas al curso</title>
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/EvalResult.css" type="text/css">

</head>
<body bgcolor="#FFFFFF" text="#000000" style="overflow-y:scroll">
 <%
  
 if (courseID != "") {

  if (Session("isTutor") > -1) {
   Sql = "SELECT Usuarios.Name AS Indentificador, Usuarios.FullName AS Nombre, Count(Conexiones_a_Cursos.[User]) AS Visitas FROM (Usuarios INNER JOIN Conexiones_a_Cursos ON Usuarios.ID = Conexiones_a_Cursos.[User]) INNER JOIN UsuariosSubGrupo ON Conexiones_a_Cursos.[User] = UsuariosSubGrupo.Usuario   GROUP BY Usuarios.Name, Usuarios.FullName, Conexiones_a_Cursos.Course, UsuariosSubGrupo.Subgrupo HAVING (Conexiones_a_Cursos.Course=" +  courseID + ") and (UsuariosSubGrupo.Subgrupo = " + Session("isTutor") +  ") " + where +  " order by Usuarios.Name";
  }
  else   
    Sql = "SELECT Usuarios.Name AS Indentificador, Usuarios.FullName AS Nombre, Count(Conexiones_a_Cursos.[User]) AS Visitas FROM (Usuarios INNER JOIN Conexiones_a_Cursos ON Usuarios.ID = Conexiones_a_Cursos.[User])  INNER JOIN Grupos_de_Usuarios ON Conexiones_a_Cursos.[User] = Grupos_de_Usuarios.[User]   GROUP BY Grupos_de_Usuarios.[Group], Usuarios.Name, Usuarios.FullName, Conexiones_a_Cursos.Course  HAVING (Conexiones_a_Cursos.Course=" +  courseID + ") " + where +  " and (Grupos_de_Usuarios.[Group] = " + Session("tutcursogrupo") +  ") " + " order by Usuarios.Name";
   
    putTable(Application("dataPath"), 'Visitas al curso ' + Session("tutcursoname"), "../../images/" + Session("skin") + "/ReportesIMG.gif", Sql, 3, "Indentificador"); 
    
 }   
 %>
</body>
 <table border="0" cellspacing="0" cellpadding="0" class="ToolBar" width="100%">
  <tr>
    <td>
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td><a href="cantusrvst.asp?uid=<%=uid%>&courseOwner=<%=Request.QueryString.Item("courseOwner")%>&IDCurso=<%=courseID%>&first=<%=last%>" class="ToolLink">&nbsp;Pr�ximos&nbsp;<%=SHOW_CANT%>&nbsp;</a></td>
        </tr>
      </table>
    </td>
  </tr>
</table>


</html>


<%
   }
  else
    Response.Redirect("../errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);     
 }
 else
   Response.Redirect("../errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>
	

































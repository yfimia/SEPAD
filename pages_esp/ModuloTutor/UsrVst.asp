<%@ Language=JavaScript %>
<!-- #include file="../../js/Adolibrary.inc" -->
<!-- #include file="../../js/library.inc" -->
<!-- #include file='../../js/user.inc' -->
<!-- #include file='inc/UsrVst.inc' -->
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
        where = " and (Conexiones_a_Cursos.Date < " + Application("dchar") +  first + Application("dchar") + ")";
      } 
    } 
    
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title><%= TXT_TITLE %></title>
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/EvalResult.css" type="text/css">

</head>
<body bgcolor="#FFFFFF" text="#000000" style="overflow-y:scroll">
 <%
 if (courseID != "") {
 
  if (Session("isTutor") > -1) {
 
   Sql = "SELECT Usuarios.Name as Identificador, Usuarios.FullName as Nombre, Conexiones_a_Cursos.Date as Fecha " +
         "FROM (Usuarios INNER JOIN Conexiones_a_Cursos ON Usuarios.ID = Conexiones_a_Cursos.[User]) INNER JOIN UsuariosSubGrupo ON Conexiones_a_Cursos.[User] = UsuariosSubGrupo.Usuario  " +
         "WHERE (((Conexiones_a_Cursos.Course)=" +  courseID + ")) and (UsuariosSubGrupo.Subgrupo = " + Session("isTutor") +  ") "  + where +  " " +
         "ORDER BY Conexiones_a_Cursos.Date DESC";
   }
   else {
   Sql = "SELECT Usuarios.Name as Identificador, Usuarios.FullName as Nombre, Conexiones_a_Cursos.Date as Fecha " +
         "FROM Usuarios INNER JOIN Conexiones_a_Cursos ON Usuarios.ID = Conexiones_a_Cursos.[User] " +
         "WHERE (((Conexiones_a_Cursos.Course)=" +  courseID + ")) "  + where +  " " +
         "ORDER BY Conexiones_a_Cursos.Date DESC";
   }
            
 //Response.Write(Sql);
   putTable(Application("dataPath"), TXT_TITLE, "../../images/" + Session("skin") + "/ReportesIMG.gif", Sql, -1, TEXT2); 
 }   
 %>
 <table border="0" cellspacing="0" cellpadding="0" class="ToolBar" width="100%">
  <tr>
    <td>
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td><a href="usrvst.asp?uid=<%=uid%>&courseOwner=<%=Request.QueryString.Item("courseOwner")%>&IDCurso=<%=courseID%>&first=<%=last%>" class="ToolLink">&nbsp;<%= TEXT1 %>&nbsp;<%=SHOW_CANT%>&nbsp;</a></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>


</html>


<%
   }
  else
    Response.Redirect("../errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);     
 }
 else
   Response.Redirect("../errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>
	

































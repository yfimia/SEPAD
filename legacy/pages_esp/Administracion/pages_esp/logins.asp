<%@ Language=JavaScript %>
<%
 Response.Expires = -1;
%>
<!-- #include file="../js/Adolibrary.inc" -->
<!-- #include file="../js/library.inc" -->
<!-- #include file='../js/user.inc' -->
<%
if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       


 
%>
<%
  if ((Session("PermissionType") == ADMINISTRATOR) || (Session("PermissionType") == PUBLICATOR))
    {
    first = "-1";
    where = "";
    if (Request.QueryString.Count >= 2) {
      first = Request.QueryString.Item("first") + "";
      
      if (first != "-1") {
        where = " where (Usuarios.Name > '" +  first + "')";
      } 
    } 

%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/EvalResult.css" type="text/css">

</head>
<body bgcolor="#FFFFFF" text="#000000" style="overflow-y:scroll">
 <%
   Sql = "SELECT Usuarios.Name as Identificador, Usuarios.FullName as Nombre, Usuarios.Logins AS Visitas " +
         "FROM Usuarios " + where +
         "ORDER BY Usuarios.Name";
   putTable(Application("dataPath"), 'Visitas a SEPAD', "../images/" + Session("skin") + "/ReportesIMG.gif", Sql, 3, "Identificador"); 
 %>
 <table border="0" cellspacing="0" cellpadding="0" class="ToolBar" width="100%">
  <tr>
    <td>
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td><a href="logins.asp?uid=<%=uid%>&first=<%=last%>" class="ToolLink">&nbsp;Próximos&nbsp;<%=SHOW_CANT%>&nbsp;</a></td>
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
    Response.Redirect("errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);     
 }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>
	

































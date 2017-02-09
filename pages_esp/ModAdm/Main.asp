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
 var amdModulo = GetAdmModulo(Request.QueryString.Item("modulo"));
 Session("admmodulo")        =  amdModulo.ID;    
 Session("admmoduloName")    =  amdModulo.Name;
 Session("admstate")        =   amdModulo.state;     
 Session("admcordinador")    =  amdModulo.Cordinador;
 Session("admgrupo")        =   amdModulo.grupo;     
 Session("admclaustro")    =    amdModulo.claustro;  
%>
<%
if ((Session("PermissionType") == ADMINISTRATOR) || (Session("userID") == Session("admcordinador")))
    {
%>

<html>
<head>
<title>Administraci&oacute;n de la modalidad académica: <%=Session("admmoduloname")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<frameset cols="260,*" frameborder="YES" border="0" framespacing="0" rows="*"> 
    <frame name="leftFrame" scrolling="YES" src="AdminTree.asp?uid=<%=uid%>">
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

<%@ Language=Jscript%>
<%
Response.Expires = 0;
%>

<!-- #include file='../../js/user.inc' -->
<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       
      Session("admcurso") = Request.QueryString.Item("course") + "";       
      Session("admcursoName") = Request.QueryString.Item("courseName") + "";       
      Session("admcursoOwner") = Request.QueryString.Item("courseOwner") + "";       
      Session("admcursomodulo")        =      Request.QueryString.Item("modulo") + "";    
      Session("admcursomoduloName")    =      Request.QueryString.Item("moduloName") + "";
      Session("admcursostate")         =      Request.QueryString.Item("state") + "";     
      Session("admcursocordinador")    =      Request.QueryString.Item("cordinador") + "";
      Session("admcursogrupo")         =      Request.QueryString.Item("grupo") + "";     
      Session("admcursoclaustro")      =      Request.QueryString.Item("claustro") + "";  


 if ((Session("PermissionType") == ADMINISTRATOR) || (Session("userID") == Session("admcursocordinador"))  || (Session("userID") == Session("admcursoOwner")))
                          {
%>

<html>
<head>
<title>Administraci&oacute;n del módulo: <%=Session("admcursoName")%> </title>
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
     Response.Redirect("../errorpage.asp?tipo=Error&short=" +  DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);     
    }

 }
 else
   Response.Redirect("../errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>

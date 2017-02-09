<%@ Language=Jscript%>
<%
Response.Expires = 0;
%>

<!-- #include file='../../js/user.inc' -->
<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

%>

<html>
<head>
<title>Administraci&oacute;n</title>
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
   Response.Redirect("../errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>

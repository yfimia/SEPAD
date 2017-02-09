<%@ Language=jScript %>
<%
Response.Expires = 0;
%>
<!-- #include file="../js/Adolibrary.inc" --> 
<!-- #include file='../js/user.inc' -->
<%
if (Request.QueryString.Item("uid") + "" == Session("uid"))
{    
	var uid = Request.QueryString.Item("uid") + "";       
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<frameset rows="*" cols="140,*" framespacing="0" frameborder="NO" border="0">
  <frame src="folders.asp?uid=<%=uid%>&fldr=INBOX" name="leftFrame" scrolling="NO" noresize>
  <frame src="inbox.asp?uid=<%=uid%>" name="mainFrame">
</frameset>
<noframes><body>
</body></noframes>
</html>
<%
}
else
{
	Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
}
%>
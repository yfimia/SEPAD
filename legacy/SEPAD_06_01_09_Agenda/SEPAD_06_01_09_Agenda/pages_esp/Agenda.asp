<%@ Language=jScript %>
<%
Response.Expires = 0;
%>
<!-- #include file="../js/Adolibrary.inc" --> 
<!-- #include file='../js/user.inc' -->
<!-- #include file='../js/date_funtions.inc' -->

<%				
if (Request.QueryString.Item("uid") + "" == Session("uid"))
{    
	var uid = Request.QueryString.Item("uid") + "";  
	
	var fAgendaFecha = new Date();
	var sAgendaFecha = ""; 		

	if (Request.QueryString.Item("sAgendaFecha").Count != 0)
	{
		sAgendaFecha = Request.QueryString.Item("sAgendaFecha") + "";
		Session("sAgendaFecha") = sAgendaFecha;
	}
	else 
	{
		var tmpAgFecha = new String (Session("sAgendaFecha"));
		if (tmpAgFecha.length != 8) 
		{
			sAgendaFecha = DateToStringShort (fAgendaFecha);			
			Session("sAgendaFecha") = sAgendaFecha;	
		}	
		else 
		{
			sAgendaFecha = Session("sAgendaFecha");
		}
	}

%>
<html>
<head>
<title>Agenda de Actividades</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<frameset rows="*" cols="155,*" framespacing="0" frameborder="NO" border="0">
  <frame src="AgendaMes.asp?uid=<%=uid%>" name="leftFrame" scrolling="NO" noresize>
  <frame src="AgendaDia.asp?uid=<%=uid%>" name="mainFrame">
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
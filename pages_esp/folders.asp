<%@ Language=jScript %>
<%  
 Response.Expires = -1;
%>
<!-- #include file='../js/user.inc' -->
<%
if (Request.QueryString.Item("uid") + "" == Session("uid"))
{    
	var uid = Request.QueryString.Item("uid") + "";       
	var fldr = Request.QueryString.Item("fldr") + "";
	if (!(fldr.toLowerCase() == "inbox" || fldr.toLowerCase() == "sent"))
	{    
		fldr = "INBOX";
	}
%>	
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<script src="../js/CheckBoxes.js" language="JavaScript"></script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>
</head>

<body class="ToolBarBodyNoBorders">
<table width="130"  border="0" align="center" cellpadding="2" cellspacing="0">
  <tr>
    <td width="2">&nbsp;</td>
    <td width="162">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2" class="SubTitle"><table  border="0" cellpadding="2" cellspacing="0" class="SubTitle">
      <tr>
        <td>Carpetas</td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td width="2">&nbsp;</td>
    <td nowrap><a href="#" class="ToolLink" onClick="MM_goToURL('parent.frames[\'mainFrame\']','inbox.asp?uid=<%=uid%>');MM_goToURL('parent.frames[\'leftFrame\']','folders.asp?uid=<%=uid%>&fldr=INBOX');return document.MM_returnValue">Bandeja de entrada</a></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td nowrap><a href="#" class="ToolLink" onClick="MM_goToURL('parent.frames[\'mainFrame\']','sent_messages.asp?uid=<%=uid%>');MM_goToURL('parent.frames[\'leftFrame\']','folders.asp?uid=<%=uid%>&fldr=SENT');return document.MM_returnValue">Mensajes Enviados </a></td>
  </tr>
</table>
</body>
</html>
<%
}
else
{
	Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
}
%>
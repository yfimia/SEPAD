<%@ Language=jScript %>
<%  
 Response.Expires = -1;
%>
<!-- #include file='../js/user.inc' -->
<!-- #include file='../js/date_funtions.inc' -->
<!-- #include file='../js/adolibrary.inc' -->
<%
if (Request.QueryString.Item("uid") + "" == Session("uid"))
{    
	var uid = Request.QueryString.Item("uid") + "";       
%>	

<%
function _put0(aI)
{
	return ((aI < 10) ? ("0" + aI ) : ("" + aI ));
}
function _selected_option(aExpr)
{
	return (aExpr ? " SELECTED " : "" );
}

	var sAgendaFecha = new String(Session("sAgendaFecha"));
	if (Request.Form.Count > 0 )
	{ 
		var anho_act = Request.Form("anho")* 1;
		var mes = Request.Form("mes")* 1;
	}
	else
	{ 
		var anho_act = sAgendaFecha.substring (0,4) * 1;
		var mes = sAgendaFecha.substring (4,6)* 1;
	}
	
	//dbwhere = " AND (Usuarios.[ID]=" + asUser +	") AND (CONVERT(char(10), DateBegin, 112) BETWEEN ('" + asDate_i + "') AND ('" +  asDate_f +"'))";

%>

<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/EvalResult.css" type="text/css">
<script src="../js/CheckBoxes.js" language="JavaScript"></script>
</head>
<script language="jscript">
function _send()
{  
	form_fecha.submit();      
} 
</script>
<body class="ToolBarBodyNoBorders">
<table width="153"  border="0" align="center" cellpadding="0" cellspacing="0">
	<tr> 
	  <td class="HeaderTable" align=right><b>&nbsp;</b></td>
	</tr>
    <tr>
     <td align = "middle"> 
     <table class="MessageTR1" width= "100%">
      <tr>
       <td class="MessageTR">
		<form name = form_fecha  id = form_fecha METHOD=POST action="AgendaMes.asp?uid=<%=uid%>" >
		    <select  id="mes" name="mes" onChange = "_send()">
            <%
            for ( var i = 0; i < 12; i++) 
            {%>
				<option <%=_selected_option(mes == _put0(i+1))%> value = '<%=_put0(i+1)%>'><%=MONTH_NAME[i]%> </option>
            <%}%>
            </select><!--
        </td>
	  </tr>
	  <tr >
	    <td class="MessageTR">
	    --><select id="anho" name="anho" onChange = "_send()">
			<%
			for ( var i = (anho_act - 3); (i < (anho_act + 4));(i++)) 
            {%>
				<option <%=_selected_option(anho_act == i)%> value = '<%=i%>'><%=i%> </option>
            <% } %>
            </select>
         <form id=form1 name=form1>
       </td></tr>
       </table>
       </td>
      </tr>
  <tr>
    <td align=center class="MessageTR1">
		<% 
		AgendaBuildCalendario (mes,anho_act); 
		
		%>
    </td>
   </tr>

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
<%@ Language=jScript %>
<%  
 Response.Expires = -1;
%>
<!-- #include file="inc/AgendaDia.inc" --> 
<!-- #include file="../js/Adolibrary.inc" --> 
<!-- #include file='../js/user.inc' -->
<!-- #include file='../js/date_funtions.inc' -->
<%


if (Request.QueryString.Item("uid") + "" == Session("uid"))
{    
	var uid = Request.QueryString.Item("uid") + "";       
	var userID = Session ("userID");
	
	var  filePath = Application("dataPath");   
    var  oConn = Server.CreateObject("ADODB.Connection");
    oConn.Open(filePath);
	
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
    // Build Query

	if (Request.Form.Count > 0)  
    {
		var SQL = "DELETE FROM AgendaEventos WHERE ";
		
		var Tot = Request.Form("Count");
		var i;
            
        //Busqueda de los mensajes a borrar      
        for (i=1; i <= Tot; i++)                                         
			if ((Request.Form("INP" + i).Count > 0) && (Request.Form("ID" + i).Count > 0))
            {
				SQL = SQL + "(ID=" + Request.Form("ID" + i) + ") OR ";
            }              

            SQL = SQL + "(1=0)";
            oConn.Execute (SQL);		 
	}
       //Response.Write(SQL);
   
function AgendaEventosDia(asUser, asDate, aoConn)
{	
	dbwhereToGroup = "(AgendaEventos.ToGroup > 0) ";
	dbwhereToSubGroup = "(AgendaEventos.ToSubGroup > 0) ";
	dbwhereToUsr = "(AgendaEventos.ToUsr > 0)";

	dbwhere = " AND (Usuarios.[ID]=" + asUser +
		") AND (('" + asDate + "'" + 
		" BETWEEN CONVERT(char(10), DateBegin, 112) AND " + 
		" CONVERT(char(10), DateEnd, 112)) OR EventType = 2) ";
		
	strFechas = ", CONVERT(char(5), DateBegin, 108) as sTimeBegin " + 
				", CONVERT(char(5), DateEnd, 108) as sTimeEnd " + 
				", CONVERT(char(8), DateBegin, 112) as sDateBegin " + 
				", CONVERT(char(8), DateEnd, 112) as sdateEnd " ; 
    
    strsql = "SELECT AgendaEventos.*, Usuarios.[ID] as [UserID] " + 
				strFechas +
				"FROM Grupos_de_Usuarios INNER JOIN " + 
				"Usuarios ON Grupos_de_Usuarios.[User] = Usuarios.ID RIGHT OUTER JOIN " +
				"AgendaEventos ON Grupos_de_Usuarios.[Group] = AgendaEventos.ToGroup " + 
				"WHERE " 
					+ dbwhereToGroup + dbwhere + 
				"UNION ( " + 
				"SELECT AgendaEventos.*, Usuarios.[ID] as [UserID] " + 
				strFechas +
				"FROM AgendaEventos INNER JOIN " +
				"Usuarios ON AgendaEventos.ToUsr = Usuarios.ID " +
				"WHERE " 
					+ dbwhereToUsr + dbwhere + 
				") " +
				"UNION ( " + 
				"SELECT AgendaEventos.*, Usuarios.[ID] as [UserID] " + 
				strFechas +
				"FROM UsuariosSubGrupo INNER JOIN " +
				"Usuarios ON UsuariosSubGrupo.Usuario = Usuarios.ID INNER JOIN " +
				"AgendaEventos ON UsuariosSubGrupo.Subgrupo = AgendaEventos.ToSubGroup " +
				"WHERE " 
					+ dbwhereToSubGroup + dbwhere + 
				") " +				
				"UNION ( " + 
				"SELECT AgendaEventos.*, Usuarios.[ID] as [UserID] " + 
				strFechas +
				"FROM AgendaEventos INNER JOIN " +
				"Usuarios ON AgendaEventos.[OwnerId] = Usuarios.ID " +
				"WHERE (1=1) " 
					+ dbwhere + 
				") " +					
				"ORDER BY AgendaEventos.DateBegin"
    
    var  filePath = Application("dataPath");   
    var  oRec     = Server.CreateObject("ADODB.Recordset");

    oRec.Open(strsql, aoConn, 1, 2);
	return oRec;
}   
//Response.Write(sAgendaFecha+ "|");
oRec = AgendaEventosDia(userID, sAgendaFecha, oConn);

%>
<html>
<head>
<title>Actividades del Dia</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/EvalResult.css" type="text/css">
<script type="text/javascript" language="javascript" src="../js/Windows.js"></script>
<script src="../js/CheckBoxes.js" language="JavaScript"></script>

</head>

<body bgcolor="#FFFFFF" text="#000000" style="overflow-y: auto;" onLoad="parent.frames['leftFrame'].location = 'AgendaMes.asp?uid=<%=uid%>';" >
<form id="Messages" method="POST" action="inbox.asp?uid=<%=uid%>">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<!--
  <tr> 
    <td class="MessageBarTD"> 
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td><a href="javascript:abreVentana('AgendaEvento',400,330,'AgendaEvento.asp?uid=<%=uid%>', 'no', 'no')" class="ToolLink">&nbsp;Nueva&nbsp;Evento&nbsp;</a></td>
          <td>|</td>
          <td><a href="javascript:onDeleteMessage()" class="ToolLink">&nbsp;Borrar&nbsp;Evento(s)&nbsp;</a></td>
          <td>|</td>
          <td><a href="#" class="ToolLink" onClick="CheckAll(GetObjectCollection( document, 'FORM'))">&nbsp;Seleccionar&nbsp;todas&nbsp;las&nbsp;Eventos&nbsp;</a></td>
        </tr>
      </table>
    </td>
  </tr>
-->
  <tr> 
    <td>
	  <table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td class="HeaderTable" align=center><b><%=sDateShortTosDateLong(sAgendaFecha)%></b></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<%
if (! oRec.EOF)
{
%>
      <table width="100%" cellpadding="0" cellspacing="1" border="0">
        <tr class="HeaderTable"  valign="middle"> 
          <td class="ToolBar" width="1%">&nbsp;</td>
          <td class="ToolBar" width="1%">&nbsp;</td>
          <td class="ToolBar" width="1%">&nbsp;</td>
          <td class="ToolBar" valign = "top" width="10%"><b><%=Comienzo%></b></td>
          <td class="ToolBar" valign = "top" width="10%"><b><%=Final%></b></td>
          <td class="ToolBar" width="77%"> <b><%=Actividad%></b> </a></td>
        </tr>
<%
}

var Alert = new Array("","!"); 
var ChechInput = "";
var trclass = 'MessageTR';
//
if (! oRec.EOF)
{
  recCount = 0;
  recActual = 0;
  while (! oRec.EOF) 
  {
	recCount = recCount + 1;
	bold = '';
    if (oRec.Fields("Readed").Value == false) 
    {
      //bold = 'Style="font-weight:bolder"';
    }
	ChechInput = ((oRec.Fields ("OwnerID").Value == userID) ? '<input type="checkbox"  name="INP' + recCount + '" id="INP' + recCount + '"> <INPUT TYPE=hidden NAME="ID' + recCount + '" VALUE="' + oRec.Fields("ID").Value  + '">' : '&nbsp;');

   	var Vinculo = '<a href="#" onclick="onDblclick(' + oRec.Fields ("ID").Value +')">';
   	var Actividad = "<b>" + oRec.Fields("Subject").Value + "</b></a><br>" + oRec.Fields("Content").Value;
   	
   	var TimeBegin = new String (oRec.Fields("sTimeBegin").Value  + "");
   	TimeBegin = ( TimeBegin + "" == "null" ? "": TimeBegin);

   	var TimeEnd = new String (oRec.Fields("sTimeEnd").Value  + "");
    TimeEnd = ( TimeEnd + "" == "null" ? "" : TimeEnd);

   	var DateBegin = new String (oRec.Fields("sDateBegin").Value  + "");
   	if (DateBegin == sAgendaFecha || DateBegin == "null")
   		DateBegin = ""
   	else 
   	{ 
   		DateBegin = "<a href = 'AgendaDia.asp?sAgendaFecha="+DateBegin+"&uid="+uid+"'>"+sShortDateToPrettyDate(DateBegin) + "</a><br>";
   		TimeBegin = "";
   	}

   	var DateEnd = new String (oRec.Fields("sDateEnd").Value  + "");
   	if (DateEnd == sAgendaFecha || DateEnd == "null")
   		DateEnd = ""
   	else
   	{
   		DateEnd = "<a href = 'AgendaDia.asp?sAgendaFecha="+DateEnd+"&uid="+uid+"'>"+sShortDateToPrettyDate(DateEnd) + "</a><br>";
   		TimeEnd = "";
	}
   	
   	if (trclass == 'MessageTR') 
   		trclass = 'MessageTR1' 
   	else 
   		trclass = 'MessageTR';
		
%>

        <tr  <%=bold%> class="<%=trclass%>"  valign="middle"> 
          <td align=center width="1%"><%=ChechInput %></td>
          <td align=center width="1%" class="PriorTD"><%=Alert[oRec.Fields("Priority").Value] %></td>
          <td align=center width="1%">&nbsp; 		</td>
          <td align=right valign = "top" width="10%"><%=DateBegin%><%=TimeBegin%> </td>
          <td align=right valign = "top" width="10%"><%=DateEnd%><%=TimeEnd%> </td>
          <td valign = "top" width="77%"><%=Vinculo%> <%=Actividad %></a></td>
        </tr>
<%                	  
          oRec.Move(1);}
%>
  <input type=hidden name="Count" value="<%=recCount%>">
     </table>
    </td>
  </tr>
</table>
<%
} else {
%>
<br>
<p class="txtregy" align="center"><strong><%=NoActividades%></strong></p>
<br>
<% 
 }
%>
<%
// Close recordset and connection
oRec.Close();
oConn.Close();
%>
<table border="0" cellspacing="0" cellpadding="0" class="ToolBar" width="100%">
  <tr>
    <td>
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td><a href="javascript:abreVentana('AgendaEvento',400,330,'AgendaEvento.asp?uid=<%=uid%>', 'no', 'no')" class="ToolLink"><%=NuevaActividad%></a></td>
          <td>|</td>
          <td><a href="javascript:onDeleteMessage()" class="ToolLink"><%=BorrarActividades%></a></td>
          <td>|</td>
          <td><a href="#" class="ToolLink" onClick="CheckAll(GetObjectCollection( document, 'FORM'))"><%=Seleccionar%></a></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<input type=hidden name='flag' id='flag' value=1>  
</form>
<script language="javascript" >
<!--
   
function onDblclick( id ) {    
     event.srcElement.parentElement.parentElement.style.fontWeight  = "lighter";
     abreVentana('',400,330,'AgendaEvento.asp?uid=<%=uid%>&ID=' + id , 'no', 'no')
}  

    
function onDeleteMessage() {      
     	Messages.all("flag").value = 0;
     	Messages.action = "AgendaDia.asp?uid=<%=uid%>";
     	Messages.submit();
}
    
function onDeleteAll() {    
    	Messages.all("flag").value = 1;
    	Messages.submit();
}    
-->          
</script>

</body>
</html>
<%
 }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>

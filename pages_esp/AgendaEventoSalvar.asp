<%@ Language=JScript %>
<%
  Response.Expires = -1;
%>
<!-- #include file='../js/user.inc' -->
<!-- #include file='../js/adolibrary.inc' -->

<%
if (Request.QueryString.Item("uid") + "" == Session("uid"))
{    
    var uid = Request.QueryString.Item("uid") + "";      
      
    eventoID = Request.Form("ID") + "";
    OwnerName = Request.Form("OwnerName") + "";
    OwnerId = Request.Form("OwnerId") + "";
    ToName = Request.Form("ToName") + "";
    
    var ToDestiny = new String(Request.Form("Select") +"");
    var ToDestinyValue = new String(Request.Form("Select") +"");
    ToDestinyValue = ToDestiny.substring (ToDestiny.search ("_")+1, ToDestiny.length );
    ToDestiny = ToDestiny.substring (0, ToDestiny.search ("_"));
        
    Subject = Request.Form("Subject") + "";
    Content = Request.Form("Content") + "";
    Priority = Request.Form("Priority") + "";
    DateBegin = Request.Form("DateBegin")+ "";
    DateBegin = (DateBegin == "null" ? DateBegin : "CONVERT(datetime, '"+DateBegin+"',103)");
    DateEnd = Request.Form("DateEnd")+ "";
    DateEnd = (DateEnd == "null" ? DateEnd : "CONVERT(datetime, '"+DateEnd+"',103)");
    EventType = Request.Form("EventType") + "";
    
	var oConn = Server.CreateObject("ADODB.Connection");  
	var oRec = Server.CreateObject("ADODB.Recordset");  
		
	var filePath = Application("filePath");
	oConn = MakeConnection( oConn, filePath );
    
    if (eventoID == "0")
    {
		// Consulta de Insercion Si ID = 0;
		sSql1 = "";
		sSql = "INSERT INTO AgendaEventos " +
				"(OwnerName, OwnerId, ToName, "+ToDestiny+", " +
				"Subject, Content, Priority, " + 
				"DateBegin, DateEnd, EventType) " +
				"VALUES " +
				"('"+OwnerName+"', "+OwnerId+", '"+ToName+"', "+ToDestinyValue+
				", '"+Subject+"', '"+Content+"', "+Priority+", " +
				""+DateBegin+", "+DateEnd+", "+EventType+") ";
		Query(sSql, oRec, oConn);
	}
	else
	{
		// Consulta de Modificacaión Si ID != 0;
		sSql1 = "UPDATE AgendaEventos " + 
				"SET "+ 
				" ToUsr = 0, ToModulo = 0, ToCourse = 0"+
				", ToGroup = 0, ToSubGroup = 0 " + 
				" WHERE AgendaEventos.[id] ="+eventoID;

		sSql = "UPDATE AgendaEventos " + 
				"SET "+ 
				"OwnerName ='"+OwnerName+"', OwnerId ="+OwnerId+", ToName ='"+ToName+
				"', " + ToDestiny + "="+ToDestinyValue+
				",Subject='"+Subject+"', Content='"+Content +
				"', Priority ="+Priority+", DateBegin ="+DateBegin+
				", DateEnd ="+DateEnd+", EventType = "+EventType+
				" WHERE AgendaEventos.[id] ="+eventoID;
		
		Query(sSql1, oRec, oConn);
		Query(sSql, oRec, oConn);
	}
	//DestroyAdoObjects( oConn, oRec );
%>
<BODY onLoad = "close();">
<script language=JSCRIPT>
  //alert(opener);
  //OJO
  if  ((opener  != null) && (opener + "" != "undefined") && (opener + ""  != "null") && (opener  != 0))
    opener.location = "AgendaDia.asp?uid=<%=uid%>";
	//opener.parent.frames["leftFrame"].location = "AgendaMes.asp?uid=<%=uid%>";
  close();
</script>
<HTML>
</BODY>
</HTML>
<%
}
else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
  //Response.Write(Request.QueryString.Item("uid"));
%>
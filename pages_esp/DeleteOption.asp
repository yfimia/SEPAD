<%@ Language=jScript %>

<%Response.Expires = -1%>

<!-- #include file='../js/adolibrary.inc' -->
<!-- #include file='../js/user.inc' -->

<html>
<%
  var oid = Request.QueryString.Item("oid");
  
  var filePath = Application("dataPath");
  var oCon = Server.CreateObject("ADODB.Connection");
  var oRec = Server.CreateObject("ADODB.Recordset");
  
  oCon.Open(filePath);
  Sql = "DELETE FROM Enlaces WHERE (ID=" + oid + ")";
  oRec.Open(Sql, oCon,3,3);
  
  oCon.Close();
  
  Response.Redirect("ManipuladorOpciones.asp");
%>
</html>
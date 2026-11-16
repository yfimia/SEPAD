<%@ Language=jScript %>

<%
  Response.Expires = -1
  
  var uid = Request.QueryString.Item("uid");
  
%>

<!-- #include file='../js/adolibrary.inc' -->
<!-- #include file='../js/user.inc' -->

<html>
<%
  var Caption = Request.Form("LinkText");
  var Url = Request.Form("LinkUrl");
  var oid = Request.QueryString.Item("oid");
  
  var filePath = Application("dataPath");
  var oCon = Server.CreateObject("ADODB.Connection");
  var oRec = Server.CreateObject("ADODB.Recordset");
  
  oCon.Open(filePath);
  Sql = "SELECT * FROM Enlaces WHERE (ID=" + oid + ")";
  oRec.Open(Sql, oCon,3,3);
  
  oRec.Fields.Item("Caption").Value = Caption;
  oRec.Fields.Item("Url").Value = Url;
  oRec.Update();
  
  oRec.Close();
  oCon.Close();
  
  Response.Redirect("ManipuladorOpciones.asp?uid=" + uid);
%>
</html>

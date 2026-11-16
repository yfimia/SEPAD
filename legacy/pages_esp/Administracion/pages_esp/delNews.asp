<%@language=jscript%>
<%
   Response.Expires = -1;
   IDNews = Request.QueryString("nid");
%>
<!-- #include file='../js/adolibrary.inc' -->
<!-- #include file='../js/user.inc' -->
<html>
<%
  var filePath = Application('dataPath');
  
  var oConn;
  var oRec;

  oConn = Server.CreateObject("ADODB.Connection");  
  oRec = Server.CreateObject("ADODB.Recordset");  
  oConn.open(filePath);
  Sql = "DELETE " +
        "FROM Noticias " +
  	    "WHERE (ID=" + IDNews + ")";
  oRec.Open(Sql, oConn,3,3);
  oConn.Close();
  Response.Redirect("news.asp?uid=" + Session("uid"));
%>
</html>
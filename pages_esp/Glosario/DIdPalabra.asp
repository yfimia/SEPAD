<%@language=jscript%>
<%
  
  var IC      = Request.QueryString.Item("IdCurso") * 1;
  var Let     = Request.QueryString.Item("Letra");
  var Word    = Request.QueryString.Item("Palabra");
  var CName   = Request.QueryString.Item("courseName");
  
  var oRec, oConn;
  
  oConn = Server.CreateObject("ADODB.Connection");
  oRec = Server.CreateObject("ADODB.Recordset");
  
  var Sql = "SELECT ID FROM Glosarios WHERE Palabra='" + Word + "'";
  
  oConn.Open(Application("filePath"));
  oRec.Open(Sql, oConn,3,3);
  
  var Idp = oRec.Fields.Item("ID").value;
  
  oRec.Close();
  oConn.Close();


  Response.Redirect("glosario.asp?IdCurso=" + IC + "&Letra=" + Let + "&IDP=" + Idp + "&courseName=" + CName);
  
%>
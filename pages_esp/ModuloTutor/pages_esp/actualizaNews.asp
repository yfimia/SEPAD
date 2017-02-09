<%@ Language=JavaScript%>
<%
  Response.Expires = -1;
%>
<%
  newsId = Request.QueryString("nid") + "";
  Titulo = Request.Form("titulo");
  Cuerpo = Request.Form("Cuerpo");
  Imagen = Request.Form("imagen");
  URL    = Request.Form("url");
  uid    = Request.QueryString.Item("uid");
  
  var filePath = Application('dataPath');
  
  var oConn;
  var oRec;
  
  oConn = Server.CreateObject("ADODB.Connection");
  oRec = Server.CreateObject("ADODB.Recordset");
  
  Consulta = "SELECT * FROM Noticias WHERE (ID=" + newsId + ")";
  oConn.open(filePath);
  oRec.open(Consulta, oConn,3,3);
  
  oRec.Fields.Item("Titulo").value = Titulo;
  oRec.Fields.Item("Cuerpo").value = Cuerpo;
  oRec.Fields.Item("Imagen").value = Imagen;
  oRec.Fields.Item("Url").value = URL;
  oRec.Update();
  oRec.Close();
  oConn.Close();
  Response.Redirect("news.asp?uid=" + uid);

%>

<html>
</html>

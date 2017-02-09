<%@language=jscript%>

<!-- #include file="../../js/adoLibrary.inc" -->
<!-- #include file='../../js/user.inc' -->

<%
  var CRecords = Request.Form.Item("CantidadRecord");
  
  var oConn;
  var oComm  = Server.CreateObject("ADODB.Command");

  var filePath = Application("dataPath");
  
  oConn = MakeConnection( oConn, filePath );
  oComm.ActiveConnection = oConn;
  
  for (var i=1;i<=CRecords+1;i++) {
      if (Request.Form("CheckBox" + i) + "" == "on") {
         //   Eliminar Palabra
         var IdP = Request.Form("PalabraIDEliminar" + i);
         //Response.Write(IdP + "<br>");
         oComm.CommandText = "DELETE FROM Glosarios WHERE (Id=" + IdP + ")";
         oComm.Execute();
      }   
      
  }
  oConn.Close;
  Response.Redirect("ModificaEliminaPalabra.asp");
%>
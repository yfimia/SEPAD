<%@language=jscript%>
<%
    var idp      = Request.QueryString.Item("idp");
    var PWord    = Request.Form.Item("TextPal");
    var SSignif  = Request.Form.Item("TextSignif");
    var idCourse = Session("admcurso");

    
  var oRec, oConn, oRecExistPalabra;
  
  oConn            = Server.CreateObject("ADODB.Connection");
  oRec             = Server.CreateObject("ADODB.Recordset");
  oRecExistPalabra = Server.CreateObject("ADODB.Recordset");
  
  var Sql = "SELECT * FROM Glosarios WHERE (Id=" + idp + ")";
  
  var filePath = Application("dataPath");
  oConn.Open( filePath );
    
    oRec.Open(Sql, oConn,3,3);
    
  if ((PWord != "") && (SSignif != "")) {
    oRec.Fields.Item("Palabra").value = PWord;
    oRec.Fields.Item("Signif").value = SSignif;
    oRec.Fields.Item("IdCurso").value = idCourse;
    oRec.Update();
  }  
  
    oRec.Close();

  oConn.Close();
  Response.Redirect("ModificaEliminaPalabra.asp");
  
%>

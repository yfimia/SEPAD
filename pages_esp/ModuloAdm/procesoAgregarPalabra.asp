<%@language=jscript%>
<%
    var PWord    = Request.Form.Item("TextPal");
    var SSignif  = Request.Form.Item("TextSignif");
    var idCourse = Session("admcurso");

    
  var oRec, oConn, oRecExistPalabra;
  
  oConn            = Server.CreateObject("ADODB.Connection");
  oRec             = Server.CreateObject("ADODB.Recordset");
  oRecExistPalabra = Server.CreateObject("ADODB.Recordset");
  
  var Sql = "SELECT * FROM Glosarios";
  var SqlExistPalabra = "SELECT Palabra FROM Glosarios WHERE Palabra='" + PWord + "'";
  
  var filePath = Application("dataPath");
  oConn.Open( filePath );
    
  oRecExistPalabra.Open(SqlExistPalabra, oConn,3,3);
  if ((oRecExistPalabra.EOF) && (PWord != "")) {
  
    oRec.Open(Sql, oConn,3,3);
    
  if ((PWord != "") && (SSignif != "")) {
    oRec.AddNew();
    oRec.Fields.Item("Palabra").value = PWord;
    oRec.Fields.Item("Signif").value = SSignif;
    oRec.Fields.Item("IdCurso").value = idCourse;
    oRec.Update();
  }  
  
    oRec.Close();

  }
  oRecExistPalabra.Close();
  oConn.Close();
  Response.Redirect("AgregarPalabra.asp");
  
%>

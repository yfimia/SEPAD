<%@language=jscript%>
<%
    var idp      = Request.QueryString.Item("idp");
    var Pregunta    = Request.Form.Item("question");
    var Respuesta  = Request.Form.Item("response");
    var IdCourse = Session("admcurso");

    
  var oRec, oConn, oRecExistPregunta;
  
  oConn            = Server.CreateObject("ADODB.Connection");
  oRec             = Server.CreateObject("ADODB.Recordset");
  oRecExistPalabra = Server.CreateObject("ADODB.Recordset");
  
  var Sql = "SELECT * FROM faq WHERE (Id=" + idp + ")";
  
  var filePath = Application("dataPath");
  oConn.Open( filePath );
    
    oRec.Open(Sql, oConn,3,3);
    
  if ((Pregunta != "") && (Respuesta != "")) {
    oRec.Fields.Item("question").value = Pregunta;
    oRec.Fields.Item("response").value = Respuesta;
    oRec.Fields.Item("course").value = IdCourse;
    oRec.Update();
  }  
  
    oRec.Close();

  oConn.Close();
  Response.Redirect("ModificaEliminaPregunta.asp");
  
%>

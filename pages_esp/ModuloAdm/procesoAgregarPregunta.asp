<%@language=jscript%>
<%
var Pregunta    = Request.Form.Item("question");
var Respuesta  = Request.Form.Item("response");
var IdCourse = Session("admcurso");
%>

<!-- #include file="../../js/user.inc" --> 

<%
    
var oRec, oConn, oRecExistPregunta;
  
oConn            = Server.CreateObject("ADODB.Connection");
oRec             = Server.CreateObject("ADODB.Recordset");
oRecExistPregunta = Server.CreateObject("ADODB.Recordset");
  
var Sql = "SELECT * FROM faq";
var SqlExistPregunta = "SELECT question FROM faq WHERE question='" + Pregunta + "' AND (course=" + IdCourse + ")";
  
var filePath = Application("dataPath");
oConn.Open( filePath );

oRecExistPregunta.Open(SqlExistPregunta, oConn,3,3);
if ((oRecExistPregunta.EOF) && (Pregunta != "")) {
  
  oRec.Open(Sql, oConn,3,3);
    
  if ((Pregunta != "") && (Respuesta != "")) {
    oRec.AddNew();
    oRec.Fields.Item("question").value = Pregunta;
    oRec.Fields.Item("response").value = Respuesta;
    oRec.Fields.Item("course").value = IdCourse;
    oRec.Update();
  }  
  
  oRec.Close();
  oRecExistPregunta.Close();
  oConn.Close();
  Response.Redirect("AgregarPregunta.asp");

}
else
{
  oRecExistPregunta.Close();
  oConn.Close();
  
  Response.Redirect("errorpage.asp?tipo=Error&short=" + EXIST_QUESTION_SHORT  + "&desc=" + EXIST_QUESTION_TEXT);
}
  
%>

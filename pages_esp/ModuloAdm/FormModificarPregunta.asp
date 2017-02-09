<%@language=jscript%>
<%
  Response.Expires = -1;
  var idPregunta = Request.QueryString.Item("idp");
  
  // Buscando datos de idPregunta;
  var oConn = Server.CreateObject("ADODB.Connection");;
  var oRec1 = Server.CreateObject("ADODB.Recordset");
  
  var filePath = Application("dataPath");
  
  oConn.Open( filePath );
  
  var Sql1 = "SELECT * FROM faq WHERE (Id=" + idPregunta + ") AND (course=" + Session("admcurso") + ")";
  oRec1.Open(Sql1, oConn,3,3);
  
  if (!oRec1.EOF) {
      var MPregunta = oRec1.Fields.Item("question");
      var MRespuesta  = oRec1.Fields.Item("response");
      var MIdCurso = oRec1.Fields.Item("course");
  } else {
      var MPregunta = "";
      var MRespuesta  = "";
      var MIdCurso = 0;
  }    
  
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link rel="stylesheet" href="css/EvalResult.css" type="text/css">
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/Main.css" type="text/css">
<!--link rel="stylesheet" href="css/Main.css" type="text/css"-->
<link href="css/Aqua/Glosary.css" rel="stylesheet" type="text/css" />
</head>

<body>
<form id="formDatosModificar" name="formDatosModificar" action="procesoModificarPregunta.asp?idp=<%=idPregunta%>" method="post">
<table width="100%" border="0" cellspacing="0" cellpadding="2">
  <tr> 
    <td class="ToolBar"><table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td><img height="54" src="../../images/<%=Session("skin")%>/EvalResult.gif" width="80"></td>
          <td class="HeaderTable">Modificar Pregunta</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
    <table border=0 cellpadding=0 cellspacing=0 width="100%">
       <tr class="MessageTR">
        <td align="center" width="1%">Pregunta:</td><td><input type=text id="question" name="question" size=60 value="<%=MPregunta%>"></td>
       </tr>
       
       <tr class="MessageTR1">
        <td align="center" width="1%" valign="top">Respuesta:</td><td><textarea id="response" name="response" cols=60 rows=10><%=MRespuesta%></textarea></td>
       </tr>
       <tr>
        <td colspan=2 align="center"><input type=submit value="Modificar"></td>
       </tr>
 </table>   
    
</body>       
</html>
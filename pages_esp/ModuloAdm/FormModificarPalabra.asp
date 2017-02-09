<%@language=jscript%>
<%
  Response.Expires = -1;
  var idPalabra = Request.QueryString.Item("idp");
  
  // Buscando datos de idpalabra;
  var oConn = Server.CreateObject("ADODB.Connection");;
  var oRec1 = Server.CreateObject("ADODB.Recordset");
  
  var filePath = Application("dataPath");
  
  oConn.Open( filePath );
  
  var Sql1 = "SELECT * FROM Glosarios WHERE (Id=" + idPalabra + ") AND (IdCurso=" + Session("admcurso") + ")";
  oRec1.Open(Sql1, oConn,3,3);
  
  if (!oRec1.EOF) {
      var MPalabra = oRec1.Fields.Item("Palabra");
      var MSignif  = oRec1.Fields.Item("Signif");
      var MIdCurso = oRec1.Fields.Item("IdCurso");
  } else {
      var MPalabra = "";
      var MSignif  = "";
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
 <table border=0 cellpadding=0 cellspacing=0 width="100%">
   <form id="formDatosModificar" name="formDatosModificar" action="procesoModificarPalabra.asp?idp=<%=idPalabra%>" method="post">
       <tr class="MessageTR">
        <td align="center" width="1%">Palabra:</td><td><input type=text id="TextPal" name="TextPal" size=60 value="<%=MPalabra%>"></td>
       </tr>
       
       <tr class="MessageTR1">
        <td align="center" width="1%" valign="top">Significado:</td><td><textarea id="TextSignif" name="TextSignif" cols=60 rows=10><%=MSignif%></textarea></td>
       </tr>
       <tr>
        <td colspan=2 align="center"><input type=submit value="Modificar"></td>
       </tr>
 </table>      
</body>       
</html>
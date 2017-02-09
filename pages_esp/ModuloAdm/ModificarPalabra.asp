<%@language=jscript%>
<!-- #include file="../../js/adoLibrary.inc" -->
<!-- #include file='../../js/user.inc' -->
<%

  Response.Expires = -1;
  var idPalabra = Request.QueryString.Item("idp");
  
  // Buscando datos de idpalabra;
  var oConn = Server.CreateObject("ADODB.Connection");;
  var oRec1 = Server.CreateObject("ADODB.Recordset");
  
  var filePath = Application("dataPath");
  
  oConn.Open( filePath );
  
  var Sql1 = "SELECT * FROM Glosarios WHERE (Id=" + idPalabra + ")";
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
<?xml version="1.0" encoding="iso-8859-1"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
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
<a name="PageTop" id="PageTop"></a> 
<table width="100%" border="0" cellspacing="0" cellpadding="4">
  <tr> 
    <td class="ToolBar"><table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td><img height="54" src="../../images/<%=Session("skin")%>/GlosarioIMG.gif" width="80"></td>
          <td class="HeaderTable">Modificar Palabra</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
     <table border=0 cellpadding=0 cellspacing=0 width="100%">
       <tr>
        <td colspan=2>
     <table border=0 cellpadding=0 cellspacing=0 width="100%">
       <tr class="ToolBar">
        <td>Palabra</td><td>&nbsp;</td><td>Significado</td><td>&nbsp;</td><td>Modificar (M)</td>
       </tr>
           <%
              var oRec, oConn;
  
              oConn = Server.CreateObject("ADODB.Connection");
              oRec = Server.CreateObject("ADODB.Recordset");
  
              var Sql = "SELECT * FROM Glosarios";
              
              oConn.open(Application("filePath"));
              oRec.open(Sql, oConn,3,3);
              
              var CantRecord = oRec.RecordCount;
              var NameCheck = 1;
              var IsPar = "FALSE";
              var clase;
              
              while (!oRec.EOF) {
                if (IsPar) { clase="MessageTR"; } else { clase="MessageTR1"; }
               IsPar = !IsPar;
           %>
             <tr class="<%=clase%>">
              <td>
               <%=oRec.Fields.Item("Palabra")%>
              </td>
              <td><font color=white>|</font></td>
              <td>
               <%=oRec.Fields.Item("Signif")%>
              </td>
              <td><font color=white>|</font></td>
              <td align="center">
               <a href="ModificarPalabra.asp?idp=<%=oRec.Fields.Item("Id")%>">M</a>
              </td>
             </tr>
           <%
               oRec.move(1);
               NameCheck++;
             }
           %>
     </table>
        </td>
       </tr>
       <tr>
         <td colspan=4>&nbsp;</td>
       </tr>
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
   </form> 
  </tr>
</table>
<%
 oRec1.Close;
 oConn.Close;
%>
</body>
</html>

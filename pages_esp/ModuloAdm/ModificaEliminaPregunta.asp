<%@language=jscript%>
<%
  Response.Expires = -1;
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
          <td><img height="54" src="../../images/<%=Session("skin")%>/EvalResult.gif" width="80"></td>
          <td class="HeaderTable">Modificar/Eliminar Preguntas</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
   <form id="formDatosEliminar" name="formDatosEliminar" action="procesoEliminarPregunta.asp" method="post">
     <table border=0 cellpadding=0 cellspacing=0 width="100%">
       <tr class="ToolBar">
        <td width="1%">&nbsp;</td><td>Pregunta</td><td>&nbsp;</td><td>Respuesta</td><td>&nbsp;</td><td>Modificar&nbsp;(M)</td>
       </tr>
           <%
              var oRec, oConn;
  
              oConn = Server.CreateObject("ADODB.Connection");
              oRec = Server.CreateObject("ADODB.Recordset");
  
              var Sql = "SELECT * FROM faq where course = " + Session("admcurso") + " ORDER BY question";
              
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
               <input type=checkbox id="CheckBox<%=NameCheck%>" name="CheckBox<%=NameCheck%>">
               <input type=hidden id="PreguntaIDEliminar<%=NameCheck%>" name="PreguntaIDEliminar<%=NameCheck%>" value="<%=oRec.Fields.Item("ID")%>">
              </td>
              <td>
               <%=oRec.Fields.Item("question")%>
              </td>
              <td><font color=white>|</font></td>
              <td>
               <%=oRec.Fields.Item("response")%>
              </td>
              <td><font color=white>|</font></td>
              <td align="center">
               <a href="FormModificarPregunta.asp?idp=<%=oRec.Fields.Item("ID")%>">M</a>
              </td>
             </tr>
           <%
               oRec.move(1);
               NameCheck++;
             }
           %>
           <tr>
            <td colspan=6 align="right">
             <input type=hidden id="CantidadRecord" name="CantidadRecord" value=<%=CantRecord%>>
             <input type="submit" value="Eliminar">
            </td>
           </tr>
     </table>
   </form> 
  </tr>
</table>
</body>
</html>

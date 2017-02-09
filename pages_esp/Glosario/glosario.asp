<%@language=jscript%>
<%

  Response.Expires = -1;
  
  var uid = Request.QueryString.Item("uid");
  
  var IC     = Request.QueryString.Item("IdCurso");
  var courseName = Request.QueryString.Item("courseName");
  var Letter = Request.QueryString.Item("Letra");
  var IDPal  = Request.QueryString.Item("IDP");
  
%>
<?xml version="1.0" encoding="iso-8859-1"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>SEPAD - Glosario</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />

<link rel="stylesheet" href="css/EvalResult.css" type="text/css">
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/Main.css" type="text/css">
<!--link rel="stylesheet" href="css/Main.css" type="text/css"-->
<link href="css/Aqua/Glosary.css" rel="stylesheet" type="text/css" />
</head>

<body>
<table width="100%" border="0" cellspacing="0" cellpadding="4">
  <tr> 
    <td class="ToolBar"><table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td><img height="54" src="../../images/<%=Session("skin")%>/GlosarioIMG.gif" width="80"></td>
          <td class="HeaderTable">Glosario</td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td class="ToolBar"><strong>Curso:&nbsp;<%=courseName%></strong></td>
    
  </tr>
  <tr> 
    <td align="right" class="BottomedLinedCell"> <table border="0" cellspacing="0" cellpadding="2">
        <tr> 
          <td align="center"><a href="DLetra.asp?IdCurso=<%=IC%>&Letra=Todas&courseName=<%=courseName%>">Todas</a></td>
          <td align="center">&nbsp;</td>
          <td align="center"><a href="DLetra.asp?IdCurso=<%=IC%>&Letra=A&courseName=<%=courseName%>">A</a></td>
          <td align="center"><a href="DLetra.asp?IdCurso=<%=IC%>&Letra=B&courseName=<%=courseName%>">B</a></td>
          <td align="center"><a href="DLetra.asp?IdCurso=<%=IC%>&Letra=C&courseName=<%=courseName%>">C</a></td>
          <td align="center"><a href="DLetra.asp?IdCurso=<%=IC%>&Letra=D&courseName=<%=courseName%>">D</a></td>
          <td align="center"><a href="DLetra.asp?IdCurso=<%=IC%>&Letra=E&courseName=<%=courseName%>">E</a></td>
          <td align="center"><a href="DLetra.asp?IdCurso=<%=IC%>&Letra=F&courseName=<%=courseName%>">F</a></td>
          <td align="center"><a href="DLetra.asp?IdCurso=<%=IC%>&Letra=G&courseName=<%=courseName%>">G</a></td>
          <td align="center"><a href="DLetra.asp?IdCurso=<%=IC%>&Letra=H&courseName=<%=courseName%>">H</a></td>
          <td align="center"><a href="DLetra.asp?IdCurso=<%=IC%>&Letra=I&courseName=<%=courseName%>">I</a></td>
          <td align="center"><a href="DLetra.asp?IdCurso=<%=IC%>&Letra=J&courseName=<%=courseName%>">J</a></td>
          <td align="center"><a href="DLetra.asp?IdCurso=<%=IC%>&Letra=K&courseName=<%=courseName%>">K</a></td>
          <td align="center"><a href="DLetra.asp?IdCurso=<%=IC%>&Letra=L&courseName=<%=courseName%>">L</a></td>
          <td align="center"><a href="DLetra.asp?IdCurso=<%=IC%>&Letra=M&courseName=<%=courseName%>">M</a></td>
          <td align="center"><a href="DLetra.asp?IdCurso=<%=IC%>&Letra=N&courseName=<%=courseName%>">N</a></td>
          <td align="center"><a href="DLetra.asp?IdCurso=<%=IC%>&Letra=O&courseName=<%=courseName%>">O</a></td>
          <td align="center"><a href="DLetra.asp?IdCurso=<%=IC%>&Letra=P&courseName=<%=courseName%>">P</a></td>
          <td align="center"><a href="DLetra.asp?IdCurso=<%=IC%>&Letra=Q&courseName=<%=courseName%>">Q</a></td>
          <td align="center"><a href="DLetra.asp?IdCurso=<%=IC%>&Letra=R&courseName=<%=courseName%>">R</a></td>
          <td align="center"><a href="DLetra.asp?IdCurso=<%=IC%>&Letra=S&courseName=<%=courseName%>">S</a></td>
          <td align="center"><a href="DLetra.asp?IdCurso=<%=IC%>&Letra=T&courseName=<%=courseName%>">T</a></td>
          <td align="center"><a href="DLetra.asp?IdCurso=<%=IC%>&Letra=U&courseName=<%=courseName%>">U</a></td>
          <td align="center"><a href="DLetra.asp?IdCurso=<%=IC%>&Letra=V&courseName=<%=courseName%>">V</a></td>
          <td align="center"><a href="DLetra.asp?IdCurso=<%=IC%>&Letra=W&courseName=<%=courseName%>">W</a></td>
          <td align="center"><a href="DLetra.asp?IdCurso=<%=IC%>&Letra=X&courseName=<%=courseName%>">X</a></td>
          <td align="center"><a href="DLetra.asp?IdCurso=<%=IC%>&Letra=Y&courseName=<%=courseName%>">Y</a></td>
          <td align="center"><a href="DLetra.asp?IdCurso=<%=IC%>&Letra=Z&courseName=<%=courseName%>">Z</a></td>
        </tr>
      </table></td>
  </tr>
  
  <%
    // Calcular cantidad de palabras a mostrar
  var oRec, oConn;
  
  oConn = Server.CreateObject("ADODB.Connection");
  oRec = Server.CreateObject("ADODB.Recordset");
  
  if (Letter != "Todas") {
  	var Sql = "SELECT Count(Palabra) As Cantidad FROM Glosarios WHERE IdCurso=" + IC + " And Palabra LIKE '" + Letter + "%'";
  } else
  {
  	var Sql = "SELECT Count(Palabra) As Cantidad FROM Glosarios WHERE IdCurso=" + IC;
  }
  
  oConn.Open(Application("filePath"));
  oRec.Open(Sql, oConn,3,3);
  
  var Cnt = oRec.Fields.Item("Cantidad") * 1;
  oRec.Close();
  if (Letter != "Todas") {
    Sql = "SELECT * FROM Glosarios WHERE IdCurso=" + IC + " And Palabra LIKE '" + Letter + "%' ORDER BY Palabra";  
  } else
  {
    Sql = "SELECT * FROM Glosarios WHERE IdCurso=" + IC + "ORDER BY Palabra";
  }
    oRec.Open(Sql, oConn,3,3);

  %>
  
  <tr> 
    <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
        <%
          var c = 1;
          while (!oRec.EOF) {
              if (c == 1) {
        %>
        
        <tr> 
          <td width="150"><a href="DIdPalabra.asp?IdCurso=<%=IC%>&Letra=<%=Letter%>&Palabra=<%=oRec.Fields.Item("Palabra")%>&courseName=<%=courseName%>"><%=oRec.Fields.Item("Palabra")%></a></td>
          <td width="90%" rowspan="<%=Cnt%>" valign="top" class="LeftLinedCellCopy"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="4">
              <tr> 
                <td>
                  <font size="+1">
                   <strong>
                       <%
                         if (IDPal == 0) {
                       %>
                                          <%=oRec.Fields.Item("Palabra")%>
                       <%                   
                           } else {
                           
                           var oRec1 = Server.CreateObject("ADODB.Recordset");
                           var Sql1 = "SELECT * FROM Glosarios WHERE ID=" + IDPal;
                           oRec1.Open(Sql1, oConn,3,3);
                           
                       %>
                          <%=oRec1.Fields.Item("Palabra")%>
                       <%    
                          
                           }
                                      
                       %>
                   </strong>
                  </font>
                </td>
              </tr>
              <tr> 
                <td>
                   <p>
                       <%
                         if (IDPal == 0) {
                       %>
                                          <%=oRec.Fields.Item("Signif")%>
                       <%                   
                           } else {
                       %>
                          <%=oRec1.Fields.Item("Signif")%>
                       <%                   
                           }
                           
                       %>
                   </p>
                 </td>
              </tr>
            </table>
            </td>
        </tr>
        <%              
              } else {
        %>
               <tr>
                 <td><a href="DIdPalabra.asp?IdCurso=<%=IC%>&Letra=<%=Letter%>&Palabra=<%=oRec.Fields.Item("Palabra")%>&courseName=<%=courseName%>"><%=oRec.Fields.Item("Palabra")%></a></td>
               </tr>
        <%                 
               }
                oRec.Move(1); c += 1;
          }
            oRec.Close();

        %>
      </table></td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td><a href="glosario.asp?uid=<%=uid%>&Letra=Todas&IDP=0&IdCurso=<%=Session("course")%>&courseName=<%=Session("courseName")%>">Ir al inicio de la pagina</a></td>
  </tr>
</table>
</body>
</html>

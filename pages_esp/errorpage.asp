<%@ Language=JavaScript%>
<%
 Response.Expires = -1;
 
 if (Request.QueryString.Count == 3) 
    {
     tipo = Request.QueryString.Item('tipo');
     short = Request.QueryString.Item('short');
     desc = Request.QueryString.Item('desc');
     
    }
  else
    {
     tipo = "";
     short = "";
     desc = "";
    } 
      

%>
<html>
<head>
<title><%=tipo%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/aqua/ErrPage.css" type="text/css">
<script type="text/javascript" language="javascript" src="../js/Windows.js"></script>

</head>

<body style="overflow-y:scroll;overflow-x:scroll">
<table border="0" cellspacing="0" cellpadding="0" align="center" class="errTiltle">
  <tr> 
    <td><img src="../images/aqua/ErrorImg.gif" width="100" height="100"></td>
    <td><%=tipo%></td>
  </tr>
</table>
<hr noshade>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ErrMessage"><%=short%></td>
  </tr>
  <tr>
    <td class="ErrMessage">
      <hr noshade>
    </td>
  </tr>
</table>
<table width="80%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td align=center class="errDescription">
      <%=desc%>
    </td>
  </tr>
</table>
<br><br>
<table width="60%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td align=center><font size =1 ><b>
       Sistema de Enseñanza Personalizada a Distancia<br>
       SEPAD<br>
       Versión 1.5<br>
       Derechos Reservados RedWay 1999-2003</b>
       </font>
    </td>
  </tr>
  <tr> 
    <td align=center>&nbsp;
    </td>
  </tr>
  
  <tr> 
    <td align=center><font size =1 ><b>
       <a HREF="javascript:history.back(1)" >Regresar</a>
       </font>
    </td>
  </tr>
  
</table>

</body>
</html>


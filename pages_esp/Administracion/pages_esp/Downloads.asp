<%@ Language=JScript %>

<%
  Response.Buffer = true;
%>

<html>
<head>
<title>Descargas</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/EvalResult.css" type="text/css">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
</head>
<body bgcolor="#ffffff" text="#000000" style="OVERFLOW-Y: scroll">
<table width="100%" border="0" cellspacing="0" cellpadding="2">

  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td><IMG height=54 src="../images/<%=Session("skin")%>/DownloadsIMG.gif" width=80></td>
          <td class="HeaderTable">Descarga de Archivos</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="2" cellpadding="0">
  <tr> 
    <td width="15%"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td align="middle"><A href="../downloads/SepadXM1.0.zip"><IMG border=0 height=30 src="../images/<%=Session("skin")%>/xmail.gif" width=30></A></td>
        </tr>
        <tr> 
          <td align="middle"></td>
        </tr>     
        <tr>     
          <td align="middle">19/05/2002</td>
        </tr>     
        
      </table>
    </td>
    <td valign="top"> 
      <table width="100%" border="0" cellspacing="2" cellpadding="0">
        <tr> 
          <td width="1%"><b>T�tulo: </b></td>
          <td width="84%"><A href="../downloads/SepadXM1.0.zip">SEPADXMail 1.0</A>&nbsp;<STRONG>(758 kb)</STRONG></td>
        </tr>  
        <tr> 
          <td colspan="2"> Cliente para acceso a SEPAD a trav�s de protocolos de correo el�tronico.</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td colspan="2"> 
      <hr noshade>
    </td>
  </tr>

  <tr> 
    <td width="15%"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td align="middle"><A href="../downloads/Sepadhp1.5.zip"><IMG border=0 height=60 src="../images/<%=Session("skin")%>/sepadhp.gif" width=60></A></td>
        </tr>
        <tr> 
          <td align="middle"></td>
        </tr>     
        <tr>     
          <td align="middle">19/05/2002</td>
        </tr>     
        
      </table>
    </td>
    <td valign="top"> 
      <table width="100%" border="0" cellspacing="2" cellpadding="0">
        <tr> 
          <td width="1%"><b>T�tulo: </b></td>
          <td width="84%"><A href="../downloads/Sepadhp1.5.zip">SepadHP 1.5</A>&nbsp;<STRONG>(1.88 mb)</STRONG></td>
        </tr>  
        <tr> 
          <td colspan="2"> Herramienta para la publicaci�n en SEPAD 1.5.</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td colspan="2"> 
      <hr noshade>
    </td>
  </tr>
</table>
</body>
</html>

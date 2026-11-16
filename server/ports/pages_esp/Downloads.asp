<%@ Language=JScript %>
<%
// Port a JScript de pages_esp/Downloads.asp (original en VBScript).
%>
<!-- #include file="../inc/filelist.inc" -->
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
<table>
<%
  var files = listFiles(Server.MapPath("../downloads/"), false);
  for (var i = 0; i < files.length; i++) {
    var f = files[i];
    if (i % 2 == 0) {
%>
      <tr>
        <td width=16 align="left" valign="middle">
          <img src="../images/icons/<%=f.icon%>" width="16" height="16">
        </td>
        <td width=300 align="left" valign="middle">
          <a href= "../downloads/<%=f.name%>"><%=f.name%> (<%=f.size%> Kb)</a>
        </td>
<%  } else { %>
        <td width=16 align="left" valign="middle">
          <img src="../images/icons/<%=f.icon%>" width="16" height="16">
        </td>
        <td  align="left" valign="middle">
          <a href= "../downloads/<%=f.name%>"><%=f.name%> (<%=f.size%> Kb)</a>
        </td>
      </tr>
<%  }
  }
%>
   </table>
</body>
</html>

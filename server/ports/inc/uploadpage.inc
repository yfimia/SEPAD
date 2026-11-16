<%
// Port a JScript de Upload/default.asp (VBScript + clase FileUploader de upload.asp).
// Parámetros que define la página que incluye este archivo:
//   UPLOAD_VPATH   ruta virtual de la carpeta destino (relativa a la página)
//   UPLOAD_LINK    prefijo de los enlaces a los archivos
//   UPLOAD_ROOT    prefijo relativo hasta la raíz del sitio ("../../" o "../../../")
//   UPLOAD_ALLOWED true si el usuario actual puede subir archivos
%>
<!-- #include file="filelist.inc" -->
<%
Response.Expires = -1;
var uploadError = "";
var physPath = Server.MapPath(UPLOAD_VPATH);

if (Request.ServerVariables("REQUEST_METHOD") + "" == "POST") {
  if (!UPLOAD_ALLOWED) {
    uploadError = "No tiene permisos para subir archivos.";
  } else {
    try {
      var up = Server.CreateObject("aspSmartUpload.SmartUpload");
      up.DeniedFilesList = "asp,asa,inc,php,php3,pl,dll,exe,bat,com,js,htm,html,shtml,svg";
      up.MaxFileSize = 4 * 1024 * 1024;
      up.Upload();
      var fso = Server.CreateObject("Scripting.FileSystemObject");
      if (!fso.FolderExists(physPath)) fso.CreateFolder(physPath);
      for (var i = 0; i < up.Files.Count; i++) {
        var file = up.Files.Item(i);
        if (!file.IsMissing) file.SaveAs(UPLOAD_VPATH + file.FileName);
      }
    } catch (e) {
      uploadError = e.description || e.message;
    }
  }
}
%>
<html>
<head>
  <title>File Manager</title>
  <link rel="stylesheet" href="<%=UPLOAD_ROOT%>css/<%=Session("skin")%>/Main.css" type="text/css">
  <link rel="stylesheet" href="<%=UPLOAD_ROOT%>css/<%=Session("skin")%>/EvalResult.css" type="text/css">
</head>
<body>
 <table width="100%" border="0" cellspacing="0" cellpadding="2">
  <tr>
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr>
          <td><IMG height=54 src="<%=UPLOAD_ROOT%>images/<%=Session("skin")%>/DownloadsIMG.gif" width=80></td>
          <td class="HeaderTable">Subir Archivos</td>
        </tr>
      </table>
    </td>
  </tr>
 </table>
  <form name="SentFile" method="post" action="default.asp" enctype="multipart/form-data">
    <table cellpadding=0 cellspacing=1 align=center width=300px>
<% if (uploadError != "") { %>
      <tr><td colspan=3><font color="#CC0000"><b><%=Server.HTMLEncode(uploadError)%></b></font><br><br></td></tr>
<% } %>
<% if (UPLOAD_ALLOWED) { %>
      <tr>
        <td colspan=3>
          <input name="newattachfile" type="file" size="60"><br>
          <input type="Submit" name="Submit" value="Subir">
           <br> <br>
        </td>
      </tr>
<% } %>
      <tr>
        <td colspan=3><h3>Lista de archivos</td>
      </tr>
<%
  var files = listFiles(physPath, UPLOAD_ALLOWED);
  for (var i = 0; i < files.length; i++) {
    var f = files[i];
%>
      <tr>
        <td width=16 align="left" valign="middle">
          <img src="<%=UPLOAD_ROOT%>images/icons/<%=f.icon%>" width="16" height="16">
        </td>
        <td width="100%" align="left" valign="middle">
          <a target="blank" href= "<%=UPLOAD_LINK%><%=f.name%>" ><%=f.name%> (<%=f.size%> Kb)</a>
        </td>
      </tr>
<% } %>
    </table>
  </form>
</body>
</html>

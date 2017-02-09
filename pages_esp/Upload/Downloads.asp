<%@ Language=VBScript %>
<% 
  Function extImage(filename)
    Dim pam_arr, pam_icon, strPmExtension
        pam_icon = "attach.gif"
        pam_arr = Split(filename, ".", -1, 1)
        If Ubound(pam_arr) > 0 Then
          strPmExtension = LCase(pam_arr(Ubound(pam_arr)))
          Select Case strPmExtension
                Case "txt"
                  pam_icon = "text_plain.gif"
                Case "bat"
                  pam_icon = "executable.gif"
                Case "exe"
                  pam_icon = "executable.gif"
                Case "com"
                  pam_icon = "executable.gif"
                Case "asp"
                  pam_icon = "application_asp.gif"
                Case "inc"
                  pam_icon = "application_asp.gif"
                Case "asa"
                  pam_icon = "application_asp.gif"
                Case "css"
                  pam_icon = "application_css.gif"
                Case "doc"
                  pam_icon = "application_doc.gif"
                Case "html"
                  pam_icon = "application_html.gif"
                Case "htm"
                  pam_icon = "application_html.gif"
                Case "shtml"
                  pam_icon = "application_html.gif"
                Case "phtml"
                  pam_icon = "application_html.gif"
                Case "pdf"
                  pam_icon = "application_pdf.gif"
                Case "xls"
                  pam_icon = "application_xls.gif"
                Case "bmp"
                  pam_icon = "image_bmp.gif"
                Case "gif"
                  pam_icon = "image_gif.gif"
                Case "jpg"
                  pam_icon = "image_jpeg.gif"
                Case "jpeg"
                  pam_icon = "image_jpeg.gif"
                Case "psd"
                  pam_icon = "image_psd.gif"
                Case "tif"
                  pam_icon = "image_tiff.gif"
                Case "tiff"
                  pam_icon = "image_tiff.gif"
                Case "rar"
                  pam_icon = "pack_rar.gif"
                Case "zip"
                  pam_icon = "pack_zip.gif"
                Case "gz"
                  pam_icon = "pack_zip.gif"
                Case "tar"
                  pam_icon = "pack_zip.gif"
                Case "eml"
                  pam_icon = "mail.gif"
                Case Else
                  pam_icon = "attach.gif"
          End Select
        End If
    extImage = pam_icon
  End Function
%>  

<html>
<head>
<title>Descargas</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/EvalResult.css" type="text/css">
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/Main.css" type="text/css">
</head>
<body bgcolor="#ffffff" text="#000000" style="OVERFLOW-Y: scroll">
<table width="100%" border="0" cellspacing="0" cellpadding="2">

  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td><IMG height=54 src="../../images/<%=Session("skin")%>/DownloadsIMG.gif" width=80></td>
          <td class="HeaderTable">Descarga de Archivos</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td colspan=3><h3>Lista de archivos</td>
  </tr>
</table>
<table>  
<% 

       Dim contador
       contador = 0 
       Dim PathToUpLoad
       idcurso = Session("course")
       PathToUpLoad = Server.MapPath("../../Courses/Course" + idcurso + "/descargas/")
        Dim pam_i, pam_icon, pam_size
        Set pam_fso = Server.CreateObject("Scripting.FileSystemObject")
        If (pam_fso.FolderExists(PathToUpload)) Then
          Set pam_f = pam_fso.GetFolder( PathToUpload )
          Set pam_fc = pam_f.Files
          pam_i   = 0
          For Each pam_f1 In pam_fc
            pam_i =       pam_i + 1
            pam_icon = extImage(pam_f1.name)
            pam_size = pam_f1.size / 1024
            If pam_size < 1       Then
              pam_size = CInt(pam_size * 10) / 10
            Else
              pam_size = CInt(pam_size)
            End If 
      if (contador Mod 2 = 0) then
      contador = contador + 1
      %>
      <tr>   
        <td width=16 align="left" valign="middle">
          <img src="../../images/icons/<%=pam_icon%>" width="16" height="16">
        </td>
        <td width=300 align="left" valign="middle">
          <a target="blank" href= "../../courses/course<%=Session("course")%>/descargas/<%=pam_f1.name%>"><%=pam_f1.name%> (<%=pam_size%> Kb)</a>
        </td>
      <% else contador = contador + 1 %>  
        <td width=16 align="left" valign="middle">
          <img src="../../images/icons/<%=pam_icon%>" width="16" height="16">
        </td>
        <td  align="left" valign="middle">
          <a target="blank" href= "../../courses/course<%=Session("course")%>/descargas/<%=pam_f1.name%>"><%=pam_f1.name%> (<%=pam_size%> Kb)</a>
        </td>
      </tr>
      <%
        End if 
          Next
        Else  
          pam_fso.CreateFolder(PathToUpload) 
        End If
      %>
   </table> 
</body>
</html>

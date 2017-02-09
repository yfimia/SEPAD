<%@language=jscript%>
<%

  Response.Expires = -1;
  
/*  var IC     = Request.QueryString.Item("IdCurso");
  var courseName = Request.QueryString.Item("courseName");
  var Letter = Request.QueryString.Item("Letra");
  var IDPal  = Request.QueryString.Item("IDP");
*/  
%>
<?xml version="1.0" encoding="iso-8859-1"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<!--link href="../tinymce/docs/style.css" rel="stylesheet" type="text/css"-->
<link rel="stylesheet" href="css/EvalResult.css" type="text/css">
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/Main.css" type="text/css">
<!--link rel="stylesheet" href="css/Main.css" type="text/css"-->
<link href="css/Aqua/Glosary.css" rel="stylesheet" type="text/css" />

<!-- tinyMCE -->
<script language="javascript" type="text/javascript" src="../tinymce/jscripts/tiny_mce/tiny_mce.js"></script>
<script language="javascript" type="text/javascript">
	tinyMCE.init({
		mode : "textareas",
		theme : "advanced",
		plugins : "table,save,advhr,advimage,advlink,emotions,iespell,insertdatetime,preview,zoom,flash,searchreplace,print",
		theme_advanced_buttons1_add_before : "save,separator",
		theme_advanced_buttons1_add : "fontselect,fontsizeselect",
		theme_advanced_buttons2_add : "separator,insertdate,inserttime,preview,zoom,separator,forecolor,backcolor",
		theme_advanced_buttons2_add_before: "cut,copy,paste,separator,search,replace,separator",
		theme_advanced_buttons3_add_before : "tablecontrols,separator",
		theme_advanced_buttons3_add : "emotions,iespell,flash,advhr,separator,print",
		theme_advanced_toolbar_location : "top",
		theme_advanced_toolbar_align : "left",
		theme_advanced_path_location : "bottom",
		content_css : "example_full.css",
	    plugin_insertdate_dateFormat : "%Y-%m-%d",
	    plugin_insertdate_timeFormat : "%H:%M:%S",
		extended_valid_elements : "a[name|href|target|title|onclick],img[class|src|border=0|alt|title|hspace|vspace|width|height|align|onmouseover|onmouseout|name],hr[class|width|size|noshade],font[face|size|color|style],span[class|align|style]",
		external_link_list_url : "example_link_list.js",
		external_image_list_url : "example_image_list.js",
		flash_external_list_url : "example_flash_list.js",
		file_browser_callback : "fileBrowserCallBack"
	});

	function fileBrowserCallBack(field_name, url, type) {
		// This is where you insert your custom filebrowser logic
		alert("Filebrowser callback: " + field_name + "," + url + "," + type);
	}
</script>
<!-- /tinyMCE -->
</head>

<body>
<a name="PageTop" id="PageTop"></a> 
<table width="100%" border="0" cellspacing="0" cellpadding="4">
  <tr> 
    <td class="ToolBar"><table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td><img height="54" src="../../images/<%=Session("skin")%>/GlosarioIMG.gif" width="80"></td>
          <td class="HeaderTable">Adicionar Palabra</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
   <form id="formDatosAgregar" name="formDatosAgregar" action="procesoAgregarPalabra.asp" method="post">
     <table border=0 cellpadding=0 cellspacing=0 width="100%">
       <tr class="MessageTR">
        <td align="center" width="1%">Palabra:</td><td><input type=text id="TextPal" name="TextPal" size=60></td>
       </tr>
       <tr class="MessageTR1">
        <td align="center" width="1%" valign="top">Significado:</td><td><textarea id="TextSignif" name="TextSignif" cols=60 rows=28></textarea></td>
       </tr>
       <tr>
        <td colspan=2 align="center" class="Toolbar"><input type=submit value="Agregar"></td>
       </tr>
     </table>
   </form> 
  </tr>
</table>
</body>
</html>

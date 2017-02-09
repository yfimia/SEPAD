<%@ Language=JScript %>
<%Response.Expires = -1;%>
<!-- #include file='../js/user.inc' -->

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

%>

<%
if (Session("PermissionType") == ADMINISTRATOR)
    {
      if ((Request.Form("Gname").Count == 0) || (Request.Form("Gname") == ""))
        {
%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<!--link href="tinymce/docs/style.css" rel="stylesheet" type="text/css"-->
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/EvalResult.css" type="text/css">

<!-- tinyMCE -->
<script language="javascript" type="text/javascript" src="tinymce/jscripts/tiny_mce/tiny_mce.js"></script>
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

<body bgcolor="#FFFFFF" text="#000000">
  <form name=newgroup id=newgroup action="agregagroup.asp?uid=<%=uid%>" method=post>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td class="HeaderTable"  align=center><b>Adicionar grupo</b></td>
        </tr>
      </table>
    </td>
  </tr>
</table>

    <table border="0" cellspacing="1" cellpadding="1" align="center">
      <tr>
        <td  align=center Class="MessageTR1">
          Nombre del nuevo grupo:
        </td>
        <td  Class="MessageTR1"align=center >
          <input name=Gname Id=Gname type=text maxlength=250 size=50 >
        </td>
      </tr>
      <tr>
        <td align=center   Class="MessageTR">
          Descripción del grupo:
        </td>
        <td align=center   Class="MessageTR">
          <textarea name=GDescript Id=GDescript" rows="24" cols="100" maxlength=250></textarea>
        </td>
      </tr>
      <tr>
        <td colspan=2 align=center Class="Toolbar">
          <input type=submit value="Adicionar">
        </td>      
      </tr>
    </table>
  </form>
<%
        }
      else
        {
          var  filePath = Application("dataPath");   
          var  oConn    = Server.CreateObject("ADODB.Connection");
          var  oComm    = Server.CreateObject("ADODB.Command");
          var  oRec     = Server.CreateObject("ADODB.Recordset");
          oConn.Open( filePath );
          oRec.Open("select * from grupos where (Name = '" + Request.Form("Gname") + "')",oConn,3,3);
          //Response.Write("select * from grupos where (Name = '" + Request.Form("Gname") + "')");
          if (oRec.EOF == false)
            { 
              oConn.Close();   
              
              Response.Redirect("errorpage.asp?tipo=Error&short=" + GRUOP_NAME_EXIST_SHORT  + "&desc=" + GRUOP_NAME_EXIST_TEXT); 
              
              Response.Write ("<center><font color = red> El grupo " + Request.Form("Gname") + " ya existe. Seleccione otro nombre para su grupo</font><br>" );
              Response.Write ('<script languaje=javascript>function pclick(){window.location="usermanager.asp";}</script>');      
              Response.Write ("<input type=Button Value=Regresar onclick='return pclick()' id=Button1 name=Button1></center>" );      ;
            }  
          else
            {
              oConn.Close();   
              oConn.Errors.Clear();
              oConn.Open( filePath );
              oRec.Open("select * from grupos",oConn,3,3);
              oRec.AddNew();
              oRec.Fields.Item("Name").Value = Request.Form("Gname") + "";
              if ((Request.Form("GDescript").Count != 0) && (Request.Form("GDescript") + "" != ""))
                { 
                  oRec.Fields.Item("Description").Value = Request.Form("GDescript") + ""; 
                }
              else oRec.Fields.Item("Description").Value = " ";   
              oRec.Update();      
              oConn.Close();   
              
              Response.Redirect("agregagroup.asp?uid=" + uid);      
              
              Response.Write ("<center><font color = red>Grupo creado satisfactoriamente</font><br>" );
              Response.Write ('<script languaje=javascript>function pclick(){window.location="usermanager.asp";}</script>');      
              Response.Write ("<input type=Button Value=Regresar onclick='return pclick()' id=Button1 name=Button1></center>" );      ;
            }  
        }
    }    
  else
    {  
     Response.Redirect("errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);                   
    }
%>

</BODY>
</HTML>
<%
  }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
  //Response.Write(Request.QueryString.Item("uid"));
%>



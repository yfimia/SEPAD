<%@language=jscript%>
<%
    Response.Expires = -1;
%>

<!-- #include file='../js/user.inc' -->
<!-- #include file='inc/crearanuncio.inc' -->
<%    
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

    idmod = Request.querystring("idmodulo");
%>
<html>
<head>
<title><%=title%></title>
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/EvalResult.css" type="text/css">
<link href="tinymce/docs/style.css" rel="stylesheet" type="text/css">

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
<body class="MainBody">
<table width="100%" border="0" cellspacing="0" cellpadding="2">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td><img src="../images/<%=Session("skin")%>/NoticiasIMG.gif" width="80" height="54"></td>
          <td class="HeaderTable"><%=stmp1%></td>
        </tr>
      </table>
    </td>
  </tr>
</table>

 <form name="MakeAnuncio" id="MakeAnuncio" method="Post" action="InsertarDatosAnuncios.asp?uid=<%=uid%>">
   <table border=0 cellpadding=0 cellspacing=0 width="100%">
      <tr class="MessageTR1">
          <td width="20%" align="right"><%=stmp2%></td>
          <td width="80%"><input class=TextArea  type=text name="title" id="title" value=""></td>
      </tr>    
      <tr class="MessageTR">
          <td valign="top" align="right" width="20%"><%=stmp3%></td>
          <td width="80%"><textarea class=TextArea  name="body" id="body" rows=24 ></textarea></td>
      </tr>    
      <tr class="MessageTR1">
          <td align="right"><%=stmp4%></td>
          <td>
             <table border=0 cellpadding=0 cellspacing=0>
               <tr>
                   <td>
                      <input type=checkbox id="ExpSoN" name="ExpSoN"><%=stmp5%>
                   </td>
               </tr>   
             </table>
          </td>
      </tr>    
      <tr>
          <table border=0 cellpadding=1 cellspacing=1>
            <tr class="MessageTR">
              <td align="right" width="40%"><%=stmp6%></td>
              <td align="left" width="5%"><%=stmp7%></td>
              <td width="5%">
                  <select name="DaysEx" id="DaysEx">
                    <%
                      var i = 1;
                      while (i<=31) {
                        Response.write("<option>");
                        Response.write(i);
                        Response.write("</option>");
                        
                        i++;
                      }
                    %>
                  </select>
              </td>
              <td align="left" width="5%"><%=stmp8%></td>
              <td width="10%">
                  <select name="MonthEx" id="MonthEx">
                    <option value=1><%=stmp9%></option>
                    <option value=2 ><%=stmp10%></option>
                    <option value=3 ><%=stmp11%></option>
                    <option value=4 ><%=stmp12%></option>
                    <option value=5 ><%=stmp13%></option>
                    <option value=6 ><%=stmp14%></option>
                    <option value=7 ><%=stmp15%></option>
                    <option value=8 ><%=stmp16%></option>
                    <option value=9 ><%=stmp17%></option>
                    <option value=10 ><%=stmp18%></option>
                    <option value=11 ><%=stmp19%></option>
                    <option value=12 ><%=stmp20%></option>
                  </select>
              </td>
              <td align="left" width="5%"><%=stmp21%></td>
              <td align="left">
                  <select name="YearEx" id="YearEx">
                    <%  
                        
                        
  var today = new Date();
  
  for (i=today.getYear();i<=today.getYear() + 50;i++)
  {
    %><option value="<%=i%>"><%=i%></option><%
  }
                        
                    %>
                  </select>
              </td>
            </tr> 
          </table>
      </tr>    
      <tr>
         <table border=0 align=center>
           <tr>
              <td><input type=hidden name="IDModulo" id="IDModulo" value=<%=idmod%>></td>
           </tr>
         </table>  
      </tr>
      <tr>
            <table border=0 cellpadding=1 cellspacing=1 align="center">
              <tr>
               <td width="100%"><input type=submit name="SendInfo" id="SendInfo" value=<%=stmp22%>></td>
              </tr> 
      </tr>
   </table>
 </form>
</body>
</html>
<%
 }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>

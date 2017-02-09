<%@ Language=JavaScript %>

<%Response.Expires = -1%>
<!-- #include file='../js/user.inc' -->
<!-- #include file='inc/viewmail.inc' -->

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

%>


<%
   var Ind;
   
   if (Request.QueryString.Count  > 0 ) 
     {
      Ind           = Request.QueryString.Item("ID");
      Tmp           = Session("MailContent");                       
       
      HTML = "<TABLE Class='MessageTable' cellspacing='0' cellpadding='0' Height='60%'>" +
              "<TR>" + 
                "<TD  height='5%' Class='MainTD' Width='10%'>" + TEXT1 + ":</TD><TD height='5%' Class='DatTD' Width='90%'>" + Tmp[Ind].From + "</TD>" +
              "</TR>"  + 
              "<TR>" + 
                "<TD Class='MainTD'' height='5%' Width='10%'>" + TEXT2 + ":</TD><TD height='5%' Class='DatTD'' Width='90%'>" + Tmp[Ind].Subject + "</TD>" +
              "</TR>" +              
              "<TR>" + 
                "<TD colspan='2' height='50%'>" +
                "<TABLE cellspacing='0' cellpadding='0' Class='MainTable' height='100%'>" +
                  "<TR>" + 
                    "<TD Width='100%' height='100%'>" + 
                   "<TEXTAREA Class='MArea' rows='24'>" + Tmp[Ind].Content +"</TEXTAREA>" + 
                "</TD>" +                 
                "</TR>" + 
                "</TABLE>" +
                "</TD>" + 
              "</TR>" + 
              "<TR Height='5%'>" + 
                "<TD colspan='2'>" +
                "<TABLE width='100%' Height='100%' cellspacing='0' cellpadding='0'>" +
                  "<TR>" + 
                    "<TD Width='100%' Height='100%' >" + 
                      "<INPUT Type='Button' Value='Volver' onclick='return Ionclick()'>" + 
                "</TD>" +                 
                "</TR>" + 
                "</TABLE>" +
                "</TD>" + 
              "</TR>" + 
             "</TABLE>";
    
     if (Session("MailContent") != null) {Session("MailContent") = null};
     if (Session("MessageID")!= null) {Session("MessageID") = null};
    } 
      
%>

<HTML>
<HEAD>
<!--link href="tinymce/docs/style.css" rel="stylesheet" type="text/css"-->
<link REL="stylesheet" TYPE="text/css" HREF="../css/<%=Session("skin")%>/SepadCss.css" />

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

<Script languaje="jscript">
   function Ionclick()
    {
     <% if (Request.QueryString("Flag") ==1){ %>
        location.href = "../beta/mail.asp";
     <%} else{ %> 
        location.href = "mailmanager.asp?uid=<%=uid%>";       
     <%}%>
    }
</Script>
</HEAD>

<BODY>
<%=HTML%>
<BODY>
</HTML>
<%
  }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
  //Response.Write(Request.QueryString.Item("uid"));
%>

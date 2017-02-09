<%@ Language=jScript %>

<%
Response.Expires = -1;
%>

<script language=vbscript  runat=Server>
function getDate
getDate = Now()
End function 
</script>
<!-- #include file='../js/adolibrary.inc' -->
<!-- #include file='../js/user.inc' -->
<%
   var      MailTo = "";
   var      MailToText = "";
   var      MailBody = "";
   var      MailSubject = "";
   var	    reply = false;
   var      forward = false;

    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       
      if (Request.Form.Count >= 4) {
         MailToText = Request.Form("MailToName");
         MailTo = Request.Form("MailToText");
         MailBody = Request.Form("MailBody");
         if (MailTo == '-1') {
           forward = true;
         if (forward)
           MailSubject = "Fw: " + Request.Form("MailSubject");
           
         }
         reply = true;
      }
         if (reply && !forward)
           MailSubject = "Re: " + Request.Form("MailSubject");


%>




<html>
<head>
<title>Nuevo Mensaje</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
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
<script language=jscript src ="../js/library.js"></script>
<script language=jscript>
  function _Send()
   {  
    if (notIsScript(Subject.value))
     {        
     
<%
if (!reply || forward)  {
%>     //alert(Select.value + " " + text);
    var p = Select.selectedIndex;
    var text = Select.options(p).text;   
    Mail.MailTo.value   = Select.value;    
    Mail.MailToText.value   = text;    
    Mail.MailBody.value     = Texto.value;      
    Mail.MailSubject.value  = Subject.value;
    Mail.MailPriority.value = Priority.checked == true?1:0
<%
  }
  else {
%>  
    Mail.MailBody.value     = Texto.value;      
    Mail.MailSubject.value  = Subject.value;
    Mail.MailPriority.value = Priority.checked == true?1:0
<%      
  }  
%>    
    Mail.submit();      
   }
   else
     alert('Asunto no válido, revise la presencia de caráteres especiales');
   }
   
//-->
</script>

</head>

<body bgcolor="#FFFFFF" text="#000000">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="MessageBarTD"> 
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td><a href="javascript:_Send()" class="ToolLink">&nbsp;Enviar&nbsp;</a></td>
          <td>|</td>
          <td><a href="javascript:close()" class="ToolLink">&nbsp;Cerrar&nbsp;</a></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td class="MessageTR1"> 
            <table border="0" cellspacing="0" cellpadding="0" width="100%">
              <tr> 
                <td width="1%">Para:</td>
                <td width="99%"> 
<%
  if (reply && !forward) {	           
%> 
                     <%=MailToText %>            
<%
  }     
  else {
%>      
                  <select id="Select" name="Select" class="ComboBox">

<%

  var oConn    = Server.CreateObject("ADODB.Connection");  
  var oRec     = Server.CreateObject("ADODB.Recordset");
  var filePath = Application("filePath");
  
                 
  oConn = MakeConnection( oConn, filePath );

  var Sql      = "SELECT * FROM Grupos WHERE (NOT ID IN (" + GUEST_GROUP + ")) ORDER BY Name";
  oRec = Query( Sql, oRec, oConn  );		
  while (oRec.EoF == false) 
   {     
    //se codifican los valores de los option de la forma [0 grupo/1 usuario] + indice   
%>      
  <OPTION selected VALUE='0_<%=oRec.Fields("ID").Value%>'>
    Grupo : <%=oRec.Fields("Name").Value%>
  </OPTION>
<%  
    oRec.Move(1);
   }

  oRec.Close();
%>

<%

  //********** Patch ( 30/6/2004 ) **************
  //Se codifican los valores de los option de la forma [0 grupo/1 usuario/2 Subgrupo] + indice      
  
  Sql  = "SELECT B.* FROM Usuarios A INNER JOIN Subgrupo B ON B.Tutor = A.ID  ORDER BY B.Name";
  oRec = Query( Sql, oRec, oConn  );		
  
  while (oRec.EoF == false) 
  {     
    
%>      
  <OPTION selected VALUE='2_<%=oRec.Fields("ID").Value%>'>
    SubGrupo : <%=oRec.Fields("Name").Value%>
  </OPTION>
<%  
    oRec.Move(1);
   }
      
   oRec.Close();
 
  //********** Patch ( 30/6/2004 ) **************
%>  

<%    
  var Sql      = "SELECT * FROM Usuarios WHERE (NOT ID IN (" + GUEST_USER + ")) ORDER BY FullName";
  oRec = Query( Sql, oRec, oConn  );		
  while (oRec.EoF == false) 
   {     
    //se codifican los valores de los option de la forma [0 grupo/1 usuario] + indice   
%>      
  <OPTION selected VALUE='1_<%=oRec.Fields("ID").Value%>'>
    <%=oRec.Fields("FullName").Value%>
  </OPTION>
<%  
    oRec.Move(1);
   }
   
  DestroyAdoObjects( oConn, oRec );  
%>

                  </select>
<%
  }
%>                  
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td class="MessageTR1"> 
            <table border="0" cellspacing="0" cellpadding="0" width="100%">
              <tr> 
                <td width="1%">Asunto:</td>
                <td width="99%"> 
                  <input type="text" id=Subject name="textfield" class="Edit" size="50" value="<%=MailSubject%>">
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td class="MessageTR1"> 
            <textarea name="textfield2" ID=Texto rows="28" cols="100" class="TextArea"><%=MailBody %></textarea>
          </td>
        </tr>
        <tr> 
          <td class="MessageTR1"> 
            <table border="0" cellspacing="0" cellpadding="1">
              <tr> 
                <td> 
                  <input type="checkbox" ID=Priority name="checkbox" value="checkbox">
                </td>
                <td>Establecer nivel de prioridad.</td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td class="MessageBarTD"> 
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td>&nbsp;&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<FORM ID=Mail NAME=Mail METHOD=POST ACTION="SendMail.asp?uid=<%=uid%>" >
  <INPUT TYPE=hidden NAME=MailTo ID=MailTo VALUE="1_<%=MailTo%>">
  <INPUT TYPE=hidden NAME=MailToText ID=MailToText VALUE="<%=MailToText%>">
  <INPUT TYPE=hidden NAME=MailBody ID=MailBody VALUE="<%=MailBody%>">  
  <INPUT TYPE=hidden NAME=MailSubject ID=MailSubject VALUE="<%=MailSubject%>">
  <INPUT TYPE=hidden NAME=MailPriority ID=MailPriority VALUE=0>      
</FORM>
</body>
</html>
<%
 }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>

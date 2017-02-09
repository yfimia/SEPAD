<%@ Language=JScript %>
<%Response.Expires = -1;%>
<!-- #include file='../../js/adoLibrary.inc' -->
<!-- #include file='../../js/user.inc' -->
<!-- #include file='../../js/change.inc' -->

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

%>

<%
 

 if ((Session("PermissionType") == ADMINISTRATOR) || (Session("userID") == Session("admcursocordinador"))  || (Session("userID") == Session("admcursoOwner")))
    {
      if ((Request.Form("Mname").Count == 0) || (Request.Form("Mname") == ""))
        {
        
        
%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<!--link href="../tinymce/docs/style.css" rel="stylesheet" type="text/css"-->
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/EvalResult.css" type="text/css">

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
<script language=jscript>

function unchange(ss) {
    	var cnt = 0;
    	s = new String(ss);
    	var cl = "<br/>";
    	    brtext = "";
    	while (s.search(cl) != -1) {
         	s = s.replace(cl,brtext);
        } 
    	var cl = "<BR/>";
    	    brtext = "";
    	while (s.search(cl) != -1) {
         	s = s.replace(cl,brtext);
        } 
        
 return(s);        
}


 function change(ss) {
    	var cnt = 0;
    	s = new String(ss);
    	var cl = "\n";
    	    brtext = "<BR/>";
    	while (s.search(cl) != -1) {
         	s = s.replace(cl,brtext); 
        }
 return(s);        
}
  function User(id, name, fullName, email) {
    this.id = id;
    this.name = name;
    this.fullName = fullName;
    this.email = email;
  }
  
  
  function findUser() {
    var user = new Array();
    user = showModalDialog("findTutor.asp?uid=<%=uid%>", "", "");
    if (user != null) {
      newgroup.Mcord.value = user[0].fullName + ' (' + user[0].name + ')';
      newgroup.Mtutorid.value = user[0].id;
    }  
  }
  
</script>

<body bgcolor="#FFFFFF" text="#000000">
  <form name=newgroup id=newgroup action="addSubGrupo.asp?uid=<%=uid%>" method=post>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td class="HeaderTable"  align=center><b>Adicionar Subgrupo Tutoreado</b></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
    <table width="100%" border="0" cellspacing="1" cellpadding="1" align="center">
      <tr>
        <td  align=right Class="MessageTR">
          Nombre:
        </td>
        <td  Class="MessageTR" align=left >
          <input name=Mname Id=Mname type=text maxlength=250 size=60 value="">
        </td>
      </tr>
      <tr>
      <tr>
        <td align=right   Class="MessageTR1">
          Tutor:
        </td>
        <td align=left   Class="MessageTR1">
          <input disabled name=Mcord Id=Mcord type=text maxlength=250 size=35 value=""><input name=Mname Id=Mname type=button size=20 value="Seleccionar..." onclick="findUser()">
          <input name=Mtutorid Id=MtutorId type=hidden maxlength=250 size=25  value="">
        </td>
      </tr>

    </table>

    <table width="100%" border="0" cellspacing="1" cellpadding="1" align="center">
      <tr>
        <td colspan=2 align=center   Class="MessageTR">
          Descripción
        </td>
      </tr>
      <tr>
        <td colspan=2 align=left   Class="MessageTR">
          <textarea name=Mdesc  rows="24" cols="78" maxlength=250></textarea>
        </td>
      </tr>
      
      <tr>
        <td widht="100%" colspan=2 align=center Class="Toolbar">
          <center><input type=submit value="Adicionar" ></center>
        </td>      
      </tr>
    </table>
  </form>
<%

        }
      else
        {
           var filePath = Application('dataPath');
           var oConn;
           var oRec;
 
           oConn = MakeConnection( oConn, filePath );
           Sql = "select SubGrupo.* FROM SubGrupo WHERE (SubGrupo.Curso = "  + Session("admcurso") + ") and (Name = '" + Request.Form("Mname") + "')";

           oRec = Query( Sql, oRec, oConn  );	
          

          if (oRec.EOF != true)
            { 
              oConn.Close();   
              
              Response.Redirect("../errorpage.asp?tipo=Error&short=Ya existe el subgrupo&desc=Ya existe un subgrupo con el nombre especificado."); 
            }  
          else
            {
              oRec.addNew();            
            
              if ((Request.Form("MName").Count != 0) && (Request.Form("Mname") + "" != ""))
                { 
                  oRec.Fields.Item("Name").Value = Request.Form("Mname") + "";
                }
              else Response.Redirect("../errorpage.asp?tipo=Error&short=Nombre no válido&desc=El nombre del subgrupo introducido no es válido o se ha dejado en blanco.");


              if ((Request.Form("Mdesc").Count != 0) && (Request.Form("Mdesc") + "" != ""))
                { 
                  oRec.Fields.Item("Desc").Value = Request.Form("Mdesc") + "";
                }
              else Response.Redirect("../errorpage.asp?tipo=Error&short=Descripción no válida&desc=La descripción del subgrupo introducida no es válida o se ha dejado en blanco.");

              if ((Request.Form("Mtutorid").Count != 0) && (!isNaN(parseInt(Request.Form("Mtutorid") + ""))))
                { 
                  oRec.Fields.Item("Tutor").Value = Request.Form("Mtutorid");
                }
              else Response.Redirect("../errorpage.asp?tipo=Error&short=Tutor no válido&desc=El tutor no ha sido seleccionado apropiadamente. Para seleccionar el tutor, presione el botón Seleccionar y escójalo de la lista.");

              oRec.Fields.Item("Curso").Value = Session("admcurso");

             
              oRec.Update();      
              oConn.Close();   
              
              Response.Redirect("addSubGrupo.asp?uid=" + uid);      
            }  
        }
    }    
  else
    {  
     Response.Redirect("../errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);                   
    }
%>

</BODY>
</HTML>
<%
  }
 else
   Response.Redirect("../errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
   
  //Response.Write(Request.QueryString.Item("uid"));
%>



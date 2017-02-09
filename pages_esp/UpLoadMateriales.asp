<%@ Language=jScript %>
<%  
 Response.Expires = -1;
%>
<!-- #include file='../js/user.inc' -->
<!-- #include file='inc/UpLoadMateriales.inc' -->
<%
  if (
  	(Request.QueryString.Item("uid") + "" == Session("uid")) &&
     	(Request.QueryString.Item("material").Count != 0) &&
     	(Request.QueryString.Item("lesson").Count != 0)
     )
    {    
      var uid 	    = Request.QueryString.Item("uid") + "";             	                  
      var  filePath = Application("dataPath");   
      var  oConn    = Server.CreateObject("ADODB.Connection");      
      var  oRec     = Server.CreateObject("ADODB.Recordset");
      var  mid	    = Request.QueryString.Item("material") + "";
      var url       = Request.QueryString.Item("url") + "";
      var sql	    = 'SELECT * from Ficheros WHERE (ID = ' + mid + ')';            
      var lesson    = Request.QueryString.Item("lesson") + "";
      
      oConn.Open( filePath );	      
      oRec.Open( sql, oConn, 3, 3);
      
      var _title       = oRec.Fields.Item("title").Value;
      var _description = oRec.Fields.Item("description").Value;	
      
      oRec.Close();	
%>
<html>
<head>
<title><%= TITLE1 %></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/EvalResult.css" type="text/css">
<script language="JavaScript">
<!--
function SubmitForm(form_name, submit_url) {
	if (confirm("¿Está seguro que desea borrar el archivo?") == true ) {
	document.all(form_name).action = submit_url;
	document.all(form_name).submit();
	}
}

//-->
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td><img src="../images/<%=Session("skin")%>/LessonImg.gif" width="80" height="54"></td>
          <td class="HeaderTable"><%= TITLE1 %></td>
        </tr>
      </table>
    </td>
  </tr>
</table>

<form name=UpLoadMaterialForm action="UploadMaterialFile.asp?uid=<%=uid%>&lesson=<%= lesson %>&url=<%= url %>&mid=<%= mid %>" method="post" enctype="multipart/form-data">

<table border="0" cellspacing="5" cellpadding="5" align="center">
  <tr>
    <td>
      <table border="0" cellspacing="0" cellpadding="0" class="PrologueTable">
        <tr> 
          <td> 
            <table border="0" cellspacing="1" cellpadding="2">
              <tr align="left"> 
                <td colspan="4" class="ToolBar"> 
                  <table border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td class="ToolBar"><b><%= TEXT1 %>:</b></td>
                      <td class="ToolBar"> 
                        &nbsp;
                        <% //var fcourse = Session("FullCourse"); Response.Write(fcourse.MaxFile);
                            
                        %>10240&nbsp;kb&nbsp;
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>            
              <tr align="left"> 
                <td colspan="4" class="ToolBar"> 
                  <table border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td class="ToolBar"><b><%= TEXT2 %>:</b></td>
                      <td class="ToolBar"> 
                        <input type="file" name="FileName" size="70" class="Edit" >
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr align="left"> 
                <td colspan="4" class="ToolBar"> 
                  <table border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td class="ToolBar"><b><%= TEXT5 %>:</b></td>
                      <td class="ToolBar"> 
                        <input type="text" name="FileTitle" size="70" class="Edit" value="<%= _title %>">                        
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
              
              <tr> 
                <td width="100%" colspan="4" class="MessageTR1"><b><%= TEXT3 %>:</b></td>
              </tr>
              <tr> 
                <td width="100%" colspan="4" class="MessageTR1"> 
                  <textarea name="FileDescription" rows="8" cols="100"><%= _description %></textarea>
                </td>
              </tr>
              <tr> 
                <td colspan="4" class="ToolBar" align="center"> 
                  <input type="submit" name="Submit" value="<%= TEXT4 %>">
                  <input type="button" name="Submit1" value="<%= TEXT6 %>" onClick="SubmitForm('UpLoadMaterialForm', 'DeleteMaterialFile.asp?uid=<%=uid%>&lesson=<%= lesson %>&url=<%= url %>&mid=<%= mid %>')">
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
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

<%@ Language=jScript %>
<%  
 Response.Expires = -1;
%>
<!-- #include file='../js/user.inc' -->
<!-- #include file='inc/UpLoadFile.inc' -->
<%
  if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       
      		
%>
<html>
<head>
<title><%= TITLE1 %></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/EvalResult.css" type="text/css">
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

<form name=UpLoadFileForm action="UploadAndClasify.asp?uid=<%=uid%>" method="post" enctype="multipart/form-data">

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
                        &nbsp;<% var fcourse = Session("FullCourse"); Response.Write(fcourse.MaxFile) %>&nbsp;kb&nbsp;
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
              <tr> 
                <td width="100%" colspan="4" class="MessageTR1"><b><%= TEXT3 %>:</b></td>
              </tr>
              <tr> 
                <td width="100%" colspan="4" class="MessageTR1"> 
                  <textarea name="FileDescription" rows="8" cols="100"></textarea>
                </td>
              </tr>
              <tr> 
                <td colspan="4" class="ToolBar" align="center"> 
                  <input type="submit" name="Submit" value="<%= TEXT4 %>">
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

<%@ Language=jScript %>
<%
/*
    Esta página se usa para enviar un email a un usuario específico. 
    Especialmente diseñada para los Foros de discusión
    
    PARAMETROS
    ----------
    
    uid        -- El ID que genera la sesión cuando el usuario se conecta 
                  por primera vez.
                  
    UserTo     -- ID del usuario al que se le enviará en email
    
    UserNameTo -- Nombre completo del usuario al que se le enviará en email
    	              
    Course     -- ID del curso al que se le piden los temas de discusión.
    
    courseName -- Nombre del curso

  */

  Response.Expires = -1;
%>
<!-- #include file='../js/adolibrary.inc' -->
<!-- #include file='../js/user.inc' -->
<%  

  var uid = "";
  if (Request.QueryString.Item("uid").Count != 0) {
    uid = Request.QueryString.Item("uid") + "";
  }

  if (uid == Session("uid"))
    {    
      var UserTo = 0;
      if (Request.QueryString.Item("UserTo").Count != 0) {
        UserTo = Request.QueryString.Item("UserTo")
      }  

      var UserNameTo = 0;
      if (Request.QueryString.Item("UserNameTo").Count != 0) {
        UserNameTo = Request.QueryString.Item("UserNameTo")
      }  
      
      var CourseInUse = 0;
      if (Request.QueryString.Item("Course").Count != 0) {
        CourseInUse = Request.QueryString.Item("Course")
      }  

      var courseNameInUse = "";
      if (Request.QueryString.Item("courseName").Count != 0) {
        courseNameInUse = Request.QueryString.Item("courseName")
      }  
      
      //var valid = ((isUserInGroup(Session("UserID"),'Curso_' + courseNameInUse)) || (isUserInGroupByID(Session("UserID"),ADMIN_GROUP)) || (Session("permissionType") == ADMINISTRATOR));
      var valid = true;
      
      if (valid) {
        
      
%>
<script language=vbscript  runat=Server>
function getDate
 getDate = Now()
End function 
</script>
<!-- #include file='../js/user.inc' -->
<html>
<head>
<title>Nuevo Mensaje</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<script language=jscript src ="../js/library.js"></script>

<script language=jscript>
<!--
  function _Send()
   {  
     Mail.MailBody.value     = Mbody.value;
     Mail.MailSubject.value  = MSubject.value;
     Mail.MailPriority.value = MPriority.checked == true?1:0
     Mail.submit();      
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
                  <select id="Select" name="select1" class="ComboBox">
                    <option selected>
                      <%=UserNameTo%>
                    </option>
                  </select>
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
                  <input type="text" id="MSubject" name="MSubject" class="Edit" size="50" value="">
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td class="MessageTR1"> 
            <textarea name="Mbody" ID="Mbody" rows="12" class="TextArea"></textarea>
          </td>
        </tr>
        <tr> 
          <td class="MessageTR1"> 
            <table border="0" cellspacing="0" cellpadding="1">
              <tr> 
                <td> 
                  <input type="checkbox" ID="MPriority" name="MPriority" value="checkbox">
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
<FORM ID=Mail NAME=Mail METHOD=POST ACTION="SendForumMail.asp?uid=<%=uid%>" >
  <INPUT TYPE=hidden NAME=MailToText ID=MailToText VALUE="<%=UserTo%>">
  <INPUT TYPE=hidden NAME=MailBody ID=MailBody VALUE="">  
  <INPUT TYPE=hidden NAME=MailSubject ID=MailSubject VALUE="">
  <INPUT TYPE=hidden NAME=MailPriority ID=MailPriority VALUE=0>      
</FORM>
<%            
      }    
      else {
        Response.Redirect("errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);                   
      }
    }  
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>
</body>
</html>

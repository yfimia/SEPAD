<%@ Language=jScript %>
<%
/*
    Esta página se usa para enviar una nueva respuesta a un mensaje de
    los foros de discusión. 
    
    PARAMETROS
    ----------
    
    uid        -- El ID que genera la sesión cuando el usuario se conecta 
                  por primera vez.
                  
    msgid      -- Id del mensaje al que se le insertará una nueva opinión
                  
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
      var CourseInUse = 0;
      if (Request.QueryString.Item("Course").Count != 0) {
        CourseInUse = Request.QueryString.Item("Course")
      }  

      var pagenumber = 0;
      if (Request.QueryString.Item("pn").Count != 0) {
        pagenumber = Request.QueryString.Item("pn")
      }  

      var msgid = 0;
      if (Request.QueryString.Item("msgid").Count != 0) {
        msgid = Request.QueryString.Item("msgid")
      }  
      
      var TopicID = 0;
      if (Request.QueryString.Item("Topic").Count != 0) {
        TopicID = Request.QueryString.Item("Topic");
      }  

      var courseNameInUse = "";
      if (Request.QueryString.Item("courseName").Count != 0) {
        courseNameInUse = Request.QueryString.Item("courseName")
      }  
      
      
      if (Session("valid")) {
        var  filePath = Application("dataPath");   
        var  oConn    = Server.CreateObject("ADODB.Connection");
        var  oRec     = Server.CreateObject("ADODB.Recordset");
        oConn.Open(filePath);
        
        
        oRec.Open("SELECT ForumAnswers.ID AS AID, ForumMessages.ID AS IDMSG1, ForumMessages.AceptaRespuestas MActivo, ForumTopics.ID AS TID, ForumTopics.Activo AS TActivo " +
                  "FROM ForumTopics INNER JOIN (ForumMessages INNER JOIN ForumAnswers ON ForumMessages.ID = ForumAnswers.Origen) ON ForumTopics.ID = ForumMessages.Topic " +
                  "WHERE (ForumMessages.ID = " + msgid + ")",oConn,3,3);
        
        if ((oRec.EOF == true) || ((oRec.Fields.Item("MActivo").value == true) && (oRec.Fields.Item("TActivo").value == true))){
          oRec.Close(); 
%>
<script language=vbscript  runat=Server>
function getDate
 getDate = Now()
End function 
</script>
<!-- #include file='../js/user.inc' -->
<html>
<head>
<title>Nueva opinión</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<script language=jscript src ="../js/library.js"></script>

<script language=jscript>
<!--
  function _Send()
   {  
     Mail.MsgBody.value = Mbody.value;
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
          <td><a href="javascript:_Send()" class="ToolLink">&nbsp;Aceptar&nbsp;</a></td>
          <td>|</td>
          <td><a href="javascript:close()" class="ToolLink">&nbsp;Cancelar&nbsp;</a></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td class="MessageTR1"> 
            <textarea name="Mbody" ID="Mbody" rows="12" class="TextArea"></textarea>
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
<FORM ID=Mail NAME=Mail METHOD=POST ACTION="AceptaNewForumAnswer.asp?uid=<%=uid%>&course=<%=CourseInUse%>&Topic=<%=TopicID%>&msgid=<%=msgid%>&pn=<%=pagenumber%>&courseName=<%=courseNameInUse%>" >
  <INPUT TYPE=hidden NAME=msgid ID=msgid VALUE="<%=msgid%>">  
  <INPUT TYPE=hidden NAME=MsgBody ID=MsgBody VALUE="">  
</FORM>
</body>
</html>
<%            
          }
        else
          {
            Response.Redirect("errorpage.asp?tipo=Error&short='Operación no válida'&desc='El mensaje escogido no permite enviarle opiniones'");                   
          }
      }    
      else {
        Response.Redirect("errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);                   
      }
    }  
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>

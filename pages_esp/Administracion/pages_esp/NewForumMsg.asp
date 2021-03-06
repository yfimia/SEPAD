<%@ Language=jScript %>
<%
/*
    Esta p�gina se usa para enviar un nuevo mensaje a los foros de 
    discusi�n. 
    
    PARAMETROS
    ----------
    
    uid        -- El ID que genera la sesi�n cuando el usuario se conecta 
                  por primera vez.
                  
    TopicID    -- Id del tema al que se le insertar� el nuevo mensaje              
                  
    Course     -- ID del curso al que se le piden los temas de discusi�n.
    
    courseName -- Nombre del curso
    
    pagenumber -- numero del bloque o pagina que se esta mostrando

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

      var TopicID = 0;
      if (Request.QueryString.Item("Topic").Count != 0) {
        TopicID = Request.QueryString.Item("Topic")
      }  
      
      var pagenumber = 0;
      if (Request.QueryString.Item("pn").Count != 0) {
        pagenumber = Request.QueryString.Item("pn")
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

        //Chequeo si esta habilitado el tema
        oRec.Open("SELECT * FROM ForumTopics WHERE (ID = " + TopicID + ") and (Activo = " + Application("dtrue") + ")",oConn,3,3);
        if (oRec.EOF == false) {
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
<title>Nuevo Mensaje</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<script language=jscript src ="../js/library.js"></script>

<script language=jscript>
<!--
  function _Send()
   {  
     Mail.MsgBody.value = Mbody.value;
     Mail.MsgAcepta.value  = MAcepta.checked == true?1:0
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
        <tr> 
          <td class="MessageTR1"> 
            <table border="0" cellspacing="0" cellpadding="1">
              <tr> 
                <td> 
                  <input type="checkbox" ID="MAcepta" name="MAcepta" value=0>
                </td>
                <td>Aceptar respuestas.</td>
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
<FORM ID=Mail NAME=Mail METHOD=POST ACTION="AceptaNewForumMsg.asp?uid=<%=uid%>&course=<%=CourseInUse%>&Topic=<%=TopicID%>&pn=<%=pagenumber%>&courseName=<%=courseNameInUse%>" >
  <INPUT TYPE=hidden NAME=MsgTopic ID=MsgTopic VALUE="<%=TopicID%>">  
  <INPUT TYPE=hidden NAME=MsgBody ID=MsgBody VALUE="">  
  <INPUT TYPE=hidden NAME=MsgAcepta ID=MsgAcepta VALUE=0>      
</FORM>
</body>
</html>
<%            
        }
        else{
          Response.Redirect("errorpage.asp?tipo=Error&short=Mensaje inactivo" + "&desc=El mensaje seleccionado no est� activado para el uso" );                   
        }
        oConn.Close();
      }    
      else {
        Response.Redirect("errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);                   
      }
    }  
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>

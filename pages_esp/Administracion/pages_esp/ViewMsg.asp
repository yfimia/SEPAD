<%@ Language=jScript %>

<%Response.Expires = -1%>
<!-- #include file='../js/adolibrary.inc' -->
<!-- #include file='../js/user.inc' -->
<!-- #include file='inc/ViewMsg.inc' -->


<%
   
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
    
   var Ind;  
   if (Request.QueryString.Count  > 0 ) {
     } 
      var uid = Request.QueryString.Item("uid") + "";       
//      MarkReaded( Request.QueryString.Item("ID") );
      

      var  filePath = Application("dataPath");   
      var  oConn    = Server.CreateObject("ADODB.Connection");
      var  oRec     = Server.CreateObject("ADODB.Recordset");
      oConn.Open(filePath);
      oRec.Open("SELECT * from Mensajeria Where (ID = " + Request.QueryString.Item("ID") + ")",oConn,3,3);    
%>

<script language=vbscript  runat=Server>
function getDate
getDate = Now()
End function 
</script>

<script language=jscript>
  function reply() {
<%  
    if (oRec.Fields("FromID").Value + "" != "null") {
%>    

      Mail.action = "NewMsg.asp?uid=<%=uid%>";
      Mail.MailToText.value   = "<%=oRec.Fields("FromID").Value%>";
      Mail.MailToName.value   = "<%=oRec.Fields("From").Value%>";
      Mail.MailBody.value     = Texto.innerText;      
      Mail.MailSubject.value  = "<%=oRec.Fields("Subject").Value%>";
      Mail.MailPriority.value = "<%=oRec.Fields("Priority").Value == true?1:0%>";
      Mail.submit();     
<%       
    }  
    else 
      {
%>    
      alert('<%= TEXT1 %>.');       
<%   
      }      
%>      
  }

  function forward() {
      Mail.action = "NewMsg.asp?uid=<%=uid%>";
      Mail.MailToText.value   = "-1";
      Mail.MailToName.value   = "";
      Mail.MailBody.value     = Texto.innerText;      
      Mail.MailSubject.value  = "<%=oRec.Fields("Subject").Value%>";
      Mail.MailPriority.value = "<%=oRec.Fields("Priority").Value == true?1:0%>";
      Mail.submit();     
  }
  
</script>

<html>
<head>
<title><%=oRec.Fields("Subject").Value%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<script language=jscript src ="../js/library.js"></script>


</head>

<body bgcolor="#FFFFFF" text="#000000">
      
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="MessageBarTD" align="right"> 
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td><a href="javascript:close()" class="ToolLink">&nbsp;<%= TEXT2 %>&nbsp;</a></td>
          <td>&nbsp;|&nbsp;</td>          
          <td><a href="javascript:reply()" class="ToolLink">&nbsp;<%= TEXT3 %>&nbsp;</a></td>                  
          <td>&nbsp;|&nbsp;</td>          
          <td><a href="javascript:forward()" class="ToolLink">&nbsp;<%= TEXT4 %>&nbsp;</a></td>                  
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
                <td width="1%" valign="top"><b><%= TEXT5 %>:&nbsp;&nbsp;</b></td>
                <td width="99%"><%=oRec.Fields("From").Value %></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td class="MessageTR1"> 
            <table border="0" cellspacing="0" cellpadding="0" width="100%">
              <tr> 
                <td width="1%" valign="top"><b><%= TEXT6 %>:&nbsp;&nbsp;</b></td>
                <td width="99%"><%=oRec.Fields("ToText").Value %></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td class="MessageTR1"> 
            <table border="0" cellspacing="0" cellpadding="0" width="100%">
              <tr> 
                <td width="1%" valign="top"><b><%= TEXT7 %>:&nbsp;&nbsp;</b></td>
                <td width="99%"><%=oRec.Fields("Subject").Value %></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td class="MessageTR1"> 
            <textarea name="content" ID=Texto rows="12" class="TextArea"><%=oRec.Fields("Content").Value %></textarea>
          </td>
        </tr>
        <tr> 
          <td class="MessageTR1">&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td class="MessageBarTD" align="right"> 
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td><a href="javascript:close()" class="ToolLink">&nbsp;<%= TEXT2 %>&nbsp;</a></td>
          <td>&nbsp;|&nbsp;</td>          
          <td><a href="javascript:reply()" class="ToolLink">&nbsp;<%= TEXT3 %>&nbsp;</a></td>                  
          <td>&nbsp;|&nbsp;</td>          
          <td><a href="javascript:forward()" class="ToolLink">&nbsp;<%= TEXT4 %>&nbsp;</a></td>                  
        </tr>
      </table>
    </td>
  </tr>
</table>
<FORM ID=Mail NAME=Mail METHOD=POST ACTION="SendMail.asp?uid=<%=uid%>" >
  <INPUT TYPE=hidden NAME=MailToText ID=MailToText VALUE="">
  <INPUT TYPE=hidden NAME=MailToName ID=MailToName VALUE="">
  <INPUT TYPE=hidden NAME=MailBody ID=MailBody VALUE="">  
  <INPUT TYPE=hidden NAME=MailSubject ID=MailSubject VALUE="">
  <INPUT TYPE=hidden NAME=MailPriority ID=MailPriority VALUE=0>      
</FORM>
</body>
</html>
<%
      oRec.Fields("Readed").Value = true;
      oRec.Update();

      oRec.Close();

      oConn.Close();

  }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
  //Response.Write(Request.QueryString.Item("uid"));
%>

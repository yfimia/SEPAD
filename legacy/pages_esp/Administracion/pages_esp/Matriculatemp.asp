<%@ Language=JScript %>

<%
  Response.Expires = -1;
%>
<!-- #include file="../js/user.inc" -->
<%  
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

  

%>


    
<HTML>
<HEAD>
<link REL="stylesheet" TYPE="text/css" HREF="../css/<%=Session("skin")%>/main.css" />
<title>Solicitud de matrícula para <%=Session("courseName")%></title>
</HEAD>
<BODY style="overflow-y:hidden">
<%
      var  filePath = Application("dataPath");   
 			var  oConn    = Server.CreateObject("ADODB.Connection");
      var  oRec     = Server.CreateObject("ADODB.Recordset");

      oRec.CursorLocation = 3;
      oRec.CursorType     = 3;
      oConn.Open(filePath);

      oRec.Open("SELECT * FROM TMP_Matriculas_Cursos where ([User] = " + Session("UserID") + ') and (Modulo = '  + Session("modulo") + ')',oConn,3,3); 
      if (oRec.EOF) {
        oRec.AddNew();
        oRec.Fields.Item("User").Value = Session("UserID");
        oRec.Fields.Item("Modulo").Value = Session("modulo");
        oRec.Update();
%>
<table border="0" cellspacing="0" cellpadding="3" class="MessageTR1" width="100%">
  <tr>
    <td class="Header4">
    <br>
        Su matrícula ha sido procesada, si es aceptado se 
        le notificará mediante la mensajería interna del sistema.
    <br>
    <br>
    </td>
  </tr>
</table>
<table border="0" cellspacing="0" cellpadding="3" class="ToolBar" width="100%">
  <tr>
    <td>
      <center>
        <table border="0" cellspacing="0" cellpadding="3" class="ToolBar">
          <tr>
            <td>    
              <a href="javascript:close()" class="ToolLink">&nbsp;Cerrar&nbsp;</a>
            </td>
          </tr>
        </table>
      </center>      
    </td>
  </tr>
</table>

<%        
      }  
    else
      {
%>
<table border="0" cellspacing="0" cellpadding="3" class="MessageTR1" width="100%">
  <tr>
    <td class="Header4">
    <br>
        Ya existe una solictud suya que aún no ha sido procesada, si es aceptado se 
        le notificará mediante la mensajería interna del sistema.
    <br>
    <br>
    </td>
  </tr>
</table>
<table border="0" cellspacing="0" cellpadding="3" class="ToolBar" width="100%">
  <tr>
    <td>
      <center>
        <table border="0" cellspacing="0" cellpadding="3" class="ToolBar">
          <tr>
            <td>    
              <a href="javascript:close()" class="ToolLink">&nbsp;Cerrar&nbsp;</a>
            </td>
          </tr>
        </table>
      </center>      
    </td>
  </tr>
</table>

<%      
      }
      oRec.Close()
      oConn.Close();   
%>      
</BODY>
</HTML>
<%
 }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>

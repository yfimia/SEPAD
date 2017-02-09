<%@ Language=JScript %>
<%
/*
    Esta página muestra la información de un usuario dado.
    
    PARAMETROS
    ----------
    
    uid        -- El ID que genera la sesión cuando el usuario se conecta 
                  por primera vez.
                  
    userid     -- ID del usuario del cual se desea ver la información.
    
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
      var userid = -1;
      if (Request.QueryString.Item("userid").Count != 0) {
        userid = Request.QueryString.Item("userid")
      }  

      var valid = (Session("UserID") != GUEST_USER);
      
      if (valid) {
        var  filePath = Application("dataPath");   
        var  oConn    = Server.CreateObject("ADODB.Connection");
        var  oComm    = Server.CreateObject("ADODB.Command");
        var  oRec     = Server.CreateObject("ADODB.Recordset");
        oRec.CursorLocation = 3;
        oRec.CursorType     = 3;
        oConn.Open(filePath);
        oComm.ActiveConnection = oConn;
        oComm.CommandText = "SELECT * FROM Usuarios Where ID = " + userid;
        oRec = oComm.Execute();    
%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<Title>Datos del autor</Title>
</HEAD>
<body bgcolor="#ffffff" text="#000000">

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar" align="middle"><h1><b>Datos del autor</b></h1></td>
  </tr>
</table> 
<br>
<table width="98%" border="0" cellspacing="0" cellpadding="0" align=center>
  <tr>
    <td colspan=2 class="ToolBar" align="middle">
      &nbsp;
    </td>       
  </tr>
<%if (!oRec.EOF){ %>  
  <tr>
    <td class=MessageTR1>
      Nombre y apellidos
    </td>       
    <td class=MessageTR1>
      <%=oRec.Fields.Item("Fullname").value%>
    </td>       
  </tr>
  <tr>
    <td class=MessageTR>
      Nick
    </td>       
    <td class=MessageTR>
      <%=oRec.Fields.Item("name").value%>
    </td>       
  </tr>
  <tr>
    <td class=MessageTR1>
      Email
    </td>       
    <td class=MessageTR1>
      <%=oRec.Fields.Item("email").value%>
    </td>       
  </tr>
  <tr>
    <td class=MessageTR>
      Tipo de usuario
    </td>       
    <td class=MessageTR>
      <%if (oRec.Fields.Item("permissiontype").value == USER)
          Response.Write(USER_TEXT)
        else
          if (oRec.Fields.Item("permissiontype").value == PUBLICATOR)  
            Response.Write(PUBLICATOR_TEXT)
          else  
            if (oRec.Fields.Item("permissiontype").value == ADMINISTRATOR)  
              Response.Write(ADMINISTRATOR_TEXT)
            else
              Response.Write(GUEST);
      %>
    </td>       
  </tr>
  
<%}
  else{
%>  
  <tr>
    <td colspan=2 class="MessageTR" align="middle">
      Usuario no válido
    </td>       
  </tr>
<%}%>  
  <tr>
    <td colspan=2 class="ToolBar" align="middle">
      &nbsp;
    </td>       
  </tr>
</table>
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
</HTML>

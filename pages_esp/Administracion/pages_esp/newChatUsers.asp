<%@ Language=JScript %>
<%
 Response.Expires = -1;
 ChatRefreshTime = 5000;
%>
<!-- #include file='../js/user.inc' -->
<!-- #include file='../js/chat.inc' -->
<%
  sala = -1;
  if (Request.QueryString.Item("salaid").count != 0){
    sala = Request.QueryString.Item("salaid");
  }
  var  filePath = Application("dataPath");   
  var  oConn    = Server.CreateObject("ADODB.Connection");
  var  oRec     = Server.CreateObject("ADODB.Recordset");
  oRec.CursorLocation = 3;
  oRec.CursorType     = 3;
  
  oConn.Open(filePath);

  //Verifico si es valida la sala
  oRec.Open("SELECT * FROM Salas WHERE (ID = " + sala + ")", oConn, 3, 3);    
  salapublica = false;
  if (oRec.EOF == false){
    nameofsala = oRec.Fields.Item("Nombre").value;
    salapublica = oRec.Fields.Item("Publica").value;
    salamoderador = oRec.Fields.Item("Moderador").value;
    oRec.Close();
    
    //Verifico si el usuario actual puede entrar a la sala
    oRec.Open("SELECT * FROM ChatCanEnter WHERE (Sala = " + sala + ") and (Usuario = " + Session("userID") + ")", oConn, 3, 3);     
    if ((salapublica) || (salamoderador == Session("userID")) || (oRec.EOF == false) || (Session("PermissionType") == ADMINISTRATOR)){
      oRec.Close();
    
      //
      ya = false;
      oRec.Open("SELECT * FROM ChatOnlineUsers " +
                "WHERE (Usuario = " + Session("userID") + ") and (Sala = " + sala + ")", oConn, 3, 3);    
      if (oRec.EOF == false){
        ya = true;
      }          
      oRec.Close();
  
      if (ya == false){
        oRec.Open("SELECT * From ChatOnlineUsers", oConn, 3, 3);    
        oRec.AddNew();
        oRec.Fields.Item("Usuario").value  = Session("userID");
        oRec.Fields.Item("Sala").value     = sala;
        oRec.Update();
        oRec.Close(); 
      } 
      //

      //Actualizo los datos del usuario actual
     if (Application("SQL") == false) {
      
       oRec.Open("UPDATE ChatOnlineUsers SET Fecha = Now() WHERE (Usuario = " + Session("userID") + ") and (Sala = " + sala + ")", oConn, 3, 3);    
       oRec.Open("DELETE FROM ChatOnlineUsers WHERE (DateDiff('n', ChatOnlineUsers.Fecha,Now()) > 10)", oConn, 3, 3);    
       //DateDiff('n', ChatOnlineUsers.Fecha,Now()) As TimeElapsed,
       oRec.Open("SELECT ChatOnlineUsers.Usuario,ChatOnlineUsers.Sala,Usuarios.Name,Usuarios.Sexo " + 
                 "FROM Usuarios INNER JOIN ChatOnlineUsers ON Usuarios.ID = ChatOnlineUsers.Usuario " + 
                 "WHERE (ChatOnlineUsers.Sala=" + sala + ")", oConn, 3, 3);    

     }
     else {
       oRec.Open("UPDATE ChatOnlineUsers SET Fecha = GETDATE() WHERE (Usuario = " + Session("userID") + ") and (Sala = " + sala + ")", oConn, 3, 3);    
       oRec.Open("DELETE FROM ChatOnlineUsers WHERE (DATEDIFF(minute, ChatOnlineUsers.Fecha,GETDATE()) > 10)", oConn, 3, 3);    
       //DATEDIFF(Minute, ChatOnlineUsers.Fecha,GETDATE()) As Timeelapsed,
       oRec.Open("SELECT ChatOnlineUsers.Usuario, Usuarios.Name,Usuarios.Sexo, ChatOnlineUsers.Sala " + 
                 "FROM Usuarios INNER JOIN ChatOnlineUsers ON Usuarios.ID = ChatOnlineUsers.Usuario " + 
                 "WHERE (ChatOnlineUsers.Sala=" + sala + ")", oConn, 3, 3);    
     
     }  
  
%>

<HTML>
<HEAD>
  <title>Sala de charlas</title>
  <link rel="stylesheet" href="../css/<%=Session("skin")%>/Chat.css" type="text/css">
  <script type="text/javascript" language="javascript" src="../js/Windows.js"></script>
</HEAD>


<BODY>
  <table width="100%" border=0 cellpadding=0 cellspacing=0>
    <tr  height=15  valign=middle>
      <td colspan=2  class="Toolbar"  valign=top>
        <b>[Sala : <%=nameofsala%> (<font color=lightGreen><%=oRec.RecordCount%></font>) ]</b>
      </td>
    </tr>
    <%
      while (oRec.EOF == false){
    %>
    <tr>
      <td width="1%">
        <%
          if (oRec.Fields.Item("Sexo").value){
        %>
        <img border=0 src="../images/<%=Session("skin")%>/chat/male.gif">      
        <%}
          else{
        %>
        <img border=0 src="../images/<%=Session("skin")%>/chat/female.gif">      
        <%}%>
      </td>
      <td  align=left>
        <%
          if (oRec.Fields.Item("Usuario").value == Session("userID") + ""){
        %>
        [<b><font color=red><%=oRec.Fields.Item("name").value%></font></b>]
        <%
        }
        else{
        %>
        [<b><a href="javascript:abreVentana( 'SalaPrivadaDe<%=Session("userID")%>Para<%=oRec.Fields.Item("name").value%>', 800, 100, 'newChatPrivateMessanger.asp?salaid=<%=sala%>&userid=<%=oRec.Fields.Item("Usuario").value%>');"><%=oRec.Fields.Item("name").value%></a></b>] 
        <%
        }
        %>
      </td>
    </tr>
    <%oRec.Move(1);
      }
      oRec.Close();
    %>
  </table>
  
  
<script language=jscript>
  function actualiza(){
    location="newChatUsers.asp?salaid=<%=sala%>";
    window.setTimeout("actualiza()",<%=ChatRefreshTime%>)
  }
  
  window.setTimeout("actualiza()",<%=ChatRefreshTime%>)
</script>
  
</BODY>
</HTML>
<%  }
    else{
      Response.Redirect("errorpage.asp?tipo=Error&short=El usuario no tiene suficientes privilegios." + "&desc=" + INVALID_ROOM_USER);
    }
  }
  else{
    Response.Redirect("errorpage.asp?tipo=Error&short=Error con la sala seleccionada." + "&desc=" + INVALID_ROOM);
  }
  oConn.Close();
%>

<%@ Language=JScript %>
<%
 Response.Expires = -1;
 ChatRefreshTime = Session("ChatTimeToRefresh") * 1000;
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

  //Verifico si la sala es valida
  oRec.Open("SELECT * FROM Salas WHERE (ID = " + sala + ")", oConn, 3, 3);    
  salapublica = false;
  if (oRec.EOF == false){
    nameofsala = oRec.Fields.Item("Nombre").value;
    salapublica = oRec.Fields.Item("Publica").value;
    salamoderador = oRec.Fields.Item("Moderador").value;
    oRec.Close();
    
    //si me envian la opcion a efectuar
    if (Request.QueryString.Item("option").count != 0){
      opt = Request.QueryString.Item("option");
      
      //Casos por ahora solo la opcion 3
      if (opt == 3) {
        //Si tiene permiso para efectuar las operaciones
        if ((salamoderador == Session("userID")) || (Session("PermissionType") == ADMINISTRATOR)) {
          //si me envian el id del mensaje
          if (Request.QueryString.Item("msgid").count != 0){
            mymsgid = Request.QueryString.Item("msgid");
            oRec.Open("DELETE FROM ChatMenssages WHERE ID = " + mymsgid, oConn, 3, 3);    
          }  
        }
      }
    }  
       
    //Verifico si el usuario actual puede entrar a la sala
    oRec.Open("SELECT * FROM ChatCanEnter WHERE (Sala = " + sala + ") and (Usuario = " + Session("userID") + ")", oConn, 3, 3);     
    if ((salapublica) || (salamoderador == Session("userID")) || (oRec.EOF == false) || (Session("PermissionType") == ADMINISTRATOR)){
      oRec.Close();

     if (Application("SQL") == false) {
  
      oRec.Open("SELECT TOP " + Application("TOPMSG") + " ChatMenssages.ID, ChatMenssages.Texto, ChatMenssages.Color, ChatMenssages.Negrita, ChatMenssages.Cursiva, ChatMenssages.Subrayada, ChatMenssages.Publico, ChatMenssages.UserOrigen, ChatMenssages.Sala, Hour(ChatMenssages.FechaHora) AS HoraA,Minute(ChatMenssages.FechaHora) AS MinuteA,Second(ChatMenssages.FechaHora) AS SecondA,Caritas.Imagen,ChatMenssages.Carita,Usuarios.Name " +
                "FROM Usuarios INNER JOIN (Caritas RIGHT JOIN ChatMenssages ON Caritas.ID = ChatMenssages.Carita) ON Usuarios.ID = ChatMenssages.UserOrigen " +
                "WHERE (ChatMenssages.sala = " + sala + ") and (ChatMenssages.Publico = " + Application("dtrue") + ") " +
                "ORDER BY ChatMenssages.ID DESC", oConn, 3, 3);
     } 
     else {
      oRec.Open("SELECT TOP " + Application("TOPMSG") + " ChatMenssages.ID, ChatMenssages.Texto, ChatMenssages.Color, ChatMenssages.Negrita, ChatMenssages.Cursiva, ChatMenssages.Subrayada, ChatMenssages.Publico, ChatMenssages.UserOrigen, ChatMenssages.Sala, DATEPART(hour, ChatMenssages.FechaHora ) AS HoraA, DATEPART(minute, ChatMenssages.FechaHora) AS MinuteA, DATEPART(second, ChatMenssages.FechaHora) AS SecondA, Caritas.Imagen,ChatMenssages.Carita,Usuarios.Name " +
                "FROM Usuarios INNER JOIN (Caritas RIGHT JOIN ChatMenssages ON Caritas.ID = ChatMenssages.Carita) ON Usuarios.ID = ChatMenssages.UserOrigen " +
                "WHERE (ChatMenssages.sala = " + sala + ") and (ChatMenssages.Publico = " + Application("dtrue") + ") " +
                "ORDER BY ChatMenssages.ID DESC", oConn, 3, 3);
     } 


%>

<HTML>
<HEAD>
  <title>Sala de charlas</title>
  <link rel="stylesheet" href="../css/<%=Session("skin")%>/chat.css" type="text/css">
  <script type="text/javascript" language="javascript" src="../js/Windows.js"></script>
</HEAD>
<BODY>
  <table width=100% border=0 cellspacing=1 cellpadding=1>
    <tr height=15  valign=middle>
      <td colspan=6  class="Toolbar"  valign=top>
        <table height=10 width="100%"><tr><td><b>[ <u>Mensajes</u> <u>públicos</u> ]</b></td><td align=left><a href="newChatSavePublicMsg.asp?salaid=<%=sala%>" target=_blank><img alt="Salvar conversación" title="Salvar conversación" border=0 height=15 src="disck.gif"></a></td></tr></table>
      </td>
    </tr>
<%
    while (oRec.EOF == false){
%>  
    <tr>
      <td width="1%">
        <small>
          <%=oRec.Fields.Item("HoraA").value%>:<%=oRec.Fields.Item("MinuteA").value%>:<%=oRec.Fields.Item("SecondA").value%>
        </small>
      </td>
      <td width="0%">
        [     
      </td>
      <td width="1%">
        <%
          if (oRec.Fields.Item("UserOrigen").value == Session("userID") + ""){
        %>
        <b><%=oRec.Fields.Item("name").value%></b>
        <%}
          else{
        %>
        <b><a href="javascript:abreVentana( 'PMD<%=Session("userID")%>para<%=oRec.Fields.Item("UserOrigen").value%>', 800, 100, 'newChatPrivateMessanger.asp?salaid=<%=sala%>&userid=<%=oRec.Fields.Item("UserOrigen").value%>');"><%=oRec.Fields.Item("name").value%></a></b>      
        <%}
        %>
      </td>
      <td width="0%">
        ]     
      </td>
      <td width="1%" align=center  valign=middle>
        <%if (oRec.Fields.Item("Carita").value == 0){%>
          <img src="../images/<%=Session("skin")%>/chat/emotions/none.gif">          
        <%}
        else{
        %>
          <img src="../images/<%=Session("skin")%>/chat/emotions/<%=oRec.Fields.Item("Imagen").value%>">
        <%}%>  
      </td>
      <td width="97%">
        <%
          str = oRec.Fields.Item("Texto").value;
          if (oRec.Fields.Item("Subrayada").value){
            str = "<u>" + str + "</u>";
          }  
          if (oRec.Fields.Item("Cursiva").value){
            str = "<i>" + str + "</i>";
          }  
          if (oRec.Fields.Item("Negrita").value){
            str = "<b>" + str + "</b>";
          }  
          str = "<font color=" + oRec.Fields.Item("Color").value + ">" + str + "</font>"

          Adminoption = "";
          if ((salamoderador == Session("userID")) || (Session("PermissionType") == ADMINISTRATOR)) {
            Adminoption = "<a href=newChatPublicMsg.asp?salaid=" + sala + "&option=3&msgid=" + oRec.Fields.Item("ID").value + "><img Alt='Borrar mensaje' title='Borrar mensaje' border=0 src='../images/deletemsg.gif'></a>"
          }
          
          str = str + Adminoption;

          oRec.Move(1);       


        %>
        <span>
          <%=str%>
        </span>
      </td>  
    </tr>  
<%
    }
  oRec.Close(); 
%>  
  </table>
</BODY>
</HTML>
<script language=jscript>
  function actualiza(){
    location = "newChatPublicMsg.asp?salaid=<%=sala%>";
    window.setTimeout("actualiza()",<%=ChatRefreshTime%>);
  }
  window.setTimeout("actualiza()",<%=ChatRefreshTime%>)
</script>
<%  }
    else{
      oRec.Close();
      oConn.Close();
      Response.Redirect("errorpage.asp?tipo=Error&short=El usuario no tiene suficientes privilegios." + "&desc=" + INVALID_ROOM_USER);
    }//Puede o no puede entrar a la sala
  }
  else{
    oRec.Close();
    oConn.Close();
    Response.Redirect("errorpage.asp?tipo=Error&short=Error con la sala seleccionada." + "&desc=" + INVALID_ROOM);
  }//Sala valida o no
  oConn.Close();
%>

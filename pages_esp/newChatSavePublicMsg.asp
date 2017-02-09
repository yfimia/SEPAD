<%@ Language=JScript %>
<%
 Response.Expires = -1;
 Response.AddHeader ("Content-Type", "text/html");
 Response.AddHeader ("Content-transfer-encoding", "binary");
 var newDateObj = new Date();
 Response.AddHeader ("Content-Disposition", "attachment; filename=chatpublico" + newDateObj.getDate() + "_" + newDateObj.getMonth() + "_" + newDateObj.getYear() + ".htm");
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
       
    //Verifico si el usuario actual puede entrar a la sala
    oRec.Open("SELECT * FROM ChatCanEnter WHERE (Sala = " + sala + ") and (Usuario = " + Session("userID") + ")", oConn, 3, 3);     
    if ((salapublica) || (salamoderador == Session("userID")) || (oRec.EOF == false) || (Session("PermissionType") == ADMINISTRATOR)){
      oRec.Close();

     if (Application("SQL") == false) {
  
      oRec.Open("SELECT TOP " + Application("TOPMSG") + " ChatMenssages.ID, ChatMenssages.Texto, ChatMenssages.Color, ChatMenssages.Negrita, ChatMenssages.Cursiva, ChatMenssages.Subrayada, ChatMenssages.Publico, ChatMenssages.UserOrigen, ChatMenssages.Sala, Hour(ChatMenssages.FechaHora) AS HoraA,Minute(ChatMenssages.FechaHora) AS MinuteA,Second(ChatMenssages.FechaHora) AS SecondA,Caritas.Texto As TextoCarita,ChatMenssages.Carita,Usuarios.Name " +
                "FROM Usuarios INNER JOIN (Caritas RIGHT JOIN ChatMenssages ON Caritas.ID = ChatMenssages.Carita) ON Usuarios.ID = ChatMenssages.UserOrigen " +
                "WHERE (ChatMenssages.sala = " + sala + ") and (ChatMenssages.Publico = " + Application("dtrue") + ") " +
                "ORDER BY ChatMenssages.ID DESC", oConn, 3, 3);
     } 
     else {
      oRec.Open("SELECT TOP " + Application("TOPMSG") + " ChatMenssages.ID, ChatMenssages.Texto, ChatMenssages.Color, ChatMenssages.Negrita, ChatMenssages.Cursiva, ChatMenssages.Subrayada, ChatMenssages.Publico, ChatMenssages.UserOrigen, ChatMenssages.Sala, DATEPART(hour, ChatMenssages.FechaHora ) AS HoraA, DATEPART(minute, ChatMenssages.FechaHora) AS MinuteA, DATEPART(second, ChatMenssages.FechaHora) AS SecondA, Caritas.Texto As TextoCarita,ChatMenssages.Carita,Usuarios.Name " +
                "FROM Usuarios INNER JOIN (Caritas RIGHT JOIN ChatMenssages ON Caritas.ID = ChatMenssages.Carita) ON Usuarios.ID = ChatMenssages.UserOrigen " +
                "WHERE (ChatMenssages.sala = " + sala + ") and (ChatMenssages.Publico = " + Application("dtrue") + ") " +
                "ORDER BY ChatMenssages.ID DESC", oConn, 3, 3);
     } 
%>
<HTML>
<HEAD>
  <title>Sala de charlas: <%Response.Write(newDateObj.getDate() + "-" + newDateObj.getMonth() + "-" + newDateObj.getYear());%></title>
  <style>
    BODY
      {   
        FONT-FAMILY: Verdana;
        FONT-SIZE: xx-small;
        MARGIN: 0px;
        background-color : #FFFAF0;
        color : Black;
      }
    TABLE
      {
        FONT-FAMILY: Verdana;
        FONT-SIZE: xx-small
      }

    .ToolBar
      {
        BACKGROUND-COLOR: #09a99d;
        COLOR: #ffffff;
        FONT-FAMILY: Verdana;
        FONT-SIZE: xx-small;
        MARGIN-LEFT: 0px;
        MARGIN-RIGHT: 0px;
        PADDING-LEFT: 2px;
        PADDING-RIGHT: 2px
      }
  </style>
</HEAD>
<BODY>
  <table width=100% border=0 cellspacing=1 cellpadding=1>
    <tr height=15  valign=middle>
      <td colspan=5  class="Toolbar"  valign=top>
        <table height=10 width="100%"><tr><td><b>[ <u>Mensajes</u> <u>públicos</u> ]</b><%Response.Write(newDateObj.getDate() + "-" + newDateObj.getMonth() + "-" + newDateObj.getYear());%></td></tr></table>
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
        <b><%=oRec.Fields.Item("name").value%></b>      
        <%}
        %>
      </td>
      <td width="0%">
        ]     
      </td>
      <td width="97%">
        <%if (oRec.Fields.Item("Carita").value != 0){%>
          [<%=oRec.Fields.Item("TextoCarita").value%>]
        <%}%>  
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

<%@ Language=JScript %>
<!-- #include file='../js/user.inc' -->
<%
  Response.Expires = -1;
 
  pagenumber = 0;
  if (Request.QueryString.Item("pn").count != 0){
    pagenumber = Request.QueryString.Item("pn");
  }

  var  filePath = Application("dataPath");   
  var  oConn    = Server.CreateObject("ADODB.Connection");
  var  oRec     = Server.CreateObject("ADODB.Recordset");
  
  oConn.Open(filePath);

  
  saladelete = -2;
  
  operation = 0;
  if (Request.QueryString.Item("op").count != 0){
    operation = Request.QueryString.Item("op");
  }
  
  /*
  Lista de operaciones
  --------------------
  0 - No operacion
  
  1 - Eliminar
        sd - trae el ID de la sala a borrar
  */
  
  if (operation == 1){
      saladelete = -1;
      if (Request.QueryString.Item("sd").count != 0){
        saladelete = Request.QueryString.Item("sd");
        if (Session("PermissionType") == ADMINISTRATOR){
          oRec.Open("DELETE FROM Salas WHERE ID = " + saladelete , oConn, 3, 3);    
        }
        else{
          oRec.Open("DELETE FROM Salas WHERE (ID = " + saladelete + ") and (Moderador = " + Session("userID") + ")", oConn, 3, 3);    
        }  
      }  
  }
  
  //Actualizo los usuarios online
  if (Application("SQL") == false) {
    oRec.Open("DELETE FROM ChatOnlineUsers WHERE (DateDiff('n', ChatOnlineUsers.Fecha,Now()) > 10)", oConn, 3, 3);    
  }
  else {
    oRec.Open("DELETE FROM ChatOnlineUsers WHERE (DATEDIFF(minute, ChatOnlineUsers.Fecha,GETDATE()) > 10)", oConn, 3, 3);    
  }  
  
  oRec.Open("SELECT Salas.ID,Salas.Objetivo, Salas.Moderador,Salas.Publica, Salas.Nombre, count(*) As Cantidad " + 
            "FROM Salas, ChatOnlineUsers " + 
            "WHERE Salas.ID = ChatOnlineUsers.Sala " + 
            "GROUP BY Salas.ID,Salas.Objetivo,Salas.Moderador, Salas.Publica,Salas.Nombre " + 
            "UNION SELECT Salas.ID, Salas.Objetivo,Salas.Moderador, Salas.Publica, Salas.Nombre, 0 As Cantidad " +  
            "FROM Salas " + 
            "WHERE NOT (Salas.ID IN (SELECT sala FROM ChatOnlineUsers))", oConn, 3, 3);    
            
%>
<HTML>
<HEAD>
  <title>Listado de las sala de charlas</title>
  <link rel="stylesheet" href="../css/<%=Session("skin")%>/EvalResult.css" type="text/css">
  <link rel="stylesheet" href="../css/<%=Session("skin")%>/Chat.css" type="text/css">
  <script type="text/javascript" language="javascript" src="../js/Windows.js"></script>
</HEAD>

<body text="#000000" style="overflow-y:auto">
  <table width="100%" border="0" cellspacing="0" cellpadding="2">
    <tr> 
      <td class="ToolBar">
        <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
          <tr> 
            <td><img src="../images/<%=Session("skin")%>/ChatIMG.gif" width="80" height="54"></td>
            <td class="HeaderTable">Lista de salas del Chat</td>
          </tr>
        </table>
      </td>
    </tr>
  </table>  
  <table border="0" cellspacing="0" cellpadding="0" class="ToolBar" width="100%">
    <tr>
      <td>
        <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
          <tr> 
            <td nowrap>
              <a href="newChatAddRoom.asp" class="ToolLink">
                &nbsp;Agregar nueva sala&nbsp;
              </a>
            </td>
            <td>|</td>
            <td>
              <a href="newChatRoomList.asp?pn=<%=pagenumber + 1%>" class="ToolLink">
                &nbsp;Próximas&nbsp;<%=SHOW_CANT%>&nbsp;
              </a>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td> 
        <table width="100%" border="0" cellspacing="1" cellpadding="0">
          <tr> 
            <td width="1%" class="ToolBar" align="center"></td>         
            <td width="1%" class="ToolBar" align="center"></td>
            <td width="1%" class="ToolBar" align="center">
              <b>Usuarios<br>Online</b>
            </td>
            <td width="24%" class="ToolBar" align="center">
              <b>Nombre</b>
            </td>
            <td width="73%" class="ToolBar" align="center">
              <b>Objetivo</b>
            </td>
            <td width="1%" class="ToolBar" align="center"></td>
          </tr>
<%    
      
      var clase = "MessageTR1";
      
      creg = oRec.RecordCount;
      
      if (pagenumber * SHOW_CANT < creg){
        oRec.Move(pagenumber * SHOW_CANT); 
      }
      else{
        pagenumber = 0;
      }
      
      var i = 0;
      while ((oRec.EOF == false) && (i < SHOW_CANT)){ 
        i = i + 1;
        if (clase == "MessageTR1"){clase = "MessageTR";} else {clase = "MessageTR1";};
%>
          <tr> 
            <td width="1%" class="<%=clase%>">           
              <%
                if ((oRec.Fields.Item("Moderador").value == Session("userID")) || (Session("PermissionType") == ADMINISTRATOR)){
              %>
              <a href="javascript:abreVentana('S<%=oRec.Fields.Item("ID").value%>',640,480,'newChatURE.asp?salaid=<%=oRec.Fields.Item("ID").value%>','no','no')">
                <img title="Administrar" border=0 src="../images/<%=Session("skin")%>/Admin.gif">
              </a>
              <%}%>
            </td>
            <td width="1%" class="<%=clase%>">
              <%if (oRec.Fields.Item("publica").value){%>
                <img src="../images/<%=Session("skin")%>/icon_folder.gif" title="Sala publica">    
              <%}
                else{
              %>
                <img src="../images/<%=Session("skin")%>/icon_folder_locked.gif" title="Sala privada">    
              <%}%>  
            </td>
            <td width="1%" class="<%=clase%>"  align=center>
              <b><%=oRec.Fields.Item("Cantidad").value%></b>
            </td>
            <td width="24%" class="<%=clase%>">
              <a href="javascript:abreVentana('S<%=Session("userID")%>a<%=oRec.Fields.Item("ID").value%>',600,400,'newChat.asp?salaid=<%=oRec.Fields.Item("ID").value%>','no','no')">
                <b>[<%=oRec.Fields.Item("Nombre").value%>]</b>
              </a>  
            </td>
            <td width="77%" class="<%=clase%>">
              <%=oRec.Fields.Item("Objetivo").value%>
            </td>
            <td width="1%" class="<%=clase%>">           
              <%
                if ((oRec.Fields.Item("Moderador").value == Session("userID")) || (Session("PermissionType") == ADMINISTRATOR)){
              %>
              <a href="newChatRoomList.asp?pn=<%=pagenumber%>&op=1&sd=<%=oRec.Fields.Item("ID").value%>">
                <img title="Eliminar" border=0 src="../images/<%=Session("skin")%>/admintree/denied.gif">
              </a>
              <%}%>              
            </td>
          </tr> 
<%
        oRec.Move(1);
      }  
      oRec.Close();
%>      
        </table>  
        <table border="0" cellspacing="0" cellpadding="0" class="ToolBar" width="100%">
          <tr>
            <td>
              <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
                <tr> 
                  <td nowrap>
                    <a href="newChatAddRoom.asp" class="ToolLink">
                      &nbsp;Agregar nueva sala&nbsp;
                    </a>
                  </td>
                  <td>|</td>
                  <td>
                    <a href="newChatRoomList.asp?pn=<%=pagenumber + 1%>" class="ToolLink">
                      &nbsp;Próximas&nbsp;<%=SHOW_CANT%>&nbsp;
                    </a>
                  </td>
                </tr>
              </table>
              <hr>
              <table border="0" cellspacing="1" cellpadding="1" class="ToolBar" width="100%">
                <tr  valign=top>
                  <td colspan=5  align=center  valign=top>
                    <%
                      i = 0;
                      while (i * SHOW_CANT < creg){
                    %>    
                    <a class="numbers" href="newChatRoomList.asp?pn=<%=i%>">
                      <b><%=i + 1%>&nbsp;</b>
                    </a>  
                    <%
                        i = i + 1;
                      }
                    %>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
      </td>  
    </tr>  
  </table>  
</BODY>
</HTML>


















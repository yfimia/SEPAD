<%@ Language=JScript %>
<%
 Response.Expires = -1;
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
  
  oConn.Open(filePath);
  
  //Verifico si la sala es valida
  oRec.Open("SELECT * FROM Salas WHERE (ID = " + sala + ")", oConn, 3, 3);    
  salapublica = false;
  if (oRec.EOF == false){
    nameofsala = oRec.Fields.Item("Nombre").value;
    salamoderador = oRec.Fields.Item("Moderador").value;
    oRec.Close();
    
    //Verifico si el usuario actual puede administrar
    if ((Session("userID") == salamoderador) || (Session("PermissionType") == ADMINISTRATOR)){
%>
<HTML>
<HEAD>
  <title>Configuración de la sala >> <%=nameofsala%></title>
  <link rel="stylesheet" href="../css/<%=Session("skin")%>/Chat.css" type="text/css">
</HEAD>
<frameset rows="170,*" frameborder="YES" border="0" framespacing="0" cols="*"> 
  <frame name="topFrame" scrolling="NO" noresize src="newChatRoomPropertiesChange.asp?salaid=<%=sala%>" >
  <frameset cols="50%,*" frameborder="YES" border="0" framespacing="0" rows="*"> 
    <frame name="leftFrame" scrolling="YES" src="newChatAddURE.asp?salaid=<%=sala%>&pn=0">
    <frame name="rightFrame" scrolling="YES" noresize src="newChatDeleteURE.asp?salaid=<%=sala%>&pn=0" >
  </frameset>
</frameset>
</HTML>
<%
    }
    else{
      oConn.Close();
      oConn.Close();
      Response.Redirect("errorpage.asp?tipo=Error&short=Acceso denegado." + "&desc=El usuario actual no tiene suficientes privilegios para efectuar la operación");    
    }
  }
  else{
    oConn.Close();
    oConn.Close();
    Response.Redirect("errorpage.asp?tipo=Error&short=Error con la sala seleccionada." + "&desc=" + INVALID_ROOM);
  }
  oConn.Close();
%>

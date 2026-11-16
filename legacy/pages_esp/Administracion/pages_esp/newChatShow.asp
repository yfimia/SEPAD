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
%>
<HTML>
<HEAD>
  <title>Sala de charlas</title>
  <link rel="stylesheet" href="../css/<%=Session("skin")%>/Chat.css" type="text/css">
</HEAD>
<frameset rows="*,*" frameborder="YES" border="0" framespacing="0" cols="*"> 
  <frame ID="publicos" name="publicos" scrolling="YES" src="newChatPublicMsg.asp?salaid=<%=sala%>">
  <frame ID="privados" name="privados" scrolling="YES" src="newChatPrivateMsg.asp?salaid=<%=sala%>">
</frameset>
</HTML>

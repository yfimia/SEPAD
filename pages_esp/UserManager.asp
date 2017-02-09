<%@ Language=JScript %>
<%
  Response.Expires = 0;
%>

<!-- #include file='../js/user.inc' -->

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

%>

<%
  if (Session("PermissionType") == ADMINISTRATOR)
    {
%>
<HTML>
<HEAD>
<link REL="stylesheet" TYPE="text/css" HREF="../css/<%=Session("skin")%>/SepadCss1.css" />
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function button1_onclick() 
  {
    location.href ="DeleteUser.asp?uid=<%=uid%>";
  } 

function button2_onclick() 
  {
    location.href ="AgregaUser.asp?uid=<%=uid%>";
  } 

function button3_onclick() 
  {
    location.href ="Deletegroup.asp?uid=<%=uid%>";
  } 

function button4_onclick() 
  {
    location.href ="agregagroup.asp?uid=<%=uid%>";
  } 

function button5_onclick() 
  {
    location.href ="agregaUsergroup.asp?uid=<%=uid%>";
  } 

function button6_onclick() 
  {
    location.href ="DeleteUsergroup.asp?uid=<%=uid%>";
  } 
  
function button7_onclick() 
  {
    location.href ="SelectUserToModify.asp?uid=<%=uid%>";
  } 
//-->
</SCRIPT>
<link REL="stylesheet" TYPE="text/css" HREF="../css/<%=Session("skin")%>/SepadCss1.css" />
</HEAD>
<BODY>
  
  <table border=0 align=center width="100%" height="100%">
    <tr>
      <td align=center>
        <input type=button value="Agregar usuario" id=button2 name=button2 LANGUAGE=javascript onclick="return button2_onclick()">
      </td>
    </tr>  
    <tr>
      <td align=center>
        <input type=button value="Borrar usuario" id=button1 name=button1 LANGUAGE=javascript onclick="return button1_onclick()">
      </td>
    </tr>  
    <tr>
      <td align=center>
        <input type=button value="Editar usuario" id=button7 name=button7 LANGUAGE=javascript onclick="return button7_onclick()">
      </td>
    </tr>  

    <tr>
      <td align=center>
        <input type=button value="Borrar grupo" id=button3 name=button3 LANGUAGE=javascript onclick="return button3_onclick()">
      </td>
    </tr>  
    <tr>
      <td align=center>
        <input type=button value="Agregar grupo" id=button4 name=button4 LANGUAGE=javascript onclick="return button4_onclick()">
      </td>
    </tr>     
    <tr>
      <td align=center>
        <input type=button value="Agregar Usuarios a un grupo" id=button4 name=button4 LANGUAGE=javascript onclick="return button5_onclick()">
      </td>
    </tr>     
    <tr>
      <td align=center>
        <input type=button value="Eliminar Usuarios de un grupo" id=button4 name=button4 LANGUAGE=javascript onclick="return button6_onclick()">
      </td>
    </tr>     

  </table>
</BODY>
</HTML>
<%
    }
  else
    {
      Response.Redirect("errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);     
    }
%>

<%
  }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
  //Response.Write(Request.QueryString.Item("uid"));
%>
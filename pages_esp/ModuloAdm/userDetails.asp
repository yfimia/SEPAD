<%@ Language=JScript %>

<%
  Response.Expires = -1;
%>

<!-- #include file='../../js/user.inc' -->
<!-- #include file='inc/userDetails.inc' -->

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

 if ((Session("PermissionType") == ADMINISTRATOR) || (Session("userID") == Session("admcursocordinador"))  || (Session("userID") == Session("admcursoOwner")))
    {

    var where = " where (ID = '" +  Request.QueryString("ID") + "')";
       
     
    
%>

<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/Main.css" type="text/css">
<title><%= TXT_TITLE %></title>
</HEAD>
<body bgcolor="#FFFFFF" text="#000000">


<%    
      var  filePath = Application("dataPath");   
      var  oConn    = Server.CreateObject("ADODB.Connection");
      var  oRec     = Server.CreateObject("ADODB.Recordset");
      oConn.Open(filePath);
      oRec.Open("select * from Usuarios where (id = '" + Request.QueryString("id") + "')",oConn,3,3);    
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr width="100%"> 
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <td class=Toolbar width="100%" align=left><b><%= TEXT1 %></b></td>
    </table>
  </tr>
</table>  
<table width="100%" border="0" cellspacing="0" cellpadding="0">    
  <tr width="100%">
    <td width="40%" Class="MessageTR" align=right><%= TEXT2 %>:</td>
    <td width="60%" Class="MessageTR" align=center><%=oRec.Fields.Item("FullName").Value%></td>
  </tr>
  <tr width="100%">
    <td width="40%" Class="MessageTR1" align=right><%= TEXT3 %>:</td>
    <td width="60%" Class="MessageTR1" align=center><%=oRec.Fields.Item("Name").Value%></td>
  </tr>
  <tr width="100%">
    <td width="40%" Class="MessageTR" align=right><%= TEXT4 %>:</td>
    <td width="60%" Class="MessageTR" align=center><%=oRec.Fields.Item("Email").Value%></td>
  </tr>  
  <tr width="100%">
    <td width="40%" Class="MessageTR1" align=right><%= TEXT5 %>:</td>
    <td width="60%" Class="MessageTR1" align=center><%=oRec.Fields.Item("ci").Value%></td>
  </tr>  
  <tr width="100%">
    <td width="40%" Class="MessageTR" align=right><%= TEXT6 %>:</td>
    <td width="60%" Class="MessageTR" align=center><%=oRec.Fields.Item("fechaNac").Value%></td>
  </tr> 
  <tr width="100%">
    <td width="40%" Class="MessageTR1" align=right><%= TEXT7 %>:</td>
    <td width="60%" Class="MessageTR1" align=center>
      <%if (oRec.Fields.Item("sexo").Value == true) {Response.Write("Masculino")} else {Response.Write("Femenino")}%>
    </td>
  </tr>
  <tr width="100%">
    <td width="40%" Class="MessageTR" align=right><%= TEXT8 %>:</td>
    <td width="60%" Class="MessageTR" align=center><%=oRec.Fields.Item("direccion").Value%></td>
  </tr> 
</table>  
<table width="100%" border="0" cellspacing="0" cellpadding="0">      
  <tr>
    <td width="35%" Class="MessageTR1" align=right><%= TEXT9 %>:</td>
    <td width="15%" Class="MessageTR1" align=center><%=oRec.Fields.Item("cpostal").Value%></td>
    <td width="35%" Class="MessageTR1" align=right><%= TEXT10 %>:</td>
    <td width="15%" Class="MessageTR1" align=center><%=oRec.Fields.Item("telefono").Value%></td>
  </tr> 
</table>  
<table width="100%" border="0" cellspacing="0" cellpadding="0">      
  <tr width="100%">
    <td width="40%" Class="MessageTR" align=right><%= TEXT11 %>:</td>
    <td width="60%" Class="MessageTR" align=center><%=oRec.Fields.Item("Pais").Value%></td>
  </tr> 
  <tr width="100%">
    <td width="40%" Class="MessageTR1" align=right><%= TEXT12 %>:</td>
    <td width="60%" Class="MessageTR1" align=center><%=oRec.Fields.Item("Provincia").Value%></td>
  </tr> 
  <tr width="100%">
    <td width="40%" Class="MessageTR" align=right><%= TEXT13 %>:</td>
    <td width="60%" Class="MessageTR" align=center><%=oRec.Fields.Item("titulo").Value%></td>
  </tr> 
  <tr width="100%">
    <td width="40%" Class="MessageTR1" align=right><%= TEXT14 %>:</td>
    <td width="60%" Class="MessageTR1" align=center><%=oRec.Fields.Item("profesion").Value%></td>
  </tr> 
  <tr width="100%">
    <td width="40%" Class="MessageTR" align=right><%= TEXT15 %>:</td>
    <td width="60%" Class="MessageTR" align=center><%=oRec.Fields.Item("especialidad").Value%></td>
  </tr> 
  <tr width="100%">
    <td width="40%" Class="MessageTR1" align=right><%= TEXT16 %>:</td>
    <td width="60%" Class="MessageTR1" align=center><%=oRec.Fields.Item("annosgrad").Value%></td>
  </tr> 
  <tr width="100%">
    <td width="40%" Class="MessageTR" align=right><%= TEXT17 %>:</td>
    <td width="60%" Class="MessageTR" align=center><%=oRec.Fields.Item("institucion").Value%></td>
  </tr> 
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr width="100%"> 
    <td class=Toolbar width="100%" align=left><b><%= TEXT18 %></b></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr width="100%">
    <td width="40%" Class="MessageTR1" align=right><%= TEXT19 %>:</td>
    <td width="60%" Class="MessageTR1" align=center><%=oRec.Fields.Item("trabactual").Value%></td>
  </tr> 
  <tr width="100%">
    <td width="40%" Class="MessageTR" align=right><%= TEXT20 %>:</td>
    <td width="60%" Class="MessageTR" align=center><%=oRec.Fields.Item("cargoactual").Value%></td>
  </tr> 
  <tr width="100%">
    <td width="40%" Class="MessageTR1" align=right><%= TEXT21 %>:</td>
    <td width="60%" Class="MessageTR1" align=center><%=oRec.Fields.Item("otros").Value%></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr width="100%"> 
    <td class=Toolbar width="100%" align=left><b><%= TEXT22 %></b></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr width="100%">
    <td width="40%" Class="MessageTR" align=right><%= TEXT23 %>:</td>
    <td width="60%" Class="MessageTR" align=center><%=oRec.Fields.Item("catDoc").Value%></td>
  </tr> 
  <tr width="100%">
    <td width="40%" Class="MessageTR1" align=right><%= TEXT24 %>:</td>
    <td width="60%" Class="MessageTR1" align=center><%=oRec.Fields.Item("catInv").Value%></td>
  </tr> 
  <tr width="100%">
    <td width="40%" Class="MessageTR" align=right><%= TEXT25 %>:</td>
    <td width="60%" Class="MessageTR" align=center><%=oRec.Fields.Item("catCient").Value%></td>
  </tr>
</table>
<br>
<!--table border="0" cellspacing="0" cellpadding="0" class="ToolBar" width="100%">
  <tr>
    <td>
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td><a href="javascript:history.back()" class="ToolLink">&nbsp;Regresar&nbsp;</a></td>
        </tr>
      </table>
    </td>
  </tr>
</table-->


<%      
      oRec.Close();

      oConn.Close();
%>

</BODY>
</HTML>

<%
    }
  else  
    {
     Response.Redirect("../errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);         
    } 

  }
 else
   Response.Redirect("../errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
  //Response.Write(Request.QueryString.Item("uid"));
%>

<%@ Language=JScript %>

<%
  Response.Expires = -1;
%>

<!-- #include file='../../js/user.inc' -->

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

  if ((Session("PermissionType") == ADMINISTRATOR) || (Session("userID") == Session("admcordinador")))
    {

    var where = " where (ID = '" +  Request.QueryString("ID") + "')";
       
     
    
%>

<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/Main.css" type="text/css">
<title>Datos Generales</title>
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
      <td class=Toolbar width="100%" align=left><b>Datos Generales</b></td>
    </table>
  </tr>
</table>  
<table width="100%" border="0" cellspacing="0" cellpadding="0">    
  <tr width="100%">
    <td width="40%" Class="MessageTR" align=right>Nombre&nbsp;y&nbsp;Apellidos:</td>
    <td width="60%" Class="MessageTR" align=center><%=oRec.Fields.Item("FullName").Value%></td>
  </tr>
  <tr width="100%">
    <td width="40%" Class="MessageTR1" align=right>Identificador:</td>
    <td width="60%" Class="MessageTR1" align=center><%=oRec.Fields.Item("Name").Value%></td>
  </tr>
  <tr width="100%">
    <td width="40%" Class="MessageTR" align=right>Email:</td>
    <td width="60%" Class="MessageTR" align=center><%=oRec.Fields.Item("Email").Value%></td>
  </tr>  
  <tr width="100%">
    <td width="40%" Class="MessageTR1" align=right>N&uacute;mero&nbsp;de&nbsp;Identidad:</td>
    <td width="60%" Class="MessageTR1" align=center><%=oRec.Fields.Item("ci").Value%></td>
  </tr>  
  <tr width="100%">
    <td width="40%" Class="MessageTR" align=right>Fecha&nbsp;de&nbsp;Nacimiento:</td>
    <td width="60%" Class="MessageTR" align=center><%=oRec.Fields.Item("fechaNac").Value%></td>
  </tr> 
  <tr width="100%">
    <td width="40%" Class="MessageTR1" align=right>Sexo:</td>
    <td width="60%" Class="MessageTR1" align=center>
      <%if (oRec.Fields.Item("sexo").Value == true) {Response.Write("Masculino")} else {Response.Write("Femenino")}%>
    </td>
  </tr>
  <tr width="100%">
    <td width="40%" Class="MessageTR" align=right>Direcci&oacute;n:</td>
    <td width="60%" Class="MessageTR" align=center><%=oRec.Fields.Item("direccion").Value%></td>
  </tr> 
</table>  
<table width="100%" border="0" cellspacing="0" cellpadding="0">      
  <tr width="100%">
    <td width="35%" Class="MessageTR1" align=right>C&oacute;digo&nbsp;Postal:</td>
    <td width="15%" Class="MessageTR1" align=center><%=oRec.Fields.Item("cpostal").Value%></td>
    <td width="35%" Class="MessageTR1" align=right>Tel&eacute;fono:</td>
    <td width="15%" Class="MessageTR1" align=center><%=oRec.Fields.Item("telefono").Value%></td>
  </tr> 
</table>  
<table width="100%" border="0" cellspacing="0" cellpadding="0">      
  <tr width="100%">
    <td width="40%" Class="MessageTR" align=right>Pa&iacute;s:</td>
    <td width="60%" Class="MessageTR" align=center><%=oRec.Fields.Item("Pais").Value%></td>
  </tr> 
  <tr width="100%">
    <td width="40%" Class="MessageTR1" align=right>Provincia:</td>
    <td width="60%" Class="MessageTR1" align=center><%=oRec.Fields.Item("Provincia").Value%></td>
  </tr> 
  <tr width="100%">
    <td width="40%" Class="MessageTR" align=right>T&iacute;tulo:</td>
    <td width="60%" Class="MessageTR" align=center><%=oRec.Fields.Item("titulo").Value%></td>
  </tr> 
  <tr width="100%">
    <td width="40%" Class="MessageTR1" align=right>Profesi&oacute;n:</td>
    <td width="60%" Class="MessageTR1" align=center><%=oRec.Fields.Item("profesion").Value%></td>
  </tr> 
  <tr width="100%">
    <td width="40%" Class="MessageTR" align=right>Especialidad:</td>
    <td width="60%" Class="MessageTR" align=center><%=oRec.Fields.Item("especialidad").Value%></td>
  </tr> 
  <tr width="100%">
    <td width="40%" Class="MessageTR1" align=right>Años de graduado:</td>
    <td width="60%" Class="MessageTR1" align=center><%=oRec.Fields.Item("annosgrad").Value%></td>
  </tr> 
  <tr width="100%">
    <td width="40%" Class="MessageTR" align=right>Instituci&oacute;n:</td>
    <td width="60%" Class="MessageTR" align=center><%=oRec.Fields.Item("institucion").Value%></td>
  </tr> 
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr width="100%"> 
    <td class=Toolbar width="100%" align=left><b>Perfiles&nbsp;Laborales</b></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr width="100%">
    <td width="40%" Class="MessageTR1" align=right>Trabajo&nbsp;Actual:</td>
    <td width="60%" Class="MessageTR1" align=center><%=oRec.Fields.Item("trabactual").Value%></td>
  </tr> 
  <tr width="100%">
    <td width="40%" Class="MessageTR" align=right>Cargo&nbsp;Actual:</td>
    <td width="60%" Class="MessageTR" align=center><%=oRec.Fields.Item("cargoactual").Value%></td>
  </tr> 
  <tr width="100%">
    <td width="40%" Class="MessageTR1" align=right>Otros:</td>
    <td width="60%" Class="MessageTR1" align=center><%=oRec.Fields.Item("otros").Value%></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr width="100%"> 
    <td class=Toolbar width="100%" align=left><b>Perfiles&nbsp;Investigativos</b></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr width="100%">
    <td width="40%" Class="MessageTR" align=right>Categor&iacute;a&nbsp;Docente:</td>
    <td width="60%" Class="MessageTR" align=center><%=oRec.Fields.Item("catDoc").Value%></td>
  </tr> 
  <tr width="100%">
    <td width="40%" Class="MessageTR1" align=right>Categor&iacute;a&nbsp;Investigativa:</td>
    <td width="60%" Class="MessageTR1" align=center><%=oRec.Fields.Item("catInv").Value%></td>
  </tr> 
  <tr width="100%">
    <td width="40%" Class="MessageTR" align=right>Categor&iacute;a&nbsp;Cient&iacute;fica:</td>
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

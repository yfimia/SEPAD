<%@ Language=jScript %>

<%Response.Expires = -1%>
<!-- #include file='../js/adolibrary.inc' -->
<!-- #include file='../js/user.inc' -->

<%
  var uid = Request.QueryString.Item("uid");
%>

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/EvalResult.css" type="text/css">
<script language=jscript src ="../js/library.js"></script>
<script language=jscript src ="../js/windows.js"></script>


</head>

<body bgcolor="#FFFFFF" text="#000000">
      
<%
  var filePath = Application("dataPath");
  var oCon = Server.CreateObject("ADODB.Connection");
  var oRec = Server.CreateObject("ADODB.Recordset");
  
  oCon.Open(filePath);
  Sql = "SELECT * FROM Enlaces";
  oRec.Open(Sql, oCon,3,3);
%>

<form name="EnlacesDisp" id="EnlacesDisp" ACTION="processOptions.asp?uid=<%=uid%>" METHOD=POST>
<table cellpadding=0 cellspacing=0 border=0 width="100%">
        <tr class="Toolbar"> 
          <td><b>Manipulando opciones</b></td>
        </tr>
   <tr>
     <td>
       <table border=0 cellpadding=0 cellspacing=1 width="100%">
        <tr>
         <td class="ToolBar" align="center"><b>Texto del enlace</b></td>
         <td class="ToolBar" align="center"><b>URL del enlace</b></td>
         <td class="ToolBar" align="center"><b>Modificar (M)</b></td>
         <td class="ToolBar" align="center"><b>Eliminar (E)</b></td>
       </tr>
     </td>   
  <%
    var cont = 0;
    while (!oRec.EOF) {
  %>
  <%
   if (cont == 0) {
      cont = 1;
  %> 
     <tr class="MessageTR">
  <%
    } else
    if (cont == 1) {
     cont = 0;
  %>
     <tr class="MessageTR1">
  <%
    }
  %> 
     <td>
       <%=oRec.Fields.Item("caption").Value%>
     </td>
     <td>
       <%=oRec.Fields.Item("url").Value%>
     </td>
     <td align="center">
      <a href="ModifyOption.asp?oid=<%=oRec.Fields.Item('ID').Value%>&capt=<%=oRec.Fields.Item("caption").Value%>&url=<%=oRec.Fields.Item("url").Value%>">M</a>
     </td>
     <td align="center">
      <a href="DeleteOption.asp?oid=<%=oRec.Fields.Item('ID').Value%>">E</a>
     </td>
   </tr>
  <%
    oRec.Move(1);
    }
  %> 
</table>

</form>

<%
  oRec.Close();
  oCon.Close();
%>
<hr>
<form name="AddLink" id="AddLink" ACTION="AddLink.asp?uid=<%uid%>" METHOD=POST>
<table cellpadding=0 cellspacing=1 border=0 width="100%">
   <tr> 
     <td width=100% class="Toolbar"><b>Agregar Opci&oacuten</b></td>
   </tr>
   <tr>     
     <td>
      <table cellpadding=0 cellspacing=1 border=0 width="100%">
        <tr class="MessageTR">
         <td>Texto del enlace</td>
         <td><input type=text name="LinkText" id="LinkText"></td>
        </tr>  
        <tr class="MessageTR1">
         <td>URL del enlace</td>
         <td><input type=text name="LinkUrl" id="LinkUrl"></td>
        </tr>  
       </table> 
     </td>
   </tr> 
</table>

   <table cellpadding=0 cellspacing=0 border=0 width="100%">
     <tr>
       <td align="center">
        <input type=submit value="Adicionar" id=submit1 name=submit1>
       </td>
     </tr>
   </table>
   
</form>

</body>
</html>

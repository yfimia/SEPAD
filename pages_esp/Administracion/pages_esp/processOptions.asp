<%@ Language=jScript %>

<%Response.Expires = -1%>
<!-- #include file='../js/adolibrary.inc' -->
<!-- #include file='../js/user.inc' -->

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/EvalResult.css" type="text/css">
<script language=jscript src ="../js/library.js"></script>


</head>

<body bgcolor="#FFFFFF" text="#000000">
      
<table width="100%" border="0" cellspacing="0" cellpadding="2">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar" width="100%">
        <tr> 
          <td width=100% class="HeaderTable"><b>Opciones disponibles</b></td>
        </tr>
      </table>
    </td>
  </tr>
</table>

<%
  var filePath = Application("dataPath");
  var oCon = Server.CreateObject("ADODB.Connection");
  var oRec = Server.CreateObject("ADODB.Recordset");
  
  oCon.Open(filePath);
  Sql = "SELECT * FROM Enlaces";
  oRec.Open(Sql, oCon,3,3);
%>

<form name="EnlacesDisp" id="EnlacesDisp" ACTION="" METHOD=POST>

<table cellpadding=0 cellspacing=0 border=1 width="100%">
   <tr>
     <td class="ToolBar">
       &nbsp
     </td>
     <td class="ToolBar">
       Texto del enlace
     </td>
     <td class="ToolBar">
       URL del enlace
     </td>
     <td class="ToolBar">
       Modificar S/N
     </td>
   </tr>
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
       <input type=checkbox name="OpActiva" id="OpActiva">
     </td>
     <td>
       <%=oRec.Fields.Item("Caption").Value%>
     </td>
     <td>
       <%=oRec.Fields.Item("url").Value%>
     </td>
     <td>
      <a href="">Modificar</a>
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

</body>
</html>

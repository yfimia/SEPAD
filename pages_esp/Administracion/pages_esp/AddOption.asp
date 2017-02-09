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
          <td width=100% class="HeaderTable"><b>Agregar Opci&oacuten</b></td>
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

<form name="AddLink" id="AddLink" ACTION="AddLink.asp" METHOD=POST>

<table cellpadding=0 cellspacing=0 border=1 width="100%">
   <tr class="MessageTR">
     <td class="ToolBar">
       Texto del enlace
     </td>
     <td>
       <input type=text name="LinkText" id="LinkText" size=36>
     </td>
   </tr>  
   <tr class="MessageTR1">
     <td class="ToolBar">
       URL del enlace
     </td>
     <td>
       <input type=text name="LinkUrl" id="LinkUrl" size=36>
     </td>
   </tr>  
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

<%
  oRec.Close();
  oCon.Close();
%>

</body>
</html>

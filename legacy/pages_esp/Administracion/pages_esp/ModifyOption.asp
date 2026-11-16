<%@ Language=jScript %>

<%
  Response.Expires = -1
  
  var uid = Request.QueryString.Item("uid");
  
%>
<!-- #include file='../js/adolibrary.inc' -->
<!-- #include file='../js/user.inc' -->

<html>
<head>
<title>Modificar Enlace</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/EvalResult.css" type="text/css">
<script language=jscript src ="../js/library.js"></script>


</head>

<body bgcolor="#FFFFFF" text="#000000">
<%
  var Caption = Request.QueryString.Item("capt");
  var Url = Request.QueryString.Item("url");
  var oid = Request.QueryString.Item("oid");
  
  var filePath = Application("dataPath");
  var oCon = Server.CreateObject("ADODB.Connection");
  var oRec = Server.CreateObject("ADODB.Recordset");
  
  oCon.Open(filePath);
  Sql = "SELECT * FROM Enlaces";
  oRec.Open(Sql, oCon,3,3);
%>

<form name="ModLink" id="ModLink" ACTION="ModLink.asp?oid=<%=oid%>&uid=<%=uid%>" METHOD=POST>

<table cellpadding=0 cellspacing=1 border=0 width="100%">
   <tr> 
     <td width=100% class="ToolBar"><b>Modificar Opci&oacuten</b></td>
   </tr>
   <tr>
    <td>
     <table cellpadding=0 cellspacing=1 border=0 width="100%">
       <tr class="MessageTR">
         <td>Texto del enlace</td>
         <td><input type=text name="LinkText" id="LinkText" value="<%=Caption%>"></td>
       </tr>  
       <tr class="MessageTR1">
         <td>URL del enlace</td>
         <td><input type=text name="LinkUrl" id="LinkUrl" value="<%=Url%>"></td>
       </tr>  
      </td> 
     </table>  
   </tr>
</table>

   <table cellpadding=0 cellspacing=0 border=0 width="100%">
     <tr>
       <td align="center">
        <input type=submit value="Actualizar">
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

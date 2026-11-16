<%@ Language=JavaScript %>
<%
 Response.Expires = -1;
%>

<!-- #include file="../js/course.inc" -->
<!-- #include file='../js/adolibrary.inc' -->
<!-- #include file='../js/user.inc' --> 


<%
   if (Request.QueryString.Item("uid") + "" == Session("uid"))
   {    
     var uid = Request.QueryString.Item("uid") + "";
     var material = Request.QueryString.Item("material");
     var lesson = Request.QueryString.Item("lesson");
     var userID = Session("userID");
     
     var  filePath = Application("dataPath");   
     var  oConn    = Server.CreateObject("ADODB.Connection");
     var  oComm    = Server.CreateObject("ADODB.Command");
     var  oRec     = Server.CreateObject("ADODB.Recordset");
     oConn.Open( filePath );
     oRec.Open("select * from Lecciones where (ID = " + Request.QueryString.Item("lesson") + ")",oConn,3,3);
     var leccion = oRec.Fields("Name").Value;
     oRec.Close;
     oRec.Open("select * from Ficheros where (ID = " + Request.QueryString.Item("material") + ")",oConn,3,3);
     var mater = oRec.Fields("title").Value;
     oRec.Close;
     oRec.Open("select * from Anotaciones where ((alumn = " + userID + ") and (Material = " + material + "))",oConn,3,3);
%>     

<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/EvalResult.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<SCRIPT LANGUAGE=javascript src="../js/md5.js"></script>
<SCRIPT LANGUAGE=javascript src="../js/user.js"></script>
<script language=javascript>

function elimina(leccion, mater, uid)
{
  window.navigate("docs.asp?uid=" + uid + "&lesson=" + leccion + "&material=" + mater + "&eliminar=SI");
}

</script>

<form name=anotac id=anotac action="docs.asp?uid=<%=uid%>&lesson=<%=lesson%>&material=<%=material%>" method=post  LANGUAGE=javascript >
 
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td class="ToolBar">
        <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
          <tr> 
            <td class="HeaderTable" align=center><b>Anotaciones</b></td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
  
  <table border="0" cellspacing="1" cellpadding="1" align="center" width="100%">
    <tr width="100%">
      <td align="right" cellspacing="0" cellpadding="0" width="50%" class="ToolBar">Lecci&oacuten:</td>
      <td cellspacing="0" cellpadding="0" width="50%" class="ToolBar"><%=leccion%></td>
    </tr>
    <tr width="100%">
      <td align="right" cellspacing="0" cellpadding="0" width="50%" class="ToolBar">Material:</td>
      <td cellspacing="0" cellpadding="0" width="50%" class="ToolBar"><%=mater%></td>
    </tr>
    <tr width="100%">
      <td colspan="2" class=MessageTR width="100%">
        <textarea class="Textarea" align="center" id="texto" name="texto" rows="14"><%if (!oRec.EOF) {%><%=oRec.Fields("Texto").Value%><%;}%></textarea>
      </td>
    </tr>
    <tr width="100%">
      <table border="0" cellspacing="1" cellpadding="1" align="center" width="100%">
        <tr width="100%">
          <td  Class="Toolbar" align="right" width="100%">
            <input type=submit value="Salvar" id=submit1 name=submit1>
          </td>
          <td  Class="Toolbar" align="center" width="100%">
            <input type=button value="Eliminar" id=submit2 name=submit2 onclick="elimina(<%=lesson%>,<%=material%>,<%=uid%>)">
          </td>          
        </tr>        
      </table>
    </tr>
     
  </table>

</form>

</body>
</html>            


<%
     oRec.Close;
   }
   else
   {  
     Response.Redirect("errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);         
   }

%>
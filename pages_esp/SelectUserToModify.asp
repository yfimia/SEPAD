<%@ Language=JScript %>
<%Response.Expires = -1;%>
<!-- #include file='../js/user.inc' -->

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

 if (Session("PermissionType") == ADMINISTRATOR)
    {
    first = "-1";
    where = "";
    if (Request.QueryString.Count >= 2) {
      first = Request.QueryString.Item("first") + "";
      
      if (first != "-1") {
        where = " and (Name > '" +  first + "')";
      } 
    } 
    
%>


<HTML>
 <HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
 </HEAD>
<body bgcolor="#FFFFFF" text="#000000">
<%    
    
      var  filePath = Application("dataPath");   
      var  oConn    = Server.CreateObject("ADODB.Connection");
      var  oComm    = Server.CreateObject("ADODB.Command");
      var  oRec     = Server.CreateObject("ADODB.Recordset");
      oRec.CursorLocation = 3;
      oRec.CursorType     = 3;
      oConn.Open(filePath);
      oComm.ActiveConnection = oConn;
      oComm.CommandText = "SELECT Top " + SHOW_CANT + " * FROM usuarios where (ID <> " + GUEST_USER + ")" +  where + " Order By Name";
      oRec = oComm.Execute();    
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar" align="center"><b>Listado de usuarios</b></td>
  </tr>
  <tr> 
    <td> 
      <table width="100%" border="0" cellspacing="1" cellpadding="0">
        <tr> 
          <td width="20%" class="ToolBar" align="center"><b>Identificador</b></td>
          <td width="40%" class="ToolBar" align="center"><b>Nombre y apellidos</b></td>
          <td width="40%" class="ToolBar" align="center"><b>Correo electr�nico</b></td>
        </tr>
<%    
      var clase = "MessageTR1";
      
      var i = 0;
      var last = "-1"
      while (oRec.EOF == false)
        { 
          i = i + 1;
          if (clase == "MessageTR1") clase = "MessageTR"; else clase = "MessageTR1";
          last = oRec.Fields.Item("Name").value;          
%>
        <tr> 
          <td width="20%" class="<%=clase%>"><a href="ModifyUser.asp?uid=<%=uid%>&user=<%=oRec.Fields.Item("ID").value%>"><%=oRec.Fields.Item("name").value%></a></td>
          <td width="40%" class="<%=clase%>"><%=oRec.Fields.Item("FullName").value%></td>
          <td width="40%" class="<%=clase%>"><a href="mailto:<%=oRec.Fields.Item("email").value%>" ><%=oRec.Fields.Item("email").value%></a></td>
        </tr> 
<%          
          oRec.Move(1);
        }  
      if (i > 0) 
        {
          oRec.MoveFirst();
//          Response.Write('<tr><td><INPUT id=Buton name=Buton type=Submit value=Borrar></td><td><INPUT id=Buton name=Buton type=Submit value=Regresar onclick="history.back(-1)"></td></tr>');
        }
      else
        {
%>        
          <tr><td align="center" width="100%" colspan="4" class="MessageTR">No quedan m�s elementos.</td></tr>        
<%          
        }  
%>        
<table border="0" cellspacing="0" cellpadding="0" class="ToolBar" width="100%">
  <tr>
    <td>
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td><a href="SelectUserToModify.asp?uid=<%=uid%>&first=<%=last%>" class="ToolLink">&nbsp;Pr�ximos&nbsp;<%=SHOW_CANT%>&nbsp;</a></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<%        
    }
  else  
    {
    Response.Redirect("errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);         
    }
%>

</BODY>
</HTML>
<%
  }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
  //Response.Write(Request.QueryString.Item("uid"));
%>
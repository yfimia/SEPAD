<%@ Language=JScript %>
<%
  Response.Expires = -1;
  
%>
<!-- #include file='../../js/user.inc' -->
<!-- #include file='inc/verEliminarSubGrupo.inc' -->
<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       
    first = "-1";
    where = "";
    if (Request.QueryString.Item("first").Count > 0) {
      first = Request.QueryString.Item("first") + "";
      
      if (first != "-1") {
        where = " and (Name > '" +  first + "')";
      } 
    } 
    

%>

<%  


  if ((Session("PermissionType") == ADMINISTRATOR) || (Session("userID") == Session("admcursocordinador"))  || (Session("userID") == Session("admcursoOwner")))
    {
%>

<html>
<head>
<title><%= TXT_TITLE %></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/EvalResult.css" type="text/css">
<script src="../../js/CheckBoxes.js" language="JavaScript"></script>
<script language="JavaScript">
<!--
function close_window() {
    window.close();
}
//-->
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form name=Delnews action="ConfirmDeleteSubGrupo.asp?uid=<%=uid%>" method="post">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td class="HeaderTable"  align=center><b><%= TEXT1 %></b</td>
        </tr>
      </table>
    </td>
  </tr>
</table>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  
<%    
      var  filePath = Application("dataPath");
      var  oConn    = Server.CreateObject("ADODB.Connection");
      var  oComm    = Server.CreateObject("ADODB.Command");
      var  oRec     = Server.CreateObject("ADODB.Recordset");
      oRec.CursorLocation = 3;
      oRec.CursorType     = 3;
      oConn.Open(filePath);
      oComm.ActiveConnection = oConn;

      
%>
  <tr> 
    <td> 
      <table width="100%" border="0" cellspacing="1" cellpadding="0">
        <tr> 
          <td width="1%" class="ToolBar" align="center">&nbsp; </td>
          <td width="20%" class="ToolBar" align="center"><b><%= TEXT2 %></b></td>
          <td width="44%" class="ToolBar" align="center"><b><%= TEXT3 %></td>
          <td width="30%" class="ToolBar" align="center"><b><%= TEXT4 %></b></td>
          <td width="5%" class="ToolBar" align="center"></td>
        </tr>
        <tr> 
<%    
      var clase = "MessageTR1";
      

      oComm.CommandText = "SELECT Usuarios.FullName, Usuarios.Email, SubGrupo.*   "  +
                          "FROM Usuarios INNER JOIN SubGrupo ON Usuarios.ID = SubGrupo.Tutor " +
                          "WHERE (((SubGrupo.Curso)=" + Session("admcurso") + ")) Order By SubGrupo.Name";

      oRec = oComm.Execute();
      var i = 0;
      last = "-1";          
      while (oRec.EOF == false)
        {
          last = oRec.Fields.Item("Name").value;          
          i = i + 1;
          if (clase == "MessageTR1") clase = "MessageTR"; else clase = "MessageTR1";
          
          
%>            
          <td width="1%" class="<%=clase%>"> 
            <input type="checkbox" id=Check<%=i%> name=check<%=i%> >
          </td>
          <td width="20%" class="<%=clase%>"><%=oRec.Fields.Item("Name").value%></td>
          <td width="44%" class="<%=clase%>"><%=oRec.Fields.Item("Desc").value%></td>
          <td width="30%" class="<%=clase%>"><a href="mailto:<%=oRec.Fields.Item("email").value%>" ><%=oRec.Fields.Item("fullName").value%></a></td>
          <td width="5" class="<%=clase%>"><a href="modifySubGrupo.asp?uid=<%=uid%>&sgrupo=<%=oRec.Fields.Item("ID").value%>" >Modificar</a></td>
          
<%              
          oRec.Move(1);
%>
        </tr>
<%
          
        }  
      if (i > 0) 
        {
          oRec.MoveFirst();
//          Response.Write('<tr><td><INPUT id=Buton name=Buton type=Submit value=Borrar></td><td><INPUT id=Buton name=Buton type=Submit value=Regresar onclick="history.back(-1)"></td></tr>');
        }
      else
        {
%>        
          <tr><td align="center" width="100%" colspan="4" class="MessageTR"><%= TEXT5 %>.</td></tr>        
<%          
        }  
      
      Session("UserList") = oRec;
      Session("Conection") = oConn;
      Session("Command") = oComm;
%>
      </table>
    </td>
  </tr>
  <tr>
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td nowrap><a href="javascript:Delnews.submit()" class="ToolLink">&nbsp;<%= TEXT6 %>&nbsp;</a></td>
          <td>|</td>
          <td><a href="#" class="ToolLink" onClick="CheckAll(GetObjectCollection( document, 'FORM'))">&nbsp;<%= TEXT7 %>&nbsp;</a></td>
        </tr>
      </table>
    </td>
  </tr>
<%
    }
  else  
    {
     Response.Redirect("../errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);         
    } 
%>
  
</table>
</form>      

</body>
</html>
<%
  }
 else
   Response.Redirect("../errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>

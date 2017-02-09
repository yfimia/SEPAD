<%@ Language=JScript %>
<%
  Response.Expires = -1;
%>
<!-- #include file='../js/user.inc' -->

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       


if (Session("PermissionType") == ADMINISTRATOR)
    {
    first = "-1";
    
    where = "";
    if (Request.QueryString.Count >= 4) {
      first = Request.QueryString.Item("first") + "";
      
      if ((first != "-1")) {
        where = " and ((Usuarios.Name > '" +  first + "'))";
      } 
    } 
    
   group = Request.QueryString.Item("group") + "";
   groupname = Request.QueryString.Item("groupname") + "";
   
%>

<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<script src="../js/CheckBoxes.js" language="JavaScript"></script>
</HEAD>
<body bgcolor="#FFFFFF" text="#000000">
<form name=Delnews action="ConfirmDeleteUserOneGroup.asp?uid=<%=uid%>&group=<%=group%>&groupname=<%=groupname%>&first=<%=last%>" method="post">
<%    
      var  filePath = Application("dataPath");   
      var  oConn    = Server.CreateObject("ADODB.Connection");
      var  oComm    = Server.CreateObject("ADODB.Command");
      var  oRec     = Server.CreateObject("ADODB.Recordset");
      oRec.CursorLocation = 3;
      oRec.CursorType     = 3;
      oConn.Open(filePath);
      oComm.ActiveConnection = oConn;
      oComm.CommandText = "SELECT Usuarios.Name, Usuarios.ID as UID, Usuarios.FullName, Usuarios.Email, Grupos_de_Usuarios.[User] AS UserID, Grupos_de_Usuarios.[Group] AS GroupID " +
                          "FROM Usuarios INNER JOIN Grupos_de_Usuarios ON Usuarios.ID = Grupos_de_Usuarios.[User] " + 
                          "WHERE (((Grupos_de_Usuarios.[Group])=" + group + ")) " +
                          where + " Order By Usuarios.Name";
      oRec = oComm.Execute();
%>
<table border="0" cellspacing="0" cellpadding="0" class="ToolBar" width="100%">
  <tr>
    <td>
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td nowrap><a href="AgregaUserGroup.asp?uid=<%=uid%>&first=a&group=<%=group%>" class="ToolLink">&nbsp;Adicionar&nbsp;Usuarios&nbsp;</a></td>
          <td>|</td>
          <td nowrap><a href="javascript:Delnews.submit()" class="ToolLink">&nbsp;Eliminar&nbsp;Usuarios&nbsp;</a></td>
          <td>|</td>
          <td><a href="#" class="ToolLink" onClick="CheckAll(GetObjectCollection( document, 'FORM'))">&nbsp;Seleccionar&nbsp;Todo&nbsp;</a></td>
          <td>|</td>
          <td><a href="Deleteuseronegroup.asp?uid=<%=uid%>&group=<%=group%>&groupname=<%=groupname%>&first=<%=last%>" class="ToolLink">&nbsp;Próximos&nbsp;<%=SHOW_CANT%>&nbsp;</a></td>
        </tr>
      </table>
    </td>
  </tr>
</table>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar" align="center"><b>Listado de usuarios del grupo: <%= groupname%></b></td>
  </tr>
  <tr> 
    <td> 
      <table width="100%" border="0" cellspacing="1" cellpadding="0">
        <tr> 
          <td width="5%" class="ToolBar" align="center"></td>
          <td width="20%" class="ToolBar" align="center"><b>Identificador</b></td>
          <td width="40%" class="ToolBar" align="center"><b>Nombre y apellidos</b></td>
          <td width="35%" class="ToolBar" align="center"><b>Correo eléctronico</b></td>
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
          <td width="5%" class="<%=clase%>"><input type="checkbox" id=Check<%=i%> name=check<%=i%> ></td>
          <td width="20%" class="<%=clase%>"><a href="ModifyUser.asp?uid=<%=uid%>&user=<%=oRec.Fields.Item("UID").value%>"><%=oRec.Fields.Item("name").value%></a></td>
          <td width="40%" class="<%=clase%>"><%=oRec.Fields.Item("FullName").value%></td>
          <td width="35%" class="<%=clase%>"><a href="mailto:<%=oRec.Fields.Item("email").value%>" ><%=oRec.Fields.Item("email").value%></a></td>
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
          <tr><td align="center" width="100%" colspan="4" class="MessageTR">No quedan más elementos.</td></tr>        
<%          
        }  
%>        
<table border="0" cellspacing="0" cellpadding="0" class="ToolBar" width="100%">
  <tr>
    <td>
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td nowrap><a href="AgregaUserGroup.asp?uid=<%=uid%>&first=a&group=<%=group%>" class="ToolLink">&nbsp;Adicionar&nbsp;Usuarios&nbsp;</a></td>
          <td>|</td>
          <td nowrap><a href="javascript:Delnews.submit()" class="ToolLink">&nbsp;Eliminar&nbsp;Usuarios&nbsp;</a></td>
          <td>|</td>
          <td><a href="#" class="ToolLink" onClick="CheckAll(GetObjectCollection( document, 'FORM'))">&nbsp;Seleccionar&nbsp;Todo&nbsp;</a></td>
          <td>|</td>
          <td><a href="Deleteuseronegroup.asp?uid=<%=uid%>&group=<%=group%>&groupname=<%=groupname%>&first=<%=last%>" class="ToolLink">&nbsp;Próximos&nbsp;<%=SHOW_CANT%>&nbsp;</a></td>
        </tr>
      </table>
    </td>
  </tr>
</table>

<%
      Session("CoursesList") = oRec;
      Session("Conection") = oConn;
      Session("Command") = oComm;
%>
 </form>
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

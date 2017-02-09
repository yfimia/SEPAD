<%@ Language=JScript %>
<%
  Response.Expires = -1;
  
%>
<!-- #include file='../js/user.inc' -->
<!-- #include file='inc/verEliminaralumnado.inc' -->
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
    coursename = Request.QueryString.Item('coursename') + '';
  course     = Request.QueryString.Item('course') + '';

  Session('course')  = course;
  Session('courseName') = coursename;


  if ((Session("PermissionType") == ADMINISTRATOR) || (Session("PermissionType") == PUBLICATOR))
    {
%>

<html>
<head>
<title><%= TITLE1 %></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<script src="../js/CheckBoxes.js" language="JavaScript"></script>
<script language="JavaScript">
<!--
function close_window() {
    window.close();
}
//-->
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form name=Delnews action="ConfirmDeleteAlumn.asp?uid=<%=uid%>" method="post">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar" align="center"><b><%= TITLE2 %> 
      <%=coursename%></b></td>
  </tr>
<%    
      var  filePath = Application("dataPath");
      var  oConn    = Server.CreateObject("ADODB.Connection");
      var  oComm    = Server.CreateObject("ADODB.Command");
      var  oRec     = Server.CreateObject("ADODB.Recordset");
      oRec.CursorLocation = 3;
      oRec.CursorType     = 3;
      oConn.Open(filePath);
      oComm.ActiveConnection = oConn;

      oComm.CommandText = "SELECT * FROM grupos where (Name = 'Curso_" + coursename + "')";
      oRec = oComm.Execute();
      var groupID = oRec.Fields.Item('ID').value;
      oRec.Close();
      
%>
  <tr> 
    <td> 
      <table width="100%" border="0" cellspacing="1" cellpadding="0">
        <tr> 
          <td width="1%" class="ToolBar" align="center">&nbsp; </td>
          <td width="20%" class="ToolBar" align="center"><b><%= TEXT1 %></b></td>
          <td width="39%" class="ToolBar" align="center"><b><%= TEXT2 %></b></td>
          <td width="40%" class="ToolBar" align="center"><b><%= TEXT3 %></b></td>
        </tr>
        <tr> 
          <input  name="groupID" type="hidden" value=<%=groupID%> >        
<%    
      var clase = "MessageTR1";
      

      oComm.CommandText = "SELECT Top " + SHOW_CANT + "   Usuarios.ID, Usuarios.email,Usuarios.Name, Usuarios.FullName FROM Usuarios INNER JOIN Grupos_de_Usuarios ON Usuarios.ID = Grupos_de_Usuarios.[User] WHERE (((Grupos_de_Usuarios.[Group])=" + groupID + ")) " +  where + " Order By Name";
      oRec = oComm.Execute();
      var i = 0;
      last = "-1";          
      while (oRec.EOF == false)
        {
          last = oRec.Fields.Item("Name").value;          
          i = i + 1;
          if (clase == "MessageTR1") clase = "MessageTR"; else clase = "MessageTR1";
          
          
          if (oRec.Fields.Item("FullName").value == null)
            {
%>            
          <td width="1%" class="<%=clase%>"> 
            <input type="checkbox" id=Check<%=i%> name=check<%=i%> >
          </td>
          <td width="20%" class="<%=clase%>"><%=oRec.Fields.Item("name").value%></td>
          <td width="39%" class="<%=clase%>"><%=oRec.Fields.Item("name").value%></td>
          <td width="40%" class="<%=clase%>"><a href="mailto:<%=oRec.Fields.Item("email").value%>" ><%=oRec.Fields.Item("email").value%></a></td>
              
<%              
            
            }
          else
            {
%>            
          <td width="1%" class="<%=clase%>"> 
            <input type="checkbox" id=Check<%=i%> name=check<%=i%> >
          </td>
          <td width="20%" class="<%=clase%>"><%=oRec.Fields.Item("name").value%></td>
          <td width="39%" class="<%=clase%>"><%=oRec.Fields.Item("FullName").value%></td>
          <td width="40%" class="<%=clase%>"><a href="mailto:<%=oRec.Fields.Item("email").value%>" ><%=oRec.Fields.Item("email").value%></a></td>
              
<%              
            }
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
          <tr><td align="center" width="100%" colspan="4" class="MessageTR"><%= TEXT4 %>.</td></tr>        
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
          <td nowrap><a href="javascript:Delnews.submit()" class="ToolLink">&nbsp;<%= TEXT5 %>&nbsp;</a></td>
          <td>|</td>
          <td><a href="#" class="ToolLink" onClick="CheckAll(GetObjectCollection( document, 'FORM'))">&nbsp;<%= TEXT6 %>&nbsp;</a></td>
          <td>|</td>
          <td><a href="#" class="ToolLink" onClick="close_window()">&nbsp;<%= TEXT7 %>&nbsp;</a></td>
          <td>|</td>
          <td><a href="verEliminaralumnado.asp?uid=<%=uid%>&course=<%=course%>&coursename=<%=coursename%>&first=<%=last%>" class="ToolLink">&nbsp;<%= TEXT8 %>&nbsp;<%=SHOW_CANT%>&nbsp;</a></td>
        </tr>
      </table>
    </td>
  </tr>
<%
    }
  else  
    {
     Response.Redirect("errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);         
    } 
%>
  
</table>
</form>      

</body>
</html>
<%
  }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>

<%@ Language=JScript %>
<%
  Response.Expires = -1;
%>


<!-- #include file='../../js/user.inc' -->
<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

%>

<%  
  
  if ((Session("PermissionType") == ADMINISTRATOR) || (Session("userID") == Session("admcordinador")))
    {
    first = "-1";
    where = "";
    if (Request.QueryString("first").Count > 0) {
      first = Request.QueryString.Item("first") + "";
      
      if (first != "-1") {
        where = " and (Usuarios.name > '" +  first + "')";
      } 
    } 
    
%>




<html>
<head>
<title>Administraci&oacute;n de matrículas</title>
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
<form name=Delnews action="AceptaMtrCur.asp?uid=<%=uid%>" method="post">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td class="HeaderTable"  align=center><b>Solicitudes de matrículas</b></td>
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
      oComm.CommandText = "SELECT  Top " + SHOW_CANT + "  TMP_Matriculas_Cursos.ID, TMP_Matriculas_Cursos.[User], Usuarios.FullName,  Usuarios.name, Usuarios.email FROM Usuarios INNER JOIN TMP_Matriculas_Cursos ON Usuarios.ID = TMP_Matriculas_Cursos.[User] WHERE ((TMP_Matriculas_Cursos.Modulo)=" + Session("admmodulo") + ") " +  where + " Order By name";
      oRec = oComm.Execute();    
%>
  <tr> 
    <td> 
      <table width="100%" border="0" cellspacing="1" cellpadding="0">
        <tr> 
          <td width="1%" class="ToolBar" align="center">&nbsp; </td>
          <td width="20%" class="ToolBar" align="center"><b>Identificador</b></td>
          <td width="39%" class="ToolBar" align="center"><b>Nombre y apellidos</b></td>
          <td width="40%" class="ToolBar" align="center"><b>Correo Electr&oacute;nico</b></td>
        </tr>
        <tr> 
<%    
      var clase = "MessageTR1";
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
          <td width="1%" class="MessageTR"> 
            <input type="checkbox" id=Check<%=i%> name=check<%=i%> >
          </td>
          <td width="20%" class="<%=clase%>"><a href="userDetails.asp?uid=<%=uid%>&id=<%=oRec.Fields.Item("User").value%>"><%=oRec.Fields.Item("name").value%></a></td>
          <td width="39%" class="<%=clase%>"><%=oRec.Fields.Item("FullName").value%></td>
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
        }
      else
        {
%>        
          <tr><td align="center" width="100%" colspan="4" class="MessageTR">No hay solicitudes de matrículas.</td></tr>
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
          <td nowrap><a href="javascript:Delnews.submit()" class="ToolLink">&nbsp;Matricular&nbsp;</a></td>
          <td>|</td>
          <td nowrap><a href="javascript:Delnews.action = 'ConfirmDeleteMtrCur.asp?uid=<%=uid%>';Delnews.submit()" class="ToolLink">&nbsp;Denegar&nbsp;</a></td>
          <td>|</td>
          
          <td><a href="#" class="ToolLink" onClick="CheckAll(GetObjectCollection( document, 'FORM'))">&nbsp;Seleccionar&nbsp;Todo&nbsp;</a></td>
          <td>|</td>
          <td><a href="mtrcurmanager.asp?uid=<%=uid%>&first=<%=last%>" class="ToolLink">&nbsp;Próximos&nbsp;<%=SHOW_CANT%>&nbsp;</a></td>
          
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


<%@ Language=JScript %>

<%
  Response.Expires = -1;
%>
<!-- #include file='../js/user.inc' -->

 <%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       
 
  %>

<html>
<head>
<title> SubGrupo : <%=Session("sgName")%> </title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/EvalResult.css" type="text/css">

<!--
function close_window() {
    window.close();
}
//-->
</script>
</head>



<body>
  
<%    
      var  filePath = Application("dataPath");
      var  oConn    = Server.CreateObject("ADODB.Connection");
      //var  oComm    = Server.CreateObject("ADODB.Command");
      var  oRec     = Server.CreateObject("ADODB.Recordset");
      oConn.Open(filePath);
      //sql = "SELECT * FROM grupos where (ID = " + Session("admgrupo") + ")";
      
      //      var groupID = oRec.Fields.Item('ID').value;
     
    
      
      sql = "SELECT  Usuarios.ID, Usuarios.email, Usuarios.Name, Usuarios.FullName FROM Usuarios INNER JOIN UsuariosSubGrupo ON Usuarios.ID = UsuariosSubGrupo.[Usuario] WHERE (((UsuariosSubGrupo.[SubGrupo])=" + Session("SgId") + "))";
      oRec.Open(sql,oConn,3,3);
      %>
        
           
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td class="HeaderTable"  align=center><b>Alumnado</b></td>
        </tr>
        <tr> 
          <td><h5 style="color: #ffffff">SubGrupo:&nbsp;<%=Session("sgName")%></h5></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
          
        <table width="100%" border=0 cellpadding=0 cellspacing=1>  
          <tr> 
              
              <td width="20%" class="ToolBar" align="center"><b>Identificador</b></td>
              <td width="39%" class="ToolBar" align="center"><b>Nombre y apellidos</b></td>
              <td width="40%" class="ToolBar" align="center"><b>Correo Electr&oacute;nico</b></td>
          </tr>     
      <%         
      while (oRec.EOF == false) {
        
        %>
          <tr> 
             <td width="20%" class="MessageTR"><a href="userDetails.asp?uid=<%=uid%>&id=<%=oRec.Fields.Item("id")%>"><%=oRec.Fields.Item("Name")%></a></td>
             <td width="39%" class="MessageTR"><%=oRec.Fields.Item("FullName")%></td>
             <td width="40%" class="MessageTR"><a href="mailto:<%=oRec.Fields.Item("email")%>"><%=oRec.Fields.Item("email")%></a></td>
              

        </tr> 
        <%  
             oRec.Move(1);
            }
%>
    </table>
<%
  }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
  //Response.Write(Request.QueryString.Item("uid"));
%>
</body>
</html>


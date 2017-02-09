<%@ Language=JScript %>
<%
  Response.Expires = -1;
  
%>
<!-- #include file='../../js/user.inc' -->
<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       
    first = "-1";
    where = "";
    if (Request.QueryString.Item("first").Count > 0) {
      first = Request.QueryString.Item("first") + "";
      
      if (first != "-1") {
        where = " and (Cursos.Name > '" +  first + "')";
      } 
    } 
    

%>

<%  
 /*   coursename = Request.QueryString.Item('coursename') + '';
  course     = Request.QueryString.Item('course') + '';

  Session('course')  = course;
  Session('courseName') = coursename;
*/

  if ((Session("PermissionType") == ADMINISTRATOR) || (Session("userID") == Session("admcordinador")))
    {
%>

<html>
<head>
<title>Ver/Eliminar alumnado</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/EvalResult.css" type="text/css">
<script src="../../js/CheckBoxes.js" language="JavaScript"></script>
<script src="../../js/windows.js" language="JavaScript"></script>

<script language="JavaScript">
<!--
function close_window() {
    window.close();
}
//-->
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000">

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td class="HeaderTable"  align=center><b>Módulos</b</td>
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
          <td width="5%" class="ToolBar" align="center"><b>Código</b></td>
          <td width="20%" class="ToolBar" align="center"><b>Nombre</b></td>
          <td width="10%" class="ToolBar" align="center"><b>Estado</b></td>
          <td width="30%" class="ToolBar" align="center"><b>Profesor principal</b></td>
          <td width="30%" class="ToolBar" align="center"><b>Correo Electr&oacute;nico</b></td>
          <td width="4" class="ToolBar" align="center"></td>
        </tr>
        <tr> 
<%    
      var clase = "MessageTR1";
      

      oComm.CommandText = "SELECT Top " + SHOW_CANT + "  Cursos.prologue, Cursos.ID, Cursos.state, Cursos.owner, Cursos.Name, Usuarios.FullName, Usuarios.Email, Usuarios.ID AS UID FROM Usuarios INNER JOIN Cursos ON Usuarios.ID = Cursos.Owner WHERE (Cursos.modulo = " + Session("admmodulo") + ")"  +  where + " Order By Cursos.Name";
      oRec = oComm.Execute();
      var i = 0;
      last = "-1";          
      while (oRec.EOF == false)
        {
          last = oRec.Fields.Item("Name").value;          
          i = i + 1;
          if (clase == "MessageTR1") clase = "MessageTR"; else clase = "MessageTR1";
          
          
%>            
          <td width="1%" class="<%=clase%>"><a href="ConfirmDeleteCurso.asp?uid=<%=uid%>&course=<%=oRec.Fields.Item("id").value%>&prologue=<%=oRec.Fields.Item("prologue").value%>" title="Eliminar módulo"><img  width="20" height="20" border="0"  src="../../images/<%=Session("skin")%>/admintree/denied.gif"></a>  </td>
          <td width="5%" class="<%=clase%>"><%=oRec.Fields.Item("id").value%></td>
          <td width="20%" class="<%=clase%>"><%=oRec.Fields.Item("name").value%></td>
          <td width="10%" nowrap class="<%=clase%>">
<%
          
          switch (oRec.Fields.Item("state").value) {
           case MOD_ACA_NOVISIBLE: {%>No visible<% break;}
           case MOD_ACA_DISABLE  : {%>Deshabilitado<% break;}
           case MOD_ACA_INCOURSE : {%>En curso<% break;}
           case MOD_ACA_FINISHED : {%>Finalizado<% break;}
           default : {%>Deshabilitado<% break;}
          }       
              
          
%>        </td>
          <td width="30%" class="<%=clase%>"><a href="userDetails.asp?uid=<%=uid%>&id=<%=oRec.Fields.Item("owner").value%>"><%=oRec.Fields.Item("FullName").value%></a></td>
          <td width="30%" class="<%=clase%>"><a href="mailto:<%=oRec.Fields.Item("email").value%>" ><%=oRec.Fields.Item("email").value%></a></td>
          <td width="4%" class="<%=clase%>">
            <a  href="javascript:abreVentana('CourseAdm',700,400,'../moduloAdm/main.asp?uid=<%=uid%>&course=<%=oRec.Fields.Item("ID").value%>&courseName=<%=oRec.Fields.Item("Name").value%>&courseOwner=<%=oRec.Fields.Item("UID").value%>&modulo=<%=Session("admmodulo")%>&moduloName=<%=Session("admmoduloName")%>&cordinador=<%=Session("cordinador")%>&state=<%=Session("admstate")%>&grupo=<%=Session("admgrupo")%>&claustro=<%=Session("admclaustro")%>', 'yes', 'yes')" ><img title="Administrar" src="../../images/<%=Session("skin")%>/admin.gif"  width="20" height="20" border="0" ></a>
          </td>
              
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
          <tr><td align="center" width="100%" colspan="7" class="MessageTR">No hay solicitudes de matrículas.</td></tr>        
<%          
        }  
      
      oRec.Close();
      oConn.Close();
%>
      </table>
    </td>
  </tr>
  <tr>
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td><a href="verEliminarCurso.asp?uid=<%=uid%>&first=<%=last%>" class="ToolLink">&nbsp;Próximos&nbsp;<%=SHOW_CANT%>&nbsp;</a></td>
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


</body>
</html>
<%
  }
 else
   Response.Redirect("../errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>


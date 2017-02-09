<%@ Language=jScript %>
<%  
 Response.Expires = -1;
%>
<!-- #include file="../../js/Adolibrary.inc" --> 
<!-- #include file='../../js/user.inc' -->
<%
  if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

 if ((Session("isTutor") > -1) || (Session("PermissionType") == ADMINISTRATOR) || (Session("userID") == Session("tutcursocordinador"))  || (Session("userID") == Request.QueryString.Item("courseOwner")))
    {

    first = "-1";
    where = "";
    if (Request.QueryString.Item("first").Count > 0) {
      first = Request.QueryString.Item("first") + "";
      
      if (first != "-1") {
        where = " and (uploadDate < " + Application("dchar") +  first  + Application("dchar") + ")";
      } 
    } 
      		
var userID = Session("userID") + "";
var courseID = Session("tutcurso") + "";


%>
<html>
<head>
<title>Manipular trabajos recibidos</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/EvalResult.css" type="text/css">
<script type="text/javascript" language="javascript" src="../../js/Windows.js"></script>
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
<form name=Delnews action="ConfirmDeleteTrab.asp?uid=<%=uid%>&courseOwner=<%=Request.QueryString.Item("courseOwner")%>" method="post">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
     <td class="ToolBar" align="center"><b>Trabajos recibidos</b></td>  </tr>
  <tr> 
    <td> 
      <table width="100%" cellpadding="0" cellspacing="1" border="0">
        <tr> 
          <td width="1%" class="ToolBar" align="center"></td>
          <td width="20%" class="ToolBar" align="center"><b>Archivo</b></td>
          <td width="7%" class="ToolBar" align="center"><b>Tama&ntilde;o</b></td>
          <td width="42%" class="ToolBar" align="center"><b>Descripci&oacute;n</b></td>
          <td width="20%" class="ToolBar" align="center"><b>Usuario</b></td>
          <td width="10%" class="ToolBar" align="center"><b>Fecha</b></td>
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

      if (Session("isTutor") > -1) {
        var SQL        = "SELECT  Top " + SHOW_CANT + " Usuarios.ID as UID, Seminarios.ID, Seminarios.fileTitle, Seminarios.fileSize, Seminarios.fileDescription, Seminarios.fileURL, Seminarios.uploadDate, Usuarios.FullName, UsuariosSubGrupo.Subgrupo " +
                         "FROM (Usuarios INNER JOIN Seminarios ON Usuarios.ID = Seminarios.userID) INNER JOIN UsuariosSubGrupo ON Usuarios.ID = UsuariosSubGrupo.Usuario " +
                         "WHERE (Seminarios.courseID=" + courseID + ") and (UsuariosSubGrupo.Subgrupo=" + Session("isTutor") + ")  " +  where + " "  +
                         "ORDER BY Seminarios.uploadDate DESC";
     }                        
     else {
       var SQL         = "SELECT  Top " + SHOW_CANT + "  Usuarios.ID as UID, Seminarios.ID, Seminarios.fileTitle, Seminarios.fileSize, Seminarios.fileDescription, Seminarios.fileURL, Seminarios.uploadDate, Usuarios.FullName " + 
		         "FROM Usuarios INNER JOIN Seminarios ON Usuarios.ID = Seminarios.userID " + 
		         "WHERE (Seminarios.courseID=" + courseID + ") " +  where + " Order By uploadDate Desc";
     }
      
      

      oComm.CommandText = SQL;
      oRec = oComm.Execute();
      
var trclass = 'MessageTR';

size = 0;
Count = 0;
last = "-1";          
while (oRec.EoF == false) {     
      	if (trclass == 'MessageTR') {
     		trclass = 'MessageTR1'; 
     	} else {
		   trclass = 'MessageTR';
	      }	
    	Count++;       
    	last = oRec.Fields.Item("uploadDate").value;          
    	size += oRec.Fields.Item("fileSize");
	
%>
        <tr class="<%=trclass %>"  valign="middle"> 
          <td width="1%"  align="center"><input type="checkbox" id=Check<%=Count%> name=check<%=Count%> ></td>
          <td width="20%" align="left"><a target=_blank href="../../Courses/Course<%=courseID %>/Seminarios/<%=oRec.Fields.Item("fileURL").value %>"><%=oRec.Fields.Item("fileTitle") %></a></td>
          <td width="7%" align="right"><%=oRec.Fields.Item("fileSize") %>&nbsp;kb</td>
	      <td width="42%" align="left">&nbsp;&nbsp;<%=oRec.Fields.Item("fileDescription") %>&nbsp;</td>
          <td width="20%" align=center><a href="userDetails.asp?uid=<%=uid%>&courseOwner=<%=Request.QueryString.Item("courseOwner")%>&id=<%=oRec.Fields.Item("UID").value%>"><%=oRec.Fields.Item("FullName") %></a></td>
          <td width="10%" align="center"><%=oRec.Fields.Item("uploadDate") %></td>
        </tr>
<%
    	oRec.Move(1); 
}  
%>
        <tr class="<%=trclass %>"  valign="middle"> 
          <td  align="right" colspan=3 ><b> Tamaño usado: </b></td>
          <td  align="center" colspan=3 ><%=size%>&nbsp;kb </td>
        </tr>
        <tr class="<%=trclass %>"  valign="middle"> 
          <td  align="right" colspan=3><b> Tamaño máximo permitido:&nbsp;</b></td><td  align="center" colspan=3 ><%= Session("tutcursomaxmem") %> kb</td>
        </tr>
<%
     if (Count > 0) 
        {
          oRec.MoveFirst();
        }
      else
        {
%>        
          <tr><td align="center" width="100%" colspan="6" class="MessageTR">No hay trabajos enviados.</td></tr>        
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
          <td nowrap><a href="javascript:Delnews.submit()" class="ToolLink">&nbsp;Eliminar&nbsp;</a></td>
          <td>|</td>
          <td><a href="#" class="ToolLink" onClick="CheckAll(GetObjectCollection( document, 'FORM'))">&nbsp;Seleccionar&nbsp;Todo&nbsp;</a></td>
          <td>|</td>
          <td><a href="adminuploadedfiles.asp?uid=<%=uid%>&courseOwner=<%=Request.QueryString.Item("courseOwner")%>&first=<%=last%>" class="ToolLink">&nbsp;Próximos&nbsp;<%=SHOW_CANT%>&nbsp;</a></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</form>
</body>
</html>

<%
   }
  else
    Response.Redirect("../errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);     
 }
 else
   Response.Redirect("../errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>

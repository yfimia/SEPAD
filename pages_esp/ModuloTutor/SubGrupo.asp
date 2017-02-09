<%@ Language=JScript %>
<%Response.Expires = -1;%>
<!-- #include file='../../js/adoLibrary.inc' -->
<!-- #include file='../../js/user.inc' -->
<!-- #include file='../../js/change.inc' -->

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

%>

<%
 

 if ((Session("isTutor") > -1) || (Session("PermissionType") == ADMINISTRATOR) || (Session("userID") == Session("tutcursocordinador"))  || (Session("userID") == Request.QueryString.Item("courseOwner") + ""))
    {
        
        
%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/EvalResult.css" type="text/css">

<script src="../../js/windows.js" language="JavaScript"></script>

</head>
<body bgcolor="#FFFFFF" text="#000000">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td class="HeaderTable"  align=center><b>Alumnos</b></td>
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
<%      
  if (Session("isTutor") > -1) {      
%>  
        <tr> 
          <td width="1%" class="ToolBar" align="center">&nbsp; </td>
          <td width="20%" class="ToolBar" align="center"><b>Identificador</b></td>
          <td width="44%" class="ToolBar" align="center"><b>Nombre</td>
          <td width="30%" class="ToolBar" align="center"><b>Email</b></td>
          <td width="5%" class="ToolBar" align="center">&nbsp;</td>
        </tr>
        <tr> 
        
<%
 }
 else {
%>
        <tr> 
          <td width="1%" class="ToolBar" align="center">&nbsp; </td>
          <td width="15%" class="ToolBar" align="center"><b>Subgrupo</b></td>
          <td width="20%" class="ToolBar" align="center"><b>Identificador</b></td>
          <td width="24%" class="ToolBar" align="center"><b>Nombre</td>
          <td width="25%" class="ToolBar" align="center"><b>Email</b></td>
          <td width="5%" class="ToolBar" align="center">&nbsp;</td>
        </tr>
        <tr> 

<%
 }
%>        
<%    
      var clase = "MessageTR1";
      
  if (Session("isTutor") > -1) 
      oComm.CommandText = "SELECT Usuarios.* " + 
                          "FROM (Usuarios INNER JOIN UsuariosSubGrupo ON Usuarios.ID = UsuariosSubGrupo.Usuario) INNER JOIN SubGrupo ON SubGRupo.ID = UsuariosSubGrupo.Subgrupo " + 
                          "WHERE (((UsuariosSubGrupo.Subgrupo)=" + Session("isTutor") + ") and (Subgrupo.Curso = " + Session("tutcurso")  +"))";
  else
      oComm.CommandText = "SELECT Usuarios.*, SubGrupo.Name as SName " + 
                          "FROM (Usuarios INNER JOIN UsuariosSubGrupo ON Usuarios.ID = UsuariosSubGrupo.Usuario) INNER JOIN SubGrupo ON SubGRupo.ID = UsuariosSubGrupo.Subgrupo " + 
                          "WHERE SubGrupo.Curso = " + Session("tutcurso") ;
                          
                      
      oRec = oComm.Execute();
      var i = 0;
      last = "-1";          
      while (oRec.EOF == false)
        {
          last = oRec.Fields.Item("Name").value;          
          i = i + 1;
          if (clase == "MessageTR1") clase = "MessageTR"; else clase = "MessageTR1";
          
  if (Session("isTutor") > -1) {
%>            
          <td width="1%" class="<%=clase%>">&nbsp;</td>
          <td width="20%" class="<%=clase%>"><a href="userDetails.asp?uid=<%=uid%>&courseOwner=<%=Request.QueryString.Item("courseOwner")%>&id=<%=oRec.Fields.Item("ID").value%>"><%=oRec.Fields.Item("Name").value%></a></td>
          <td width="44%" class="<%=clase%>"><%=oRec.Fields.Item("fullName").value%></td>
          <td width="30%" class="<%=clase%>"><a href="mailto:<%=oRec.Fields.Item("email").value%>" ><%=oRec.Fields.Item("email").value%></a></td>
          <td width="5%" class="<%=clase%>">
            <a href="verCalifOfic.asp?uid=<%=uid%>&curso=<%=Session("tutcurso")%>&alumn=<%=oRec.Fields.Item("ID").value%>">
              <img title="Calificaciones oficiales" src="../../images/<%=Session("skin")%>/CertOfic.gif"  width="20" height="20" border="0">
            </a>
          </td>
<%              
  }
  else {
%>
          <td width="1%" class="<%=clase%>">&nbsp;</td>
          <td width="15%" class="<%=clase%>"><%=oRec.Fields.Item("SName").value%></td>
          <td width="20%" class="<%=clase%>"><a href="userDetails.asp?uid=<%=uid%>&courseOwner=<%=Request.QueryString.Item("courseOwner")%>&id=<%=oRec.Fields.Item("ID").value%>"><%=oRec.Fields.Item("Name").value%></a></td>
          <td width="24%" class="<%=clase%>"><%=oRec.Fields.Item("fullName").value%></td>
          <td width="25%" class="<%=clase%>"><a href="mailto:<%=oRec.Fields.Item("email").value%>" ><%=oRec.Fields.Item("email").value%></a></td>
          <td width="5%" class="<%=clase%>">
            <a href="verCalifOfic.asp?uid=<%=uid%>&courseOwner=<%=Request.QueryString.Item("courseOwner")%>&curso=<%=Session("tutcurso")%>&alumn=<%=oRec.Fields.Item("ID").value%>&prof=<%=Session("userID")%>">
              <img title="Calificaciones oficiales" src="../../images/<%=Session("skin")%>/CertOfic.gif"  width="20" height="20" border="0">
            </a>
          </td>
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
          <tr><td align="center" width="100%" colspan="6" class="MessageTR">No hay alumnos asignados.</td></tr>        
<%          
        }  
%>
      </table>
    </td>
  </tr>
</table>


<%
    }    
  else
    {  
     Response.Redirect("../errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);                   
    }
%>

</BODY>
</HTML>
<%
  }
 else
   Response.Redirect("../errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
   
  //Response.Write(Request.QueryString.Item("uid"));
%>



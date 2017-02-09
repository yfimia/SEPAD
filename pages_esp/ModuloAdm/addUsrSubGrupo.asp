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
  coursename = Request.QueryString.Item('coursename') + '';
  course     = Request.QueryString.Item('course') + '';
    Session('course')  = course;
  Session('courseName') = coursename;


 if ((Session("PermissionType") == ADMINISTRATOR) || (Session("userID") == Session("admcursocordinador"))  || (Session("userID") == Session("admcursoOwner")))
    {
    first = "-1";
    where = "";
    if (Request.QueryString.Item("first").Count > 0) {
      first = Request.QueryString.Item("first") + "";
      
      if (first != "-1") {
        where = " and (Name > '" +  first + "')";
      } 
    } 
    
%>
<HEAD>
<title>Matricular usuarios</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/EvalResult.css" type="text/css">
<script src="../../js/CheckBoxes.js" language="JavaScript"></script>
</HEAD>
<body bgcolor="#FFFFFF" text="#000000">
<form name=Delnews action="ConfirmAgregaUserSubGrupo.asp?uid=<%=uid%>&first=<%=last%>" method="post">

<%    
      var  filePath = Application("dataPath");   
      var  oConn    = Server.CreateObject("ADODB.Connection");
      var  oComm    = Server.CreateObject("ADODB.Command");
      var  oRec     = Server.CreateObject("ADODB.Recordset");
      var  oRec1     = Server.CreateObject("ADODB.Recordset");
      oRec.CursorLocation = 3;
      oRec.CursorType     = 3;
      oRec1.CursorLocation = 3;
      oRec1.CursorType     = 3;
      oConn.Open(filePath);
      oComm.ActiveConnection = oConn;
      oComm.CommandText = "Select Top " + SHOW_CANT + "  Usuarios.* " + 
                          "FROM Usuarios INNER JOIN Grupos_de_Usuarios ON Usuarios.ID = Grupos_de_Usuarios.[User] " +
                          "WHERE (Grupos_de_Usuarios.[Group] = " + Session("admcursogrupo") +  ") and " + 
                          "(Usuarios.ID NOT IN (SELECT UsuariosSubGrupo.Usuario FROM SubGrupo INNER JOIN UsuariosSubGrupo ON SubGrupo.ID = UsuariosSubGrupo.Subgrupo WHERE (((SubGrupo.Curso)=" + Session("admcurso") + "))))"  +  where + " Order By Usuarios.Name";
      oRec = oComm.Execute();    
%>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td width=20% class="HeaderTable"  align=center><b>Subgrupo:</b></td>
          <td width=80% class="HeaderTable"  align=center>
      <select name=GrupoADD ID=GrupoADD class="TemasAsociados" >
<%
      oComm.CommandText = "SELECT * FROM SubGrupo where (SubGrupo.Curso = "  + Session("admcurso") + ") ";
      oRec1 = oComm.Execute();    
      while (oRec1.EOF == false) 
        {
%>        
          <option selected value="<%=oRec1.Fields.Item("ID").Value%>"><%=oRec1.Fields.Item("Name").Value%></option>
<%
          oRec1.Move(1);
        }
%>        
       </select>
          
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>


<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td class="HeaderTable"  align=center><b>Listado de usuarios</b></td>
        </tr>
        <tr> 
          <td   align=right>Nota: Los alumnos que no se se listan aquí ya tienen un tutor asignado.</td>
        </tr>
        
      </table>
    </td>
  </tr>
</table>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td> 
      <table width="100%" border="0" cellspacing="1" cellpadding="0">
  
        <tr> 
          <td width="5%" class="ToolBar" align="center"></td>
          <td width="20%" class="ToolBar" align="center"><b>Identificador</b></td>
          <td width="75%" class="ToolBar" align="center"><b>Nombre y apellidos</b></td>
        </tr>
      
<%    
      var clase = "MessageTR1";
      
      var i = 0;
      var last = "-1";
      while (oRec.EOF == false)
        { 
          i = i + 1;
          if (clase == "MessageTR1") clase = "MessageTR"; else clase = "MessageTR1";
          last = oRec.Fields.Item("Name").value;          

%>
        <tr> 
          <td width="5%" class="<%=clase%>"><input type="checkbox" id=Check<%=i%> name=check<%=i%> ></td>
          <td width="20%" class="<%=clase%>"><a href="userDetails.asp?uid=<%=uid%>&id=<%=oRec.Fields.Item("ID").value%>"><%=oRec.Fields.Item("name").value%></a></td>
          <td width="75%" class="<%=clase%>"><%=oRec.Fields.Item("FullName").value%></td>
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
          <td nowrap><a href="javascript:Delnews.submit()" class="ToolLink">&nbsp;Adicionar&nbsp;</a></td>
          <td>|</td>
          <td><a href="#" class="ToolLink" onClick="CheckAll(GetObjectCollection( document, 'FORM'))">&nbsp;Seleccionar&nbsp;Todo&nbsp;</a></td>
          <td>|</td>
          <td><a href="addUsrSubGrupo.asp?uid=<%=uid%>&first=<%=last%>" class="ToolLink">&nbsp;Próximos&nbsp;<%=SHOW_CANT%>&nbsp;</a></td>
        </tr>
      </table>
    </td>
  </tr>
<%   

      Session("UserList") = oRec;
      Session("Conection") = oConn;
      Session("Command") = oComm;
%>
</table>
 </form>
<%
    }
  else  
    {
      Response.Redirect("../../errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);                       
    }
%>

</BODY>
</HTML>
<%
  }
 else
   Response.Redirect("../../errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
  //Response.Write(Request.QueryString.Item("uid"));
%>

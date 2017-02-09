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
 

 if ((Session("PermissionType") == ADMINISTRATOR) || (Session("userID") == Session("admcursocordinador"))  || (Session("userID") == Session("admcursoOwner")))
    {
      if ((Request.Form("Mname").Count == 0) || (Request.Form("Mname") == ""))
        {

           var filePath = Application('dataPath');
           var oConn;
           var oRec;
 
           oConn = MakeConnection( oConn, filePath );
           Sql = "select SubGrupo.* FROM SubGrupo WHERE (SubGrupo.Curso = "  + Session("admcurso") + ") and (ID = " + Request.QueryString.Item("sgrupo") + ")";

           oRec = Query( Sql, oRec, oConn  );	
          

          if (oRec.EOF == true)
            { 
              oConn.Close();   
              
              Response.Redirect("../errorpage.asp?tipo=Error&short=No existe el subgrupo&desc=No existe el subgrupo que usted intenta modificar."); 
            }  
        
        
%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/EvalResult.css" type="text/css">
<script src="../../js/CheckBoxes.js" language="JavaScript"></script>
</head>
<script language=jscript>

function unchange(ss) {
    	var cnt = 0;
    	s = new String(ss);
    	var cl = "<br/>";
    	    brtext = "";
    	while (s.search(cl) != -1) {
         	s = s.replace(cl,brtext);
        } 
    	var cl = "<BR/>";
    	    brtext = "";
    	while (s.search(cl) != -1) {
         	s = s.replace(cl,brtext);
        } 
        
 return(s);        
}


 function change(ss) {
    	var cnt = 0;
    	s = new String(ss);
    	var cl = "\n";
    	    brtext = "<BR/>";
    	while (s.search(cl) != -1) {
         	s = s.replace(cl,brtext); 
        }
 return(s);        
}
  function User(id, name, fullName, email) {
    this.id = id;
    this.name = name;
    this.fullName = fullName;
    this.email = email;
  }
  
  
  function findUser() {
    var user = new Array();
    user = showModalDialog("findTutor.asp?uid=<%=uid%>", "", "");
    if (user != null) {
      newgroup.Mcord.value = user[0].fullName + ' (' + user[0].name + ')';
      newgroup.Mtutorid.value = user[0].id;
    }  
  }
  
</script>

<body bgcolor="#FFFFFF" text="#000000">
  <form name=newgroup id=newgroup action="modifySubGrupo.asp?uid=<%=uid%>&sgrupo=<%=Request.QueryString.Item("sgrupo")%>" method=post>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td class="HeaderTable"  align=center><b>Modificar Subgrupo Tutoreado</b></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
    <table width="100%" border="0" cellspacing="1" cellpadding="1" align="center">
      <tr>
        <td  align=right Class="MessageTR">
          Nombre:
        </td>
        <td  Class="MessageTR" align=left >
          <input name=Mname Id=Mname type=text maxlength=250 size=60 value="<%=oRec.Fields.Item("Name").Value%>">
        </td>
      </tr>
      <tr>
      <tr>
        <td align=right   Class="MessageTR1">
          Tutor:
        </td>
        <td align=left   Class="MessageTR1">
          <input disabled name=Mcord Id=Mcord type=text maxlength=250 size=35 value="<% var user = GetUserData(oRec.Fields.Item("Tutor").Value); Response.Write(user.fullName); %>"><input name=Mname Id=Mname type=button size=20 value="Seleccionar..." onclick="findUser()">
          <input name=Mtutorid Id=MtutorId type=hidden maxlength=250 size=25  value="<%=oRec.Fields.Item("Tutor").Value%>">
        </td>
      </tr>

    </table>

    <table width="100%" border="0" cellspacing="1" cellpadding="1" align="center">
      <tr>
        <td colspan=2 align=center   Class="MessageTR">
          Descripción
        </td>
      </tr>
      <tr>
        <td colspan=2 align=left   Class="MessageTR">
          <textarea name=Mdesc  rows="8" cols="78" maxlength=250><%=oRec.Fields.Item("Desc").Value%></textarea>
        </td>
      </tr>
      
      <tr>
        <td widht="100%" colspan=2 align=center Class="Toolbar">
          <center><input type=submit value="Actualizar" ></center>
        </td>      
      </tr>
    </table>
  </form>
<%

        }
      else
        {
           var filePath = Application('dataPath');
           var oConn;
           var oRec;
 
           oConn = MakeConnection( oConn, filePath );
           Sql = "select SubGrupo.* FROM SubGrupo WHERE (SubGrupo.Curso = "  + Session("admcurso") + ") and (ID = " + Request.QueryString.Item("sgrupo") + ")";

           oRec = Query( Sql, oRec, oConn  );	
          

          if (oRec.EOF == true)
            { 
              oConn.Close();   
              
                          Response.Redirect("../errorpage.asp?tipo=Error&short=No existe el subgrupo&desc=No existe el subgrupo que usted intenta modificar."); 
            }  
          else
            {
            
              if ((Request.Form("MName").Count != 0) && (Request.Form("Mname") + "" != ""))
                { 
                  oRec.Fields.Item("Name").Value = Request.Form("Mname") + "";
                }
              else Response.Redirect("../errorpage.asp?tipo=Error&short=Nombre no válido&desc=El nombre del subgrupo introducido no es válido o se ha dejado en blanco.");


              if ((Request.Form("Mdesc").Count != 0) && (Request.Form("Mdesc") + "" != ""))
                { 
                  oRec.Fields.Item("Desc").Value = Request.Form("Mdesc") + "";
                }
              else Response.Redirect("../errorpage.asp?tipo=Error&short=Descripción no válida&desc=La descripción del subgrupo introducida no es válida o se ha dejado en blanco.");

              if ((Request.Form("Mtutorid").Count != 0) && (!isNaN(parseInt(Request.Form("Mtutorid") + ""))))
                { 
                  oRec.Fields.Item("Tutor").Value = Request.Form("Mtutorid");
                }
              else Response.Redirect("../errorpage.asp?tipo=Error&short=Tutor no válido&desc=El tutor no ha sido seleccionado apropiadamente. Para seleccionar el tutor, presione el botón Seleccionar y escójalo de la lista.");

              oRec.Fields.Item("Curso").Value = Session("admcurso");

             
              oRec.Update();      
              oConn.Close();   
              
              Response.Redirect("modifySubGrupo.asp?uid=" + uid + "&sgrupo=" + Request.QueryString.Item("sgrupo"));      
            }  
        }
%>

<form name=Delnews action="ConfirmDeleteUrsSubGrupo.asp?uid=<%=uid%>&sgrupo=<%=Request.QueryString.Item("sgrupo")%>" method="post">
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
        <tr> 
          <td width="1%" class="ToolBar" align="center">&nbsp; </td>
          <td width="20%" class="ToolBar" align="center"><b>Identificador</b></td>
          <td width="49%" class="ToolBar" align="center"><b>Nombre</td>
          <td width="30%" class="ToolBar" align="center"><b>Email</b></td>
        </tr>
        <tr> 
<%    
      var clase = "MessageTR1";
      

      oComm.CommandText = "SELECT Usuarios.* " + 
                          "FROM Usuarios INNER JOIN UsuariosSubGrupo ON Usuarios.ID = UsuariosSubGrupo.Usuario " +
                          "WHERE (((UsuariosSubGrupo.Subgrupo)=" + Request.QueryString.Item("sgrupo") + "))";
                          
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
          <td width="20%" class="<%=clase%>"><a href="userDetails.asp?uid=<%=uid%>&id=<%=oRec.Fields.Item("ID").value%>"><%=oRec.Fields.Item("Name").value%></a></td>
          <td width="49%" class="<%=clase%>"><%=oRec.Fields.Item("fullName").value%></td>
          <td width="30%" class="<%=clase%>"><a href="mailto:<%=oRec.Fields.Item("email").value%>" ><%=oRec.Fields.Item("email").value%></a></td>
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
          <tr><td align="center" width="100%" colspan="4" class="MessageTR">No hay alumnos asignados.</td></tr>        
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
        </tr>
      </table>
    </td>
  </tr>
 
</table>
</form>      


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



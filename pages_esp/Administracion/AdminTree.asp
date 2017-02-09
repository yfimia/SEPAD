<%@ Language=JScript %>
<%
  Response.Expires = 0;
%>
<!-- #include file='../../js/user.inc' -->
<!-- #include file='../../libs/response.inc' -->

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

%>
<%
  if (Session("PermissionType") == ADMINISTRATOR)
    {
%>

<html>
<head>
<title>Arbol</title>
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/listspan.css" type="text/css">
<script type="text/javascript" language="javascript" src="../../js/listspan.js"></script>
<script type="text/javascript" language="javascript" src="../../js/treeview.js"></script>

</head>

<body bgcolor="#ffffff" text="#000000" class="TreeViewBody" style="OVERFLOW-Y: scroll">
<table border="0" cellspacing="0" cellpadding="0" width="100%" class="ToolBar">
  <tr>
    <td>
      <table border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><a class="ToolLink" href="javascript:Do_Expand_All()">&nbsp;Expandir&nbsp;</a></td>
          <td>|</td>
          <td><a class="ToolLink" href="javascript:Do_Collapse_All()">&nbsp;Contraer&nbsp;</a></td>
        </tr> 
      </table>
    </td>
  </tr>
</table>

<ul>
  <li class="clsHasKidsCollapsed"><A class=TreeLnk href="../Configure.asp?uid=<%=uid%>" target=mainFrame ><IMG class=IconClass src="../../images/<%=Session("skin")%>/AdminTree/General.gif"><span>Configuración general del sistema</span></a> 
    <ul>
      <li class="clsHasKidsCollapsed"><a class="TreeLnk" target="mainFrame" ><IMG class=IconClass src="../../images/<%=Session("skin")%>/AdminTree/Courses.gif"><span>Manipulación de modalidades</span></a> 
        <ul>
          <li class="clsDocument"><A class=TreeLnk href="../AgregaMod.asp?uid=<%=uid%>" target=mainFrame >Adicionar modalidad académica</a></li>
          <li class="clsDocument"><A class=TreeLnk href="../DeleteMod.asp?uid=<%=uid%>" target=mainFrame >Ver/Eliminar modalidad académica</a></li>
          <!--li class="clsDocument"><A class=TreeLnk href="../CourseManager.asp?uid=<%=uid%>" target=mainFrame >Ver/Eliminar de módulos</a></li>
          <li class="clsDocument"><A class=TreeLnk href="../AgregaCurMod.asp?uid=<%=uid%>" target=mainFrame >Adicionar cursos a un módulo</a></li-->
        
        </ul>

      <li class="clsDocument"><A class=TreeLnk href="../ManipuladorOpciones.asp?uid=<%=uid%>" target=mainFrame >Manipulador de opciones</a></li>

      <li class="clsHasKidsCollapsed"><a class="TreeLnk" target="mainFrame"><IMG class=IconClass src="../../images/<%=Session("skin")%>/AdminTree/UsrMgr.gif"><span>Manipulador de usuarios</span></a> 
        <ul>
          <li class="clsDocument" > <A class=TreeLnk href="../AgregaUser.asp?uid=<%=uid%>" target=mainFrame >Crear usuario</a>
          <li class="clsDocument" > <A class=TreeLnk href="../Usersearch.asp?uid=<%=uid%>" target=mainFrame >Buscar Usuario</a>
          <li class="clsDocument" > <A class=TreeLnk href="../DeleteUser.asp?uid=<%=uid%>" target=mainFrame ><span>Usuarios</span></a>
          <li class="clsDocument" > <A class=TreeLnk href="../agregagroup.asp?uid=<%=uid%>" target=mainFrame >Crear grupo</a>
          <li class="clsDocument" > <A class=TreeLnk href="../Deletegroup.asp?uid=<%=uid%>" target=mainFrame >Grupos de usuarios</a>

          <li class="clsDocument" > <A class=TreeLnk href="../agregaUsergroup.asp?uid=<%=uid%>" target=mainFrame >Adicionar Usuarios a un grupo</a>
          <li class="clsDocument" > <A class=TreeLnk href="../DeleteUsergroup.asp?uid=<%=uid%>" target=mainFrame >Eliminar Usuarios de un grupo</a></li>

        </ul><!--<li class="clsHasKidsCollapsed"><a target="mainFrame" href="../UserManager.asp?uid=<%=uid%>"><img src="../../images/<%=Session("skin")%>/AdminTree/UsrMgr.gif" class="IconClass"><span>Administraci&oacute;n&nbsp;de&nbsp;Usuarios </span></a> 
        <ul>
          <li class="clsDocument"><a class="TreeLnk" target="mainFrame" href="AdminUsers.asp?uid=<%=uid%>"><img src="../../images/<%=Session("skin")%>/AdminTree/Users.gif" class="IconClass">Usuarios</a></li>
          <li class="clsDocument"><a class="TreeLnk" target="mainFrame" href="AdminGroups.asp?uid=<%=uid%>"><img src="../../images/<%=Session("skin")%>/AdminTree/Groups.gif" class="IconClass">Grupos</a></li>
        </ul> 
      </li>-->
      <li class="clsDocument"><a class="TreeLnk" href="../MtrManagerByModality.asp?uid=<%=uid%>" target="mainFrame"><IMG class=IconClass src="../../images/<%=Session("skin")%>/AdminTree/Matriculas.gif"><span>Solicitudes de matrículas</span></a> </l7i>
      

      <li class="clsHasKidsCollapsed"><a class="TreeLnk" target="mainFrame"><IMG class=IconClass src="../../images/<%=Session("skin")%>/AdminTree/inbox.gif"><span>Manipulador de mensajes</span></a> 
        <ul>
          <li class="clsDocument" > <A class=TreeLnk href="../MailManager.asp?uid=<%=uid%>" target=mainFrame >Ordenados por fecha</a>
          <li class="clsDocument" > <A class=TreeLnk href="../MailManagerPara.asp?uid=<%=uid%>" target=mainFrame >Ordenados por destinatario</a></li>
        </ul><!--li class="clsDocument"> <a class="TreeLnk" target="mainFrame" href="../MailManager.asp?uid=<%=uid%>"><img src="../../images/<%=Session("skin")%>/AdminTree/inbox.gif" class="IconClass">Manipulador de Mensajes</a></li -->
      <li class="clsDocument"> <A class=TreeLnk href="../NewsManager.asp?uid=<%=uid%>" target=mainFrame ><IMG class=IconClass src="../../images/<%=Session("skin")%>/AdminTree/News.gif">Manipulador de noticias</a>


      <li class="clsHasKidsCollapsed"><a class="TreeLnk" target="mainFrame"><!--img src="../../images/<%=Session("skin")%>/AdminTree/Matriculas.gif" class="IconClass"--><span>Estadísticas</span></a> 
        <ul>
          <li class="clsDocument" > <A class=TreeLnk href="../logins.asp?uid=<%=uid%>" target=mainFrame ><!--img src="../../images/<%=Session("skin")%>/AdminTree/Acept.gif" class="IconClass"-->Accesos al sistema</a></li>
          <li class="clsDocument" > <A class=TreeLnk href="../sestadisticas.asp?uid=<%=uid%>" target=mainFrame ><!--img src="../../images/<%=Session("skin")%>/AdminTree/Acept.gif" class="IconClass"-->Datos del sistema </a></li>
          <li class="clsDocument" > <A class=TreeLnk href="../festadisticas.asp?uid=<%=uid%>" target=mainFrame ><!--img src="../../images/<%=Session("skin")%>/AdminTree/Acept.gif" class="IconClass"-->Calificaciones</a></li>
          
        </ul><!--li class="clsDocument"> <a class="TreeLnk" target="mainFrame" href="../QySManager.asp?uid=<%=uid%>"><img src="../../images/<%=Session("skin")%>/AdminTree/Q&A.gif" class="IconClass">Manipulador&nbsp;de&nbsp;Quejas&nbsp;y&nbsp;Sugerencias</a></li -->

    </ul>	
  </li>
      <li class="clsDocument"> <A class=TreeLnk href="upload/default.asp" target=mainFrame ><IMG class=IconClass src="../../images/<%=Session("skin")%>/AdminTree/News.gif">Descargas</a>

</ul>
</body>
</html>
<script language='jscript'>
<!--
  Do_Expand_All();
//-->  
</script>


<%
    }
  else
    {
      Response.Redirect("errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);
    }

  }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>
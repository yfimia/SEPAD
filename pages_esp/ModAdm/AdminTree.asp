<%@ Language=JScript %>
<%
  Response.Expires = 0;
%>
<!-- #include file='../../js/user.inc' -->
<!-- #include file='inc/AdminTree.inc' -->

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

%>
<%
if ((Session("PermissionType") == ADMINISTRATOR) || (Session("userID") == Session("admcordinador")))
    {
%>

<html>
<head>

<title><%=TituloPagina%></title>
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/listspan.css" type="text/css">

<script type="text/javascript" language="javascript" src="../../js/listspan.js"></script>
<script type="text/javascript" language="javascript" src="../../js/treeview.js"></script>

</head>

<body bgcolor="#ffffff" text="#000000" class="TreeViewBody" style="OVERFLOW-Y: scroll">
<table border="0" celsspacing="0" cellspadding="0" width="100%" class="ToolBar">
  <tr>
    <td>
      <table border="0" celsspacing="0" cellspadding="0">
        <tr>
          <td><a class="ToolLink" href="javascript:Do_Expand_All()">&nbsp;<%=Expan%>&nbsp;</a></td>
          <td>|</td>
          <td><a class="ToolLink" href="javascript:Do_Collapse_All()">&nbsp;<%=Contr%>&nbsp;</a></td>
        </tr> 
      </table>
    </td>
  </tr>
</table>

<ul>
  <li class="clsHasKidsCollapsed"><A class=TreeLnk href="modifyMod.asp?uid=<%=uid%>" target=mainFrame ><IMG class=IconClass src="../../images/<%=Session("skin")%>/AdminTree/General.gif"><span><%=GModalidadAcadem%></span></a> 
    <ul>
      <li class="clsHasKidsCollapsed"><A class=TreeLnk href="verEliminarAlumnado.asp?uid=<%=uid%>" target=mainFrame ><IMG class=IconClass src="../../images/<%=Session("skin")%>/AdminTree/UsrMgr.gif"><span><%=AlumVerEliminar%></span></a> 
        <ul>
          <li class="clsDocument"><A class=TreeLnk href="MtrUsrCour.asp?uid=<%=uid%>" target=mainFrame ><%=AAlumno%></a></li>
          <li class="clsDocument"><A class=TreeLnk href="MtrGrpCour.asp?uid=<%=uid%>" target=mainFrame ><%=AGrupoAlumno%></a></li>
          <li class="clsDocument"><A class=TreeLnk href="mtrcurmanager.asp?uid=<%=uid%>" target=mainFrame ><%=SolicitudesMat%></a></li>
          <!--li class="clsDocument"><A class=TreeLnk href='../MatriculasManager.asp?uid=<%=uid%>&id=<%=Session("modulo")%>' target=mainFrame ><%=SolicitudesMat%></a></li-->          
        </ul>
      <li class="clsHasKidsCollapsed"><A class=TreeLnk href="verEliminarClaustro.asp?uid=<%=uid%>" target=mainFrame ><IMG class=IconClass src="../../images/<%=Session("skin")%>/AdminTree/Courses.gif"><span><%=ClaustroVerEliminar%></span></a> 
        <ul>
          <li class="clsDocument" > <A class=TreeLnk href="MtrClaustroCour.asp?uid=<%=uid%>" target=mainFrame ><%=DesignarProfesor%></a>
          <li class="clsDocument" > <A class=TreeLnk href="MtrGClaustroCour.asp?uid=<%=uid%>" target=mainFrame ><%=DesignarGrupoProfesores%></a>
        </ul>
      <li class="clsHasKidsCollapsed"><A class=TreeLnk href="verEliminarCurso.asp?uid=<%=uid%>" target=mainFrame ><IMG class=IconClass src="../../images/<%=Session("skin")%>/AdminTree/Matriculas.gif"><span><%=ModVerEliminar%></span></a> 
        <ul>
          <li class="clsDocument" > <A class=TreeLnk href="agregaCurso.asp?uid=<%=uid%>" target=mainFrame ><IMG class=IconClass src="../../images/<%=Session("skin")%>/AdminTree/Acept.gif"><%=AModulo%></a>
          <li class="clsDocument" > <A class=TreeLnk href="CambiaOrdenCurso.asp?uid=<%=uid%>" target=mainFrame ><IMG class=IconClass src="../../images/<%=Session("skin")%>/AdminTree/Acept.gif">Ordenar módulos</a>
        </ul>

    </ul>	
  </li>
</ul>
</body>
<script language='jscript'>
<!--
  
  Do_Expand_All();
//-->  
</script>
</html>

<%
    }
  else
    {
      Response.Redirect("../errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);     
    }

  }
 else
   Response.Redirect("../errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
  //Response.Write(Request.QueryString.Item("uid"));
%>
<%@ Language=JScript %>
<%
  Response.Expires = 0;
%>
<!-- #include file='../../js/user.inc' -->

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

%>
<%
if ((Session("isTutor") > -1) || (Session("PermissionType") == ADMINISTRATOR) || (Session("userID") == Session("tutcursocordinador"))  || (Session("userID") == Request.QueryString.Item("courseOwner")))
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
<table border="0" celsspacing="0" cellspadding="0" width="100%" class="ToolBar">
  <tr>
    <td>
      <table border="0" celsspacing="0" cellspadding="0">
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
  <li class="clsHasKidsCollapsed"><A class=TreeLnk href="Subgrupo.asp?uid=<%=uid%>&courseOwner=<%=Request.QueryString.Item("courseOwner")%>" target=mainFrame ><IMG class=IconClass src="../../images/<%=Session("skin")%>/AdminTree/General.gif"><span>Tutorear subgrupo</span></a> 
    <ul>
          <li class="clsDocument"><A class=TreeLnk href="AdminUpLoadedFiles.asp?uid=<%=uid%>&courseOwner=<%=Request.QueryString.Item("courseOwner")%>" target=mainFrame >Trabajos recibidos</a></li>
          <li class="clsDocument"><A class=TreeLnk href="exerlist.asp?uid=<%=uid%>&courseOwner=<%=Request.QueryString.Item("courseOwner")%>" target=mainFrame >Calificar evaluaciones</a></li>
          <li class="clsDocument"><A class=TreeLnk href="usrvst.asp?uid=<%=uid%>&courseOwner=<%=Request.QueryString.Item("courseOwner")%>" target=mainFrame >Accesos por fecha</a></li>
          <li class="clsDocument"><A class=TreeLnk href="lessonvst.asp?uid=<%=uid%>&courseOwner=<%=Request.QueryString.Item("courseOwner")%>" target=mainFrame >Cantidad de accesos por lecci�n</a></li>
          <li class="clsDocument"><A class=TreeLnk href="cantUsrVst.asp?uid=<%=uid%>&courseOwner=<%=Request.QueryString.Item("courseOwner")%>" target=mainFrame >Cantidad de accesos por alumno</a></li>
          <li class="clsDocument"><A class=TreeLnk href="estadisticas.asp?uid=<%=uid%>&courseOwner=<%=Request.QueryString.Item("courseOwner")%>" target=mainFrame >Calificaciones</a></li>
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
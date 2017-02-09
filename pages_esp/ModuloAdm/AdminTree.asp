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
 if ((Session("PermissionType") == ADMINISTRATOR) || (Session("userID") == Session("admcursocordinador"))  || (Session("userID") == Session("admcursoOwner")))
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
  <li class="clsHasKidsCollapsed"><A class=TreeLnk href="modifyMod.asp?uid=<%=uid%>" target=mainFrame ><IMG class=IconClass src="../../images/<%=Session("skin")%>/AdminTree/General.gif"><span>Generalidades del módulo</span></a> 
    <ul>
      <li class="clsHasKidsCollapsed"><A class=TreeLnk href="verEliminarSubgrupo.asp?uid=<%=uid%>" target=mainFrame ><IMG class=IconClass src="../../images/<%=Session("skin")%>/AdminTree/UsrMgr.gif"><span>Subgrupos tutoreados(Ver/Eliminar)</span></a> 
        <ul>
          <li class="clsDocument"><A class=TreeLnk href="addSubGrupo.asp?uid=<%=uid%>" target=mainFrame >Adicionar subgrupo</a></li>
          <li class="clsDocument"><A class=TreeLnk href="addUsrSubGrupo.asp?uid=<%=uid%>" target=mainFrame >Adicionar alumno</a></li>
        </ul>
      </li>
      <li class="clsHasKidsCollapsed"><IMG class=IconClass src="../../images/<%=Session("skin")%>/AdminTree/UsrMgr.gif"><span>Glosario</span> 
        <ul>
          <li class="clsDocument"><A class=TreeLnk href="AgregarPalabra.asp?uid=<%=uid%>" target=mainFrame >Adicionar palabra</a></li>
          <li class="clsDocument"><A class=TreeLnk href="ModificaEliminaPalabra.asp?uid=<%=uid%>" target=mainFrame >Modificar/Eliminar palabra</a></li>
        </ul>
      </li>
	  <li class="clsHasKidsCollapsed"><IMG class=IconClass src="../../images/<%=Session("skin")%>/AdminTree/UsrMgr.gif"><span>Preguntas Frecuentes</span> 
        <ul>
          <li class="clsDocument"><A class=TreeLnk href="AgregarPregunta.asp?uid=<%=uid%>" target=mainFrame >Adicionar pregunta</a></li>
          <li class="clsDocument"><A class=TreeLnk href="ModificaEliminaPregunta.asp?uid=<%=uid%>" target=mainFrame >Modificar/Eliminar pregunta</a></li>
        </ul>
      </li>
      <li class="clsDocument"><a href="../Upload/default.asp"  target=mainFrame class=TreeLnk><IMG class=IconClass src="../../images/<%=Session("skin")%>/AdminTree/UsrMgr.gif"><span>Descargas</span></a> 
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
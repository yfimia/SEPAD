<%@ Language=JScript %>
<%Response.Expires = -1;%>
<!-- #include file='../../js/adoLibrary.inc' -->
<!-- #include file='../../js/user.inc' -->
<!-- #include file='../../js/change.inc' -->

<%
     function quitCRLF(cad) {
       return (cad);  
     }
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      if (Request.QueryString.Item("curso").count > 0)
        var curso = Request.QueryString.Item("curso") + "";
      if (Request.QueryString.Item("alumn").count > 0)
        var alumn = Request.QueryString.Item("alumn") + "";
      var temp = GetUserData(alumn);
      alumnFullName = temp.fullName;
      var uid = Request.QueryString.Item("uid") + "";
      var ID = -1;
      var titulo = "";
      var calif = 0;
      var obs = "";
      if (Request.QueryString.Item("id").count > 0)
      {
        var dataPath = Application('dataPath');
        var ID = Request.QueryString.Item("id");
        var oConn = MakeConnection( oConn, dataPath );
        var oRec  = Server.CreateObject("ADODB.Recordset");
        Sql = "SELECT * " + 
              "FROM Calif_Ofic " +
              "WHERE ID = " + ID;
        oRec = Query( Sql, oRec, oConn  );
        titulo = oRec.Fields.Item("titulo").value;
        calif = oRec.Fields.Item("calif").value;
        obs = oRec.Fields.Item("obs").value;
      }
      if ((Session("isTutor") > -1) || (Session("PermissionType") == ADMINISTRATOR) || (Session("userID") == Session("tutcursocordinador"))  || (Session("userID") == Request.QueryString.Item("courseOwner")))
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
          <td class="HeaderTable"  align=center><b>Adicionar Calificación Oficial</b></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<br>

<center>
  <b>Curso:   </b><i><%=Session("courseName")%></i><br>
  <b>Alumno:   </b><i><a href="javascript:abreVentana('',350,280,'userDetails.asp?uid=<%=uid%>&courseOwner=<%=Request.QueryString.Item("courseOwner")%>&id=<%=alumn%>', 'yes', 'yes')"><%=alumnFullName%></a></i>
</center>
<br>
 <form name=AddCalif action="ConfirmAddCalif.asp?uid=<%=uid%>&courseOwner=<%=Request.QueryString.Item("courseOwner")%>&ID=<%=ID%>&curso=<%=curso%>&alumn=<%=alumn%>" method="post">
 <input type="hidden" name="ID" value=<%=ID%>>
 <table border="0" cellspacing="1" cellpadding="1" align="center">  
      <tr>
        <td align=center   Class="MessageTR">
          Título:
        </td>
        <td align=left   Class="MessageTR">
          <input name=COtitulo Id=COtitulo type=text maxlength=250 size=60 value="<%=quitCRLF(titulo)%>">
        </td>
      </tr>
      <tr>
        <td align=center   Class="MessageTR">
          Calificación:
        </td>
        <td align=left   Class="MessageTR">
          <input name=COcalif Id=COcalif type=text maxlength=250 size=60 value="<%=quitCRLF(calif)%>">
        </td>
      </tr>
      <tr>
        <td align=center   Class="MessageTR">
          Observaciones:
        </td>
        <td align=left   Class="MessageTR">
          <textarea name=COobs Id=COobs rows="8" cols="58" maxlength=250><%=quitCRLF(obs)%></textarea>
        </td>
      </tr>
      <tr>
        <td colspan=2 align=center Class="Toolbar">
<%
        if (ID == -1)
        {
%>
          <input type=submit value="Adicionar">
<%
        }
        else
        {
%>
          <input type=submit value="Actualizar">
<%
        }
%>
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



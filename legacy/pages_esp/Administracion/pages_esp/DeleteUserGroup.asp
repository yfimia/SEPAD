<%@ Language=JScript %>
<%
  Response.Expires = -1;
%>
<!-- #include file='../js/user.inc' -->

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       


if (Session("PermissionType") == ADMINISTRATOR)
    {
    firstU = "-1";
    firstG = "-1";
    
    where = "";
    if (Request.QueryString.Count >= 2) {
      firstU = Request.QueryString.Item("firstU") + "";
      firstG = Request.QueryString.Item("firstG") + "";
      
      if ((firstU != "-1") && (firstG != "-1")) {
        where = " and ((A.Name >= '" +  firstU + "'))";
      } 
    } 
    
%>

<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<script src="../js/CheckBoxes.js" language="JavaScript"></script>
</HEAD>
<body bgcolor="#FFFFFF" text="#000000">
<form name=Delnews action="ConfirmDeleteUserGroup.asp?uid=<%=uid%>" method="post">
<%    
      var  filePath = Application("dataPath");   
      var  oConn    = Server.CreateObject("ADODB.Connection");
      var  oComm    = Server.CreateObject("ADODB.Command");
      var  oRec     = Server.CreateObject("ADODB.Recordset");
      oRec.CursorLocation = 3;
      oRec.CursorType     = 3;
      oConn.Open(filePath);
      oComm.ActiveConnection = oConn;
      oComm.CommandText = "SELECT  Top " + SHOW_CANT + "  A.Name,B.Name As Grupo,A.ID as UserID,B.ID As GroupID FROM Usuarios As A,Grupos As B,Grupos_de_Usuarios As C Where (A.ID = C.[User]) and (B.ID = C.[Group]) and (B.ID <> 60) and (A.ID <> 580) " +  where + " Order By A.Name, B.Name";
      oRec = oComm.Execute();
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar" align="center"><b>Listado de pertencia de usuarios a grupos</b></td>
  </tr>
  <tr> 
    <td> 
      <table width="100%" border="0" cellspacing="1" cellpadding="0">
        <tr> 
          <td width="5%" class="ToolBar" align="center"></td>
          <td width="35%" class="ToolBar" align="center"><b>Identificador de usuario</b></td>
          <td width="60%" class="ToolBar" align="center"><b>Nombre de Grupo</b></td>
        </tr>
      
<%    
      var clase = "MessageTR1";
      
      var i = 0;
      var lastU = "-1"
      var lastG = "-1"
      
      while (oRec.EOF == false)
        { 
          i = i + 1;
          if (clase == "MessageTR1") clase = "MessageTR"; else clase = "MessageTR1";
          lastU = oRec.Fields.Item("Name").value;          
          lastG = oRec.Fields.Item("grupo").value;          

%>
        <tr> 
          <td width="5%" class="<%=clase%>"><input type="checkbox" id=Check<%=i%> name=check<%=i%> ></td>
          <td width="35%" class="<%=clase%>"><%=oRec.Fields.Item("name").value%></td>
          <td width="60%" class="<%=clase%>"><%=oRec.Fields.Item("grupo").value%></td>
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
          <td nowrap><a href="javascript:Delnews.submit()" class="ToolLink">&nbsp;Eliminar&nbsp;</a></td>
          <td>|</td>
          <td><a href="#" class="ToolLink" onClick="CheckAll(GetObjectCollection( document, 'FORM'))">&nbsp;Seleccionar&nbsp;Todo&nbsp;</a></td>
          <td>|</td>
          <td><a href="Deleteusergroup.asp?uid=<%=uid%>&firstU=<%=lastU%>&firstG=<%=lastG%>" class="ToolLink">&nbsp;Próximos&nbsp;<%=SHOW_CANT%>&nbsp;</a></td>
        </tr>
      </table>
    </td>
  </tr>
</table>

<%
      Session("CoursesList") = oRec;
      Session("Conection") = oConn;
      Session("Command") = oComm;
%>
 </form>
<%
    }
  else  
    {
     Response.Redirect("errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);                           
    }
%>

</BODY>
</HTML>
<%
  }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
  //Response.Write(Request.QueryString.Item("uid"));
%>

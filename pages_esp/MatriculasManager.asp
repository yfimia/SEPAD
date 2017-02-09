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
    first = "-1";
    where = "";
    var idmodulo = Request.QueryString.Item("id");
    if (Request.QueryString.Count >= 3) {
      first = Request.QueryString.Item("first") + "";
      
      if (first != "-1") {
        where = " and (Name > '" +  first + "')";
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

<script language="jscript">
  function Deneg() {
    Delnews.action="ConfirmDeleteMatricula.asp?uid=" + "<%=uid%>&id=<%=idmodulo%>";
    Delnews.submit();
  }
</script>

<form name=Delnews action="ConfirmAceptaMatricula.asp?uid=<%=uid%>&id=<%=idmodulo%>" method="post">
<%    
      var  filePath = Application("dataPath");   
      var  oConn    = Server.CreateObject("ADODB.Connection");
      var  oComm    = Server.CreateObject("ADODB.Command");
      var  oRec     = Server.CreateObject("ADODB.Recordset");
      oRec.CursorLocation = 3;
      oRec.CursorType     = 3;
      oConn.Open(filePath);
      oComm.ActiveConnection = oConn;
      oComm.CommandText = 
                          "SELECT     Top " + SHOW_CANT + "  Mod_Matricula.idmodulo, matricula.* " +
                          "FROM         Mod_Matricula INNER JOIN " +
                          "matricula ON Mod_Matricula.idmatricula = matricula.ID " + 
                          "WHERE     (Mod_Matricula.idmodulo =  " + idmodulo + ") " + where + " Order By matricula.Name " ;     
     oRec = oComm.Execute();    
%>
<table border="0" cellspacing="0" cellpadding="0" class="ToolBar" width="100%">
  <tr>
    <td>
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td nowrap><a href="MtrManagerByModality.asp?uid=<%=uid%>" class="ToolLink">&nbsp;Regresar&nbsp;</a></td>
          <td>|</td>
          <td nowrap><a href="javascript:Delnews.submit()" class="ToolLink">&nbsp;Aceptar&nbsp;</a></td>
          <td>|</td>
          <td nowrap><a href="javascript:Deneg()" class="ToolLink">&nbsp;Denegar&nbsp;</a></td>
          <td>|</td>
          <td><a href="#" class="ToolLink" onClick="CheckAll(GetObjectCollection( document, 'FORM'))">&nbsp;Seleccionar&nbsp;Todo&nbsp;</a></td>
          <!--td>|</td>
          <td><a href="matriculasManager.asp?uid=<%=uid%>&first=<%=last%>" class="ToolLink">&nbsp;Próximos&nbsp;<%=SHOW_CANT%>&nbsp;</a></td-->
        </tr>
      </table>
    </td>
  </tr>
</table>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
<%
   if (Application("AutoMtr") == "1") {
%>  
  <tr> 
    <td  align="left"><b>Nota:</b> 
    El sistema está configurado para que los usuarios se registren directamente. 
    Si desea supervisar el registro de los usuarios y que le lleguen las solicitudes a los administradores para que luego estos las acepten o descarten, desmarque la opción <b>Aceptar automáticamente todas las solicitudes de registro al sistema</b> en la <b>Cofiguración General del Sistema</b>.</td>
  </tr>
  <tr> 
    <td class="ToolBar" align="justified">
  </tr>
  
<%
  }
%>  

  <tr> 
    <td class="ToolBar" align="center"><b>Listado de solicitudes de registro al sistema</b></td>
  </tr>
  
  <tr> 
    <td> 
      <table width="100%" border="0" cellspacing="1" cellpadding="0">
        <tr> 
          <td width="5%" class="ToolBar" align="center"></td>
          <td width="15%" class="ToolBar" align="center"><b>Identificador</b></td>
          <td width="25%" class="ToolBar" align="center"><b>Nombre y apellidos</b></td>
          <td width="25%" class="ToolBar" align="center"><b>Correo eléctronico</b></td>
          <td width="30%" class="ToolBar" align="center"><b>Razón</b></td>
        </tr>
      
<%    
      var clase = "MessageTR1";
      
      var i = 0;
      var last = "-1"
      while (oRec.EOF == false)
        { 
          i = i + 1;
          if (clase == "MessageTR1") clase = "MessageTR"; else clase = "MessageTR1";
          last = oRec.Fields.Item("Name").value;          

%>
        <tr> 
          <td width="5%" class="<%=clase%>"><input type="checkbox" id=Check<%=i%> name=check<%=i%> ></td>
          <td width="15%" class="<%=clase%>"><a href="DetallesMatricula.asp?id=<%=oRec.Fields.Item("id").value%>&uid=<%=uid%>"><%=oRec.Fields.Item("name").value%></a></td>
          <td width="25%" class="<%=clase%>"><%=oRec.Fields.Item("FullName").value%></td>
          <td width="25%" class="<%=clase%>"><a href="mailto:<%=oRec.Fields.Item("email").value%>" ><%=oRec.Fields.Item("email").value%></a></td>
          <td width="30%" class="<%=clase%>"><%=oRec.Fields.Item("Razon").value%></td>
          
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
          <tr><td align="center" width="100%" colspan="5" class="MessageTR">No quedan más elementos.</td></tr>        
<%          
        }  
%>        
<table border="0" cellspacing="0" cellpadding="0" class="ToolBar" width="100%">
  <tr>
    <td>
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td nowrap><a href="MtrManagerByModality.asp?uid=<%=uid%>" class="ToolLink">&nbsp;Regresar&nbsp;</a></td>
          <td>|</td>
          <td nowrap><a href="javascript:Delnews.submit()" class="ToolLink">&nbsp;Aceptar&nbsp;</a></td>
          <td>|</td>
          <td nowrap><a href="javascript:Deneg()" class="ToolLink">&nbsp;Denegar&nbsp;</a></td>
          <td>|</td>
          <td><a href="#" class="ToolLink" onClick="CheckAll(GetObjectCollection( document, 'FORM'))">&nbsp;Seleccionar&nbsp;Todo&nbsp;</a></td>
          <!--td>|</td>
          <td><a href="matriculasManager.asp?uid=<%=uid%>&first=<%=last%>" class="ToolLink">&nbsp;Próximos&nbsp;<%=SHOW_CANT%>&nbsp;</a></td-->
        </tr>
      </table>
    </td>
  </tr>
</table>
<%      

      Session("UserList") = oRec;
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

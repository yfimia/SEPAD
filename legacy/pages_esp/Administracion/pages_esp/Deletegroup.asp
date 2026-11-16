<%@ Language=JScript %>
<%Response.Expires = -1;%>
<!-- #include file='../js/user.inc' -->

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       


if (Session("PermissionType") == ADMINISTRATOR)
    {
    first = "-1";
    where = "";
    if (Request.QueryString.Count >= 2) {
      first = Request.QueryString.Item("first") + "";
      
      if (first != "-1") {
        where = " where (Name > '" +  first + "')";
      } 
    } 
    
%>

<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<script src="../js/CheckBoxes.js" language="JavaScript"></script>
</HEAD>
<body bgcolor="#FFFFFF" text="#000000">
<form name=Delnews action="ConfirmDeleteGroup.asp?uid=<%=uid%>" method="post">
<%    
      var  filePath = Application("dataPath");   
      var  oConn    = Server.CreateObject("ADODB.Connection");
      var  oComm    = Server.CreateObject("ADODB.Command");
      var  oRec     = Server.CreateObject("ADODB.Recordset");
      oRec.CursorLocation = 3;
      oRec.CursorType     = 3;
      oConn.Open(filePath);
      oComm.ActiveConnection = oConn;
      oComm.CommandText = "SELECT  Top " + SHOW_CANT + " * FROM grupos "  +  where + " Order By Name";
      oRec = oComm.Execute();
%>
<table border="0" cellspacing="0" cellpadding="0" class="ToolBar" width="100%">
  <tr>
    <td>
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td nowrap><a href="agregagroup.asp?uid=<%=uid%>" class="ToolLink">&nbsp;Crear&nbsp;</a></td>
          <td>|</td>
          <td nowrap><a href="javascript:Delnews.submit()" class="ToolLink">&nbsp;Eliminar&nbsp;</a></td>
          <td>|</td>
          <td><a href="#" class="ToolLink" onClick="CheckAll(GetObjectCollection( document, 'FORM'))">&nbsp;Seleccionar&nbsp;Todo&nbsp;</a></td>
          <td>|</td>
          <td><a href="javascript:goNext()" class="ToolLink">&nbsp;Próximos&nbsp;<%=SHOW_CANT%>&nbsp;</a></td>
        </tr>
      </table>
    </td>
  </tr>
</table>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar" align="center"><b>Listado de grupos</b></td>
  </tr>
  <tr> 
    <td> 
      <table width="100%" border="0" cellspacing="1" cellpadding="0">
        <tr> 
          <td width="5%" class="ToolBar" align="center"></td>
          <td width="35%" class="ToolBar" align="center"><b>Nombre</b></td>
          <td width="60%" class="ToolBar" align="center"><b>Descripcion</b></td>
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
          <td width="5%" class="<%=clase%>">
<%
  if ((oRec.Fields.Item("ID").value != ADMIN_GROUP) && (oRec.Fields.Item("ID").value != PUB_GROUP) && (oRec.Fields.Item("ID").value != GUEST_GROUP) && (oRec.Fields.Item("ID").value != SEPAD_GROUP)) { 
%>        
          <input type="checkbox" id=Check<%=i%> name=check<%=i%> >
          <input type="hidden" id=ID<%=i%> name="ID<%=i%>" value="<%=oRec.Fields.Item("ID").value%>">
          <input type="hidden" id=Name<%=i%> name="Name<%=i%>"  value="<%=oRec.Fields.Item("Name").value%>" >
          
<% }
%>          
          </td>
          <td width="35%" class="<%=clase%>"><a href="deleteuseronegroup.asp?uid=<%=uid%>&group=<%=oRec.Fields.Item("ID").value%>&groupname=<%=oRec.Fields.Item("Name").value%>"><%=oRec.Fields.Item("name").value%></a></td>
          <td width="60%" class="<%=clase%>"><%=oRec.Fields.Item("description").value%></td>
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

<input type="hidden"  name="Count" value="<%=i%>" >
<input type="hidden"  name="First" value="<%=first%>" >

</table>
<table border="0" cellspacing="0" cellpadding="0" class="ToolBar" width="100%">
  <tr>
    <td>
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td nowrap><a href="agregagroup.asp?uid=<%=uid%>" class="ToolLink">&nbsp;Crear&nbsp;</a></td>
          <td>|</td>
          <td nowrap><a href="javascript:Delnews.submit()" class="ToolLink">&nbsp;Eliminar&nbsp;</a></td>
          <td>|</td>
          <td><a href="#" class="ToolLink" onClick="CheckAll(GetObjectCollection( document, 'FORM'))">&nbsp;Seleccionar&nbsp;Todo&nbsp;</a></td>
          <td>|</td>
          <td><a href="javascript:goNext()" class="ToolLink">&nbsp;Próximos&nbsp;<%=SHOW_CANT%>&nbsp;</a></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<script language="jscript">
  function goNext() {
    location.href ="DeleteGroup.asp?uid=<%=uid%>&first=<%=last%>" 
  }
</script>


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
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
    if (Request.QueryString.Count >= 2) {
      first = Request.QueryString.Item("first") + "";
      
      if (first != "-1") {
        where = " where (Name > '" +  first + "')";
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
    Delnews.action="ConfirmDeleteMatricula.asp?uid=" + "<%=uid%>";
    Delnews.submit();
  }
</script>

<form name=Delnews action="ConfirmAceptaMatricula.asp?uid=<%=uid%>" method="post">
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
                          "SELECT   Top " + SHOW_CANT + " Modulo.Name, COUNT(*) AS Solicitudes, Mod_Matricula.idmodulo " + 
                          "FROM         Mod_Matricula INNER JOIN " + 
                          "matricula ON Mod_Matricula.idmatricula = matricula.ID INNER JOIN " + 
                          "Modulo ON Mod_Matricula.idmodulo = Modulo.ID " + 
                          "GROUP BY Modulo.Name, Mod_Matricula.idmodulo ORDER BY Modulo.Name";      
      oRec = oComm.Execute();    
%>

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
          <td width="70%" class="ToolBar" align="center"><b>Modalidad Académica</b></td>
          <td width="30%" class="ToolBar" align="center"><b>SolicitudesPpendientes</b></td>
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
          <td width="70%" class="<%=clase%>" align="center"><a href="MatriculasManager.asp?id=<%=oRec.Fields.Item("idmodulo").value%>&uid=<%=uid%>"><%=oRec.Fields.Item("name").value%></a></td>
          <td width="30%" class="<%=clase%>" align="center"><%=oRec.Fields.Item("Solicitudes").value%></td>
          
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
          <tr><td align="center" width="100%" colspan="5" class="MessageTR">No quedan modalidades con solicitudes de registro pendientes.</td></tr>        
<%          
        }  
%>        
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

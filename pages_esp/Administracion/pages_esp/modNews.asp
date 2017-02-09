<%@ Language=JavaScript%>
<%
  Response.Expires = -1;
  newsid = Request.QueryString.Item("nid");
  uid = Request.QueryString.Item("uid");
  
%>

<!-- #include file="../js/adoLibrary.inc" -->
<!-- #include file="../js/gettime.inc" -->
<!-- #include file='../js/user.inc' -->

<html>
<head>
<title>Modificar noticias</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/EvalResult.css" type="text/css">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000" style="overflow-y:scroll">
<%
  filePath = Application("dataPath");
  var oConn = Server.CreateObject("ADODB.Connection");
  var oRec  = Server.CreateObject("ADODB.Recordset");
  
  Consulta = "SELECT * FROM Noticias WHERE (ID=" + newsid + ")";
  oConn.open(filePath);
  oRec.open(Consulta, oConn,3,3);

  var Titulo = oRec.Fields.Item("Titulo").value;
  var Cuerpo = oRec.Fields.Item("Cuerpo").value;
  var Imagen = oRec.Fields.Item("Imagen").value;
  var URL    = oRec.Fields.Item("Url").value;
  oRec.Close();
  oConn.Close();
  
%>
<table width="100%" border="0" cellspacing="0" cellpadding="2">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td><img src="../images/<%=Session("skin")%>/NoticiasIMG.gif" width="80" height="54"></td>
          <td class="HeaderTable">Noticias</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td>
      <hr noshade>
    </td>
  </tr>  

<SCRIPT LANGUAGE=javascript>
<!--
function mysubmit() {
  if (courForm.titulo.value != "") { courForm.submit(); }
}
//-->
</SCRIPT>

 <form id=courForm name=courForm action="actualizaNews.asp?nid=<%=newsid%>&uid=<%=uid%>" method="post" >
<table border="0" cellspacing="5" cellpadding="5" align="center">
  <tr>
    <td>
      <table border="0" cellspacing="0" cellpadding="0" class="PrologueTable">
        <tr> 
          <td> 
            <table border="0" cellspacing="1" cellpadding="2">
              <tr align="left"> 
                <td colspan="4" class="ToolBar"> 
                  <table border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td class="ToolBar"><b>Titulo:</b></td>
                      <td class="ToolBar"> 
                        <input type="text" id="titulo" name="titulo" maxlength="60" size="80" class="Edit" value="<%=Titulo%>">
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr> 
                <td width="100%" colspan="4" class="MessageTR1"> 
                  <textarea name="cuerpo" rows="13" cols="100"><%=Cuerpo%></textarea>
                </td>
              </tr>
              <tr> 
                <td class="ToolBar"> 
                  <table border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td class="ToolBar"><b>Imagen:</b></td>
                      <td class="ToolBar">
                        <select name="imagen" class="ComboBox1">
                          <% if (Imagen == "Deportes.gif") {%>
                           <option value="Deportes.gif" selected>Deportes</option>
                          <%} else {%>
                           <option value="Deportes.gif">Deportes</option>
                          <%}%> 
                          <% if (Imagen == "Cultura.gif") {%>
                           <option value="Cultura.gif" selected>Cultura</option>
                          <%} else {%>
                           <option value="Cultura.gif">Cultura</option>
                          <%}%> 
                          <% if (Imagen == "Nacionales.gif") {%>
                           <option value="Nacionales.gif" selected>Ambito Nacional</option>
                          <%} else {%>
                           <option value="Nacionales.gif">Ambito Nacional</option>
                          <%}%> 
                          <% if (Imagen == "Internacionales.gif") {%>
                           <option value="Internacionales.gif" selected>Ambito Internacional</option>
                          <%} else {%>
                           <option value="Internacionales.gif">Ambito Internacional</option>
                          <%}%> 
                          <% if (Imagen == "others.gif") {%>
                          <option value="others.gif" selected>Otras</option>
                          <%} else {%>
                          <option value="others.gif">Otras</option>
                          <%}%> 
                          <% if (Imagen == "No_image.gif") {%>
                          <option value="No_image.gif" selected>Sin Imagen</option>                          
                          <%} else {%>
                          <option value="No_image.gif">Sin Imagen</option>                          
                          <%}%> 
                        </select>
                      </td>
                    </tr>
                  </table>
                </td>
                <td class="ToolBar"> 
                  <table border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td class="ToolBar"><b>URL:</b></td>
                      <td class="ToolBar"> 
                        <input type="text" name="url" size="25" class="Edit" value="<%=URL%>">
                      </td>
                    </tr>
                  </table>
                </td>
                <td colspan="2" class="ToolBar"> 
                  <input type="button" name="Submit" value="Actualizar" onclick="mysubmit()">
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
 </form>
</body>
</html>
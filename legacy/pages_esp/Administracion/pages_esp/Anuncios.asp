<%@language = "JScript"%>
<%
    Response.Expires = -1;
%>

<!-- #include file='../js/adolibrary.inc' -->
<!-- #include file='../js/user.inc' -->
<!-- #include file="../js/gettime.inc" -->

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       
    
      var MAXNEW = Application("NewNews"); //minimo de dias para considerar una noticia como nueva.

    idmod = Request.querystring("idmodulo");
%>
<html>
<head>
<title>
  Anuncios
</title>
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/EvalResult.css" type="text/css">
<script language=jscript src="../js/windows.js"></script>
</head>
<body>
<table width="100%" border="0" cellspacing="0" cellpadding="2">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td width=1% ><img src="../images/<%=Session("skin")%>/NoticiasIMG.gif" width="80" height="54"></td>
          <td width=98% class="HeaderTable">Anuncios</td>
          <td width=1% valign=bottom align=right><a href="javascript:abreVentana('CrearAnuncios', 400, 300, 'CrearAnuncio.asp?uid=<%=uid%>&idmodulo=<%=Session("modulo")%>')" class="ToolLink">Crear&nbsp;Anuncio</td>          
        </tr>
      </table>
    </td>
  </tr>
</table>

<table border=0 cellpadding=1 cellspacing=1 width="100%">

<%
var miid = Request.QueryString("idmodulo") * 1;
var filePath = Application("dataPath");
var oCon = Server.CreateObject("ADODB.Connection");
var oRec = Server.CreateObject("ADODB.Recordset");
var  oComm    = Server.CreateObject("ADODB.Command");

      oCon.Open(filePath);

      oComm.ActiveConnection = oCon;
      
      oComm.CommandText = "DELETE FROM Anuncios WHERE (Expira=" + Application("dtrue") + ") and (FechaExpira <= " + Application("dchar") + Today() + Application("dchar") + ")";
      oComm.Execute();
      //Response.Write("DELETE FROM Anuncios WHERE (Expira=" + Application("dtrue") + ") and (FechaExpira <= " + Application("dchar") + Today() + Application("dchar") + ")");
      

var micon = ("SELECT Anuncios.ID, Anuncios.Titulo, Anuncios.Cuerpo, Anuncios.Fecha_Publicada, Usuarios.Name, Usuarios.ID as UID, Anuncios.Publisher FROM AnunciosModulo, Anuncios, Usuarios WHERE (AnunciosModulo.Modulo=" + miid + ") and (Anuncios.ID=AnunciosModulo.Anuncio) and (Anuncios.Publisher=Usuarios.ID) order by [Fecha_Publicada] Desc");
oRec.Open(micon,oCon,3,3);
 
//Aqui esta la conexion para mostrar todos los anuncios
cont = -1;
while (!oRec.EOF) {
 cont++;
 var IDAnuncio = oRec.Fields.Item("ID");
 
 
%>

<table width="100%" border="0" cellspacing="2" cellpadding="0">
  <tr> 
    <td width="15%"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td align="center">
             
         </td>
        </tr>

        <tr> 
          <td align="center"><%=oRec.Fields.Item('Fecha_Publicada') %></td>
        </tr>

<% 
   var delta = 1;
    delta = CompareDate(oRec.Fields.Item('Fecha_Publicada'),Today());
    //Response.Write("<br>" + oRec.Fields.Item('Fecha_Publicada') + "   "  + Today() + "<br>");
    //Response.Write("<br>" + delta + " <= " +   MAXNEW + "<br>");
    
    if ((delta <=  MAXNEW) && (delta >= 0))
    
      {
%>                
        <tr> 
          <td align="center"><img src="../images/<%=Session("skin")%>/NewNews.gif" width="40" height="15"></td>
        </tr>

<%
      }
%>                  
      </table>
    </td>
    <td valign="top"> 
      <table width="100%" border="0" cellspacing="2" cellpadding="0">
        <tr> 
          <td width="1%"><b>T&iacute;tulo: </b></td>
          <td width="84%">
                 <%=oRec.Fields.Item('Titulo') %>
        </tr>         
                 
       <tr>
           <td><b>Publicado&nbsp;por: </b><a href="javascript:abreVentana('',350,280,'userDetails.asp?uid=<%=uid%>&id=<%=oRec.Fields.Item('UID').value%>', 'yes', 'yes')"><%=oRec.Fields.Item("Name")%></a></td>
       </tr>
                 
        <tr> 
          <td colspan="2"><%=oRec.Fields.Item('Cuerpo') %></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td colspan=2 align="right">
<%
 if ((Session("userID") == ADMIN_USER) || (Session("userID") == Session("cordinador")) || (Session("userID") == oRec.Fields.Item("Publisher"))) {
 %>
   <a href="javascript:abreVentana('MAnuncios', 400, 300, 'MAnuncios.asp?uid=<%=uid%>&idmodulo=<%=idmod%>&idanuncio=<%=IDAnuncio%>')">Modificar</a>
 <%
 }
 if ((Session("userID") == ADMIN_USER) || (Session("userID") == Session("cordinador")) || (Session("userID") == oRec.Fields.Item("Publisher"))) {
 %>
   <a href="DelAnuncios.asp?uid=<%=uid%>&idmodulo=<%=idmod%>&idanuncio=<%=IDAnuncio%>">Eliminar</a>
 <%
 }
%> 
    </td>
  </tr>
  
  <tr> 
    <td colspan="2"> 
      <hr noshade>
    </td>
  </tr>
</table>

<%
 oRec.move(1);
}

if (cont == -1) {
 Response.write('<table cellpaddin="0" cellspacing="0" width="100%">');
 Response.write('  <tr>');
 Response.write('    <td align="center" class="MessageTR" >No hay anuncios que mostrar en este m&oacutedulo</td>');
 Response.write('  </tr>');
 Response.write('</table>');
 }

oRec.Close();
oCon.Close();
%>
</table>
</body>

<html>
<%
 }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>

<%@language = jscript%>
<%
    Response.Expires = -1;
%>
<!-- #include file= '../js/user.inc' -->
<!-- #include file= '../js/adolibrary.inc' -->

<!-- #include file= 'inc/actualizarDatosAnuncios.inc' -->

<html>
<head>
<title>
  Modificar Anuncios
</title>

 <link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css"> 
 <link rel="stylesheet" href="../css/<%=Session("skin")%>/EvalResult.css" type="text/css">

</head>
<%

    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

var filePath = Application("dataPath");
var oCon = Server.CreateObject("ADODB.Connection");
var oRec = Server.CreateObject("ADODB.Recordset");
oCon.open(filePath);

var day = Request.Form("DaysEx") * 1;
var month = Request.Form("MonthEx") + "";
var year = Request.Form("YearEx") + "";
var IdAnunM = Request.Form("IDAnuncio") * 1;
var IdMod = Request.Form("IDModulo") * 1;



oRec.open("SELECT * FROM Anuncios where ID = " + IdAnunM,oCon,3,3);

oRec.Fields.Item("Titulo").value = Request.Form("title") + "";
oRec.Fields.Item("Cuerpo").value = Request.Form("body") + "";
oRec.Fields.Item("Publisher").value = Session("userID");
if ((Request.Form("ExpSoN").Count > 0) && (Request.Form("ExpSoN") + "" == "on")) 
 oRec.Fields.Item("Expira").value = true;
else 
 oRec.Fields.Item("Expira").value = false;

 

  var mydate;
  var day=Request.Form("DaysEx") + "";
  var month=Request.Form("MonthEx") + "";
  var year=Request.Form("YearEx") + "";

var  mydate=month + "/" + day + "/" + year;

oRec.Fields.Item("FechaExpira").value = mydate;


  
oRec.update();
oRec.close();


oCon.close();
%>
<table width='100%' border=0 cellpadding=0 cellspacing=0>
  <tr>
    <td class='ToolBar'>
      <table border=0 cellspacing=0 cellpadding=0 class='ToolBar'>
        <tr>
          <td class='HeaderTable' align=center><b><%=mensage%></b></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<%
Response.Redirect("MAnuncios.asp?uid=" + uid + "&idmodulo=" + IdMod + "&idanuncio=" + IdAnunM);

%>
</html>
<%
 }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>

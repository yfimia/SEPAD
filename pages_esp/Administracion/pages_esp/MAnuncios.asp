<%@language=jscript%>
<%
    Response.Expires = -1;

    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

    idmod = Request.querystring("idmodulo");
    IdAnun = Request.querystring("idanuncio") * 1;
%>

<!-- #include file= '../js/adolibrary.inc' -->
<%

var filePath = Application("dataPath");
var oCon = Server.CreateObject("ADODB.Connection");
var oRec = Server.CreateObject("ADODB.Recordset");
oCon.open(filePath);
oRec.open("SELECT * FROM Anuncios WHERE (ID=" + IdAnun + ")",oCon,3,3);


//Cogiendo datos del anuncio a modificar
while (!oRec.EOF) {
   if (oRec.Fields.Item("ID") == IdAnun) {
          var Tt = oRec.Fields.Item("Titulo").value;
          var Cc = oRec.Fields.Item("Cuerpo").value;
          var Ex = oRec.Fields.Item("Expira").value + "";
          var fecha = Date.parse(oRec.Fields.Item("FechaExpira").value);
          }
 oRec.move(1);          
}

//Response.write(Ff);

oRec.close();


function generadia(fcad)
{
  var fecha = new Date(fcad);
  
//  Response.Write(fecha.toLocaleString());
//  return ("<option>" + fecha.getDate() + "</option>");
  aux = fecha.getDate();
  for (i=1;i<=31;i++)
  {
    if (aux == i) {%><option selected value="<%=i%>"><%=i%></option><%}
    else {%><option value="<%=i%>"><%=i%></option><%}
  }
}

function generames(fcad)
{
  var fecha = new Date(fcad);
  return (fecha.getMonth());
}

function generaanio(fcad)
{
  var fecha = new Date(fcad);
  var today = new Date();
//  Response.Write(fecha.toLocaleString());
//  return ("<option>" + fecha.getDate() + "</option>");
  aux = fecha.getYear();
  
  for (i=today.getYear();i<=today.getYear() + 50;i++)
  {
    if ((aux == i)) {%><option selected value="<%=i%>"><%=i%></option><%}
    else {%><option value="<%=i%>"><%=i%></option><%}
  }
}
%>

<html>
<head>
<title>Modificar Anuncio</title>

<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/EvalResult.css" type="text/css">
<script language=jscript>
 function GetDay(f) {
  aux = new Date(f);
  d = f.getDate();
  return(d);  
 }
</script>

</head>
<body class="MainBody">
<table width="100%" border="0" cellspacing="0" cellpadding="2">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td><img src="../images/<%=Session("skin")%>/NoticiasIMG.gif" width="80" height="54"></td>
          <td class="HeaderTable">Modificar Anuncio</td>
        </tr>
      </table>
    </td>
  </tr>
</table>

 <form name="ModifAnuncio" id="ModifAnuncio" method="Post" action="actualizarDatosAnuncios.asp?uid=<%=uid%>">
   <table border=0 cellpadding=0 cellspacing=0 width="100%">
      <tr class="MessageTR1">
          <td width="20%" align="right">T&iacutetulo:</td>
          <td width="80%"><input type=text class=TextArea name="title" id="title" value="<%=Tt%>"></td>
      </tr>    
      <tr class="MessageTR">
          <td valign="top" align="right" width="20%">Cuerpo:</td>
          <td width="80%"><textarea class=TextArea name="body" id="body" rows=10 ><%=Cc%></textarea></td>
      </tr>    
      <tr class="MessageTR1">
          <td align="right">Expira:</td>
          <td>
             <table border=0 cellpadding=0 cellspacing=0>
               <tr>
                   <td>
                      <%
                      if (Ex == "true") {%>
                       <input type=checkbox id="ExpSoN" name="ExpSoN" checked=true>S&iacute
                      <%} else
                      if (Ex == "false") {%>
                       <input type=checkbox id="ExpSoN" name="ExpSoN">S&iacute
                      <%}%>
                   </td>
               </tr>   
             </table>
          </td>
      </tr>    
      <tr>
          <table border=0 cellpadding=1 cellspacing=1>
            <tr class="MessageTR">
              <td align="right" width="40%">Fecha&nbspque&nbspexpira:</td>
              <td align="left" width="5%">D&iacutea:</td>
              <td width="5%">
                  <select name="DaysEx" id="DaysEx">
                    <%generadia(fecha)%>                  
                    </select>
              </td>
              <td align="left" width="5%">Mes:</td>
              <td width="10%">
                  <select name="MonthEx" id="MonthEx">
                  
                <%var aux = generames(fecha)%>
                <option <%if (aux == 0) {Response.Write("selected")}%> value=1>Enero</option>
                <option <%if (aux == 1) {Response.Write("selected")}%> value=2>Febrero</option>
                <option <%if (aux == 2) {Response.Write("selected")}%> value=3>Marzo</option>
                <option <%if (aux == 3) {Response.Write("selected")}%> value=4>Abril</option>
                <option <%if (aux == 4) {Response.Write("selected")}%> value=5>Mayo</option>
                <option <%if (aux == 5) {Response.Write("selected")}%> value=6>Junio</option>
                <option <%if (aux == 6) {Response.Write("selected")}%> value=7>Julio</option>
                <option <%if (aux == 7) {Response.Write("selected")}%> value=8>Agosto</option>
                <option <%if (aux == 8) {Response.Write("selected")}%> value=9>Septiembre</option>
                <option <%if (aux == 9) {Response.Write("selected")}%> value=10>Octubre</option>
                <option <%if (aux == 10) {Response.Write("selected")}%> value=11>Noviembre</option>
                <option <%if (aux == 11) {Response.Write("selected")}%> value=12>Diciembre</option>
                
                  </select>
              </td>
              <td align="left" width="5%">A&ntildeo:</td>
              <td align="left">
                  <select name="YearEx" id="YearEx">
                    <%generaanio(fecha)%>
                  </select>
              </td>
            </tr> 
          </table>
      </tr>    
      <tr>
         <table border=0 align=center>
           <tr>
              <td><input type=hidden name="IDModulo" id="IDModulo" value=<%=idmod%>></td>
           </tr>
         </table>  
      </tr>
      <tr>
         <table border=0 align=center>
           <tr>
              <td><input type=hidden name="IDAnuncio" id="IDAnuncio" value=<%=IdAnun%>></td>
           </tr>
         </table>  
      </tr>
      <tr>
            <table border=0 cellpadding=1 cellspacing=1 align="center">
              <tr>
               <td width="100%"><input type=submit name="SendInfo" id="SendInfo" value="Aceptar"></td>
              </tr> 
      </tr>
   </table>
 </form>
</body>
</html>
<%
 }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>

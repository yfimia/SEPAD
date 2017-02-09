<%@language = "JScript"%>
<!-- #include file= '../js/adolibrary.inc' -->
<html>
<%

    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

var filePath = Application("dataPath");
var oCon = Server.CreateObject("ADODB.Connection");
var oRec = Server.CreateObject("ADODB.Recordset");
oCon.open(filePath);
oRec.open("SELECT * FROM Anuncios",oCon,3,3);

oRec.addnew();
oRec.Fields.Item("Titulo").value = Request.Form("title") + "";
oRec.Fields.Item("Cuerpo").value = Request.Form("body") + "";
oRec.Fields.Item("Publisher").value = Session("userID");
if ((Request.Form("ExpSoN").Count > 0) && (Request.Form("ExpSoN") + "" == "on")) 
 oRec.Fields.Item("Expira").value = true;

 

  var mydate;
  var day=Request.Form("DaysEx") + "";
  var month=Request.Form("MonthEx") + "";
  var year=Request.Form("YearEx") + "";

var  mydate=month + "/" + day + "/" + year;

oRec.Fields.Item("FechaExpira").value = mydate;

//  mydate=Application("dchar") + CambiarMes(month) + "/" + day + "/" + year + Application("dchar");
  
oRec.update();
oRec.MoveLast();
IdAnuncio = oRec.Fields.Item("ID").value;
oRec.close();

/*
var pp = Session("userID") * 1;

var tt = Request.Form("title") + "";
var cc = Request.Form("body") + "";


oRec.open("SELECT * FROM Anuncios WHERE (Publisher=" + pp + ")", oCon,3,3);
while (!oRec.EOF) {
 if ((oRec.Fields.Item("Publisher") == pp) && (oRec.Fields.Item("Titulo") == tt) && (oRec.Fields.Item("Cuerpo") == cc)) {
   IdAnuncio = oRec.Fields.Item("ID").value * 1;
 }
 oRec.move(1);
}
oRec.close();
*/

oRec.open("SELECT * FROM AnunciosModulo",oCon,3,3);
oRec.addnew();
IdMod = Request.Form("IDModulo") * 1;
oRec.Fields.Item("Modulo").value = IdMod;
oRec.Fields.Item("Anuncio").value = IdAnuncio;
oRec.update();
oRec.close();


/*
Consulta = "UPDATE Anuncios SET FechaExpira=" + mydate + " WHERE (ID= " + IdAnuncio + ")"; 
oRec.open(Consulta, oCon,3,3);
*/


oCon.close();
Response.Redirect("CrearAnuncio.asp?uid=" + uid + "&idmodulo=" + IdMod);
%>
</html>
<%
 }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>

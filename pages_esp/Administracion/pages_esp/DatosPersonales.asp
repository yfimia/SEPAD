<%@ Language=javascript%>

<%Response.Expires = -1;%>

<SCRIPT LANGUAGE=javascript src="../js/user.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function go() 
  {
     if (!isNaN(parseInt(modif.annosgrad.value + "")))  { 
     
        if ((modif.nick.value != "") 
            && (modif.fullname.value != "") 
            && (modif.email.value != ""))
        {   
        if (modif.passw.value == modif.confirm.value)
         {
           if (modif.passw.value != "") 
             modif.passw.value = calcMD5(modif.passw.value);
           modif.confirm.value = modif.passw.value;
           modif.submit();
         }  
        else
          alert(DIFFERENT_PASSWORD_TEXT);
        } else alert("Faltan campos");  
     }
     else  alert("El campo A�os de Graduado debe ser n�merico.");  
        
  }

//-->

</SCRIPT>


<%
function generapais()

{
  var Paises = new Array(
"Afganist�n",
"Albania",
"Alemania",
"Algeria",
"Andorra",
"Angola",
"Anguila",
"Ant�rtida",
"Antigua" + " " + "y" + " " + "Barbuda",
"Antillas" + " " + "Holandesas",
"Arabia" + " " + "Saudita",
"Argentina",
"Armenia",
"Aruba",
"Australia",
"Austria",
"Azerbaiy�n",
"Bahamas",
"Bahrein",
"Bangladesh",
"Barbados",
"Belarusia",
"B�lgica",
"Belice",
"Benin",
"Bermuda",
"Bolivia",
"Bosnia" + " " + "Herzegovina",
"Botswana",
"Bouvet" + " " + "Island",
"Brasil",
"Brunei",
"Bulgaria",
"Burkina" + " " + "Faso",
"Burundi",
"But�n",
"Cabo" + " " + "Verde",
"Cambodia",
"Camer�n",
"Canad�",
"Chile",
"China",
"Chipre",
"Ciudad" + " " + "del" + " " + "Vaticano",
"Cocos" + " " + "(Keeling" + " " + "Islands)",
"Colombia",
"Comoros" + " " + "Islas",
"Congo",
"Corea" + " " + "del" + " " + "Norte",
"Corea" + " " + "del" + " " + "Sur",
"Costa" + " " + "de" + " " + "Marfil",
"Costa" + " " + "Rica",
"Croacia",
"Cuba",
"Dinamarca",
"Djibouti",
"Dominica",
"East" + " " + "Timor",
"Ecuador",
"Egipto",
"El" + " " + "Chad",
"El" + " " + "L�bano",
"El" + " " + "Salvador",
"Emiratos" + " " + "Arabes" + " " + "Unidos",
"Eslovaquia",
"Eslovenia",
"Espa�a",
"Estados" + " " + "Unidos",
"Estonia",
"Etiop�a",
"Faroe" + " " + "Islands",
"Fiji",
"Filipinas",
"Finlandia",
"Francia",
"Gab�n",
"Gambia",
"Georgia",
"Ghana",
"Gibraltar",
"Granada",
"Grecia",
"Groenlandia",
"Guadalupe",
"Guam",
"Guatemala",
"Guinea",
"Guinea" + " " + "Ecuatorial",
"Guinea-Bissau",
"Guyana",
"Guyana" + " " + "Francesa",
"Hait�",
"Heard" + " " + "and" + " " + "McDonald" + " " + "Islands",
"Holanda",
"Honduras",
"Hong" + " " + "Kong",
"Hungr�a",
"India",
"Indonesia",
"Irak",
"Ir�n",
"Irlanda",
"Isla" + " " + "Navidad",
"Isla" + " " + "Norfolk",
"Islandia",
"Islas" + " " + "Caim�n",
"Islas" + " " + "Cook",
"Islas" + " " + "Malvinas",
"Islas" + " " + "Marshall",
"Islas" + " " + "Salom�n",
"Islas" + " " + "V�rgenes",
"Israel",
"Italia",
"Jamaica",
"Jap�n",
"Jordania",
"Kazakhstan",
"Kenya",
"Kiribati",
"Kuwait",
"Kyrgyzstan",
"Laos",
"Latvia",
"Lesotho",
"Liberia",
"Libia",
"Liechtenstein",
"Lituania",
"Luxemburgo",
"Macao",
"Macedonia",
"Madagascar",
"Malasia",
"Malawi",
"Maldiva",
"Mal�",
"Malta",
"Marruecos",
"Martinica",
"Mauricios",
"Mauritania",
"M�xico",
"Micronesia",
"Moldovia",
"M�naco",
"Mongolia",
"Montserrat",
"Mozambique",
"Myanmar",
"Namibia",
"Nauru",
"Nepal",
"Nicaragua",
"Nigeria",
"Niue",
"Northern" + " " + "Mariana" + " " + "Islands",
"Noruega",
"Nueva" + " " + "Caledonia",
"Nueva" + " " + "Guinea",
"Nueva" + " " + "Zelandia",
"Om�n",
"Pakist�n",
"Palau",
"Panam�",
"Paraguay",
"Per�",
"Pitcairn",
"Polinesia" + " " + "Francesa",
"Polonia",
"Portugal",
"Puerto" + " " + "Rico",
"Quatar",
"Reino" + " " + "Unido",
"Rep�blica" + " " + "Arabe" + " " + "Siria",
"Rep�blica" + " " + "Central" + " " + "Africana",
"Rep�blica" + " " + "Checa",
"Rep�blica" + " " + "Dominicana",
"Reunion",
"Ruanda",
"Rumania",
"Rusia",
"Sahara" + " " + "Oeste",
"Saint" + " " + "Kitts" + " " + "and" + " " + "Nevis",
"Samoa",
"Samoa" + " " + "Oeste",
"San" + " " + "Marino",
"San" + " " + "Vicente",
"Santa" + " " + "Helena",
"Santa" + " " + "Luc�a",
"Santo" + " " + "Tom�s" + " " + "y" + " " + "Pr�ncipe",
"Senegal",
"Seychelles",
"Sierra" + " " + "Leona",
"Singapur",
"Somalia",
"Srilanka",
"St." + " " + "Pierre" + " " + "y" + " " + "Miquelon",
"Suazilandia",
"Sud�frica",
"Sud�n",
"Suecia",
"Suiza",
"Surinam",
"Svalbard" + " " + "y" + " " + "Jan" + " " + "Mayen" + " " + "Islands",
"Tailandia",
"Taiw�n",
"Tanzania",
"Tayikist�n",
"Togo",
"Tokelau",
"Tonga",
"Trinidad" + " " + "and" + " " + "Tobago",
"T�nez",
"Turkmenist�n",
"Turks" + " " + "y" + " " + "Caicos" + " " + "Islands",
"Turqu�a",
"Tuvalu",
"Ucrania",
"Uganda",
"Uruguay",
"Uzbekist�n",
"Vanuatu",
"Venezuela",
"Vietnam",
"Wallis" + " " + "y" + " " + "Futuna" + " " + "Islands",
"Yemen",
"Yugoslavia",
"Zaire",
"Zona" + " " + "Neutral",
"");
  
  for (i=0;i<229;i++)
  {
    if (oRec.Fields.Item("Pais").Value == Paises[i])
    {
      %><option selected><%=Paises[i]%></option><%
    }
    else
    {
      %><option><%=Paises[i]%></option><%
    }
  }
}

function generaprovincia()
{
  var Provincias = new Array("Pinar del R�o",
  			     "Ciudad" + " " + "de" + " " + "la" + " " + "Habana",
  			     "La" + " " + "Habana",
  			     "Isla" + " " + "de" + " " + "la" + " " + "Juventud",
  			     "Matanzas",
  			     "Cienfuegos",
  			     "Villa" + " " + "Clara",
  			     "Sancti Sp�ritus",
  			     "Ciego" + " " + "de" + " " + "Avila",
  			     "Camag�ey",
  			     "Las" + " " + "Tunas",
  			     "Holgu�n",
  			     "Granma",
  			     "Santiago" + " " + "de" + " " + "Cuba",
  			     "Guant�namo",
  			     "");
  			     
  for (i=0; i<15; i++)
  {
    if (oRec.Fields.Item("Provincia").Value == Provincias[i]) {%><option value="<%=Provincias[i]%>" selected ><%=Provincias[i]%></option><%}
    else {%><option  value="<%=Provincias[i]%>" ><%=Provincias[i]%></option><%}
  }  
}


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

  var today = new Date();
  var fecha = new Date(fcad);
  var aux = "";
  var x = fecha.getYear();
  if (x < 100) aux = "19" + x; else aux = x + "";
  
  for (var i=today.getYear() - 100;i<=today.getYear();i++)
  {
    if ((aux == i + "")) {%><option selected value="<%=i%>"><%=i%></option><%}
    else {%><option value="<%=i%>"><%=i%></option><%}
    
  }


}
%>

<!-- #include file='../js/user.inc' -->

<%
if (Request.QueryString.Item("uid") + "" == Session("uid"))
{    
  var uid = Request.QueryString.Item("uid") + "";
%>

<HTML>
<HEAD>
<title>Datos Personales</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/EvalResult.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000" style="overflow-y:scroll">   
<SCRIPT LANGUAGE=javascript src="../js/md5.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function modif_onsubmit() 
  {
    if (modif.passw.value != "")
      {modif.passw.value = calcMD5(modif.passw.value);}
    if (modif.confirm.value != "")
      {modif.confirm.value = calcMD5(modif.confirm.value);}
  }

//-->
</SCRIPT>

<%

var  filePath = Application("dataPath");   
var  oConn    = Server.CreateObject("ADODB.Connection");
var  oComm    = Server.CreateObject("ADODB.Command");
var  oRec     = Server.CreateObject("ADODB.Recordset");
oConn.Errors.Clear();
oConn.Open( filePath );
oRec.Open("select * from Usuarios where (ID = " + Session("userID") + ")" ,oConn,3,3);
//var fecha = Date.parse(oRec.Fields.Item("fechaNac").Value + "");
var fecha = oRec.Fields.Item("fechaNac").Value;

%>      

<form name=modif action="ConfirmModifyData.asp?uid=<%=uid%>" method="post" LANGUAGE=javascript onsubmit="return modif_onsubmit()">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td class="HeaderTable"  align=center><b>Modificar datos del usuario</b></td>
        </tr>
      </table>
    </td>
  </tr>
</table>

<table border="0" cellspacing="0" cellpadding="2" class="MessageTR1" width="100%">
  <tr> 
    <td align=center>
      <b>Reglas para completar el formulario </b>
    </td>    
        
  </tr>
  <tr> 
    <td>
	<li>Es obligatorio llenar todos los campos que est&aacute;n en negrita.</li>
	<li>La cuenta de correo debe ser v�lida, pues de especificarse, ser� utilizada por el sistema para notificarle cuando sea aceptado.</li>
	<li>No se aceptaran usuarios con el mismo identificador.</li>
	<br>
	<br>
    </td>
  </tr>
</table>



    <table border="0" cellspacing="1" cellpadding="1" align="center" width="100%">
      <tr width="100%">
        <td Class="MessageTR1" align=right width="40%"><b>Nombre y apellidos:</b></td>
        <td Class="MessageTR1"  align=left width="60%">
          <input type=text id=fullname name=fullname  size=30 maxlength=250  value="<%Response.Write(oRec.Fields.Item ("fullname").Value + "")%>">
        </td>
      </tr>
      <tr width="100%">
        <td  Class="MessageTR" align=right width="40%"><b>Identificador:</b></td>
        <td  Class="MessageTR" align=left width="60%">
        <%Response.Write(oRec.Fields.Item ("name").Value + "")%>
        <input type=hidden id=nick name=nick  size=30 maxlength=50  value="<%Response.Write(oRec.Fields.Item ("name").Value + "")%>">
        </td>
      </tr>
      <tr width="100%">
        <td  Class="MessageTR1" align=right width="40%"><b>Contrase�a:</b></td>
        <td  Class="MessageTR1" align=left width="60%">
            <input type=password id=passw name=passw value="" size=30 maxlength=50 >
        </td>
      </tr>
      <tr width="100%">
        <td  Class="MessageTR" align=right width="40%"><b>Confirme la Contrase�a:</b></td>
        <td  Class="MessageTR" align=left width="60%">
          <input name=confirm Id=confirm type=password size=30 maxlength=250 >
        </td>
      </tr>
      <tr width="100%">
        <td  Class="MessageTR1" align=right width="40%">Tipo de usuario:</td>
        <td  Class="MessageTR1" align=left width="60%">
         <%if (oRec.Fields.Item("permissiontype").Value == ADMINISTRATOR) {Response.Write("Administrador")}
           else {Response.Write("Usuario")}%>
         <input type=hidden id=TipoU name=TipoU  size=30 maxlength=50  value="">
        </td>
      </tr>
      <tr width="100%">
        <td  Class="MessageTR" align=right width="40%"><b>Email:</b></td>
        <td  Class="MessageTR" align=left width="60%">
          <input name=email Id=email type=edit size=30 maxlength=250 size="50" value="<%Response.Write(oRec.Fields.Item ("Email").Value + "")%>">
        </td>
      </tr>
      
      <tr width="100%">
        <td  Class="MessageTR1" align=right width="40%">N&uacute;mero&nbsp;de&nbsp;identidad&nbsp;permanente:</td>
        <td  Class="MessageTR1" align=left width="60%">
          <input name=ci Id=ci type=edit size=30 maxlength=20 value="<%Response.Write(oRec.Fields.Item ("ci").Value + "")%>">
        </td>
      </tr>
      
      <tr width="100%">
        <table border="0" cellspacing="1" cellpadding="1" align="center" width="100%">
          <tr width="100%">
            <td align=left width="40%" class="MessageTR" align="center">Fecha de Nacimiento:</td>
            <td Class="MessageTR" align=right width="10%">D&iacute;a:</td>
            <td Class="MessageTR" align=left width="10%">
              <select name=dia Id=dia type=text>
                <%generadia(fecha)%>
              </select>
            </td>
            <td Class="MessageTR" align=right width="10%">Mes:</td>
            <td Class="MessageTR" align=left width="10%">
              <select name=mes Id=mes type=text>
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
            <td  Class="MessageTR" align=right width="10%">A�o:</td>
            <td  Class="MessageTR" align=left width="10%">
              <select name=anio Id=anio type=text>
                <%generaanio(fecha)%>
              </select>
            </td>
          </tr>  
        </table>
      </tr>     
      <tr width="100%">
        <table border="0" cellspacing="1" cellpadding="1" align="center" width="100%">
          <tr width="100%">
            <td  Class="MessageTR1" align=right width="40%">Sexo:</td>
            <td  Class="MessageTR1" align=left width="60%">
              <select name=sexo Id=sexo type=text>
                <option <%if (oRec.Fields.Item("sexo").Value == true) {Response.Write("selected")}%> value="m">Masculino</option>
                <option <%if (oRec.Fields.Item("sexo").Value == false) {Response.Write("selected")}%> value="f">Femenino</option>
              </select>
            </td>
          </tr>
        </table>
      </tr>
      <tr width="100%">
        <table border="0" cellspacing="1" cellpadding="1" align="center" width="100%">
          <tr width="100%"> 
            <td  Class="MessageTR" align=right width="40%">Direcci&oacute;n:</td>
            <td  Class="MessageTR" align=left width="60%">
              <input name=direccion Id=direccion type=edit height="38" size="30" maxlength="250" value="<%Response.Write(oRec.Fields.Item ("direccion").Value + "")%>">
            </td>
          </tr>
        </table>
      </tr>
      <tr width="100%">
        <table border="0" cellspacing="1" cellpadding="1" align="center" width="100%">
          <tr width="100%">
            <td  Class="MessageTR1" align=right width="35%">C&oacute;digo Postal:</td>
            <td  Class="MessageTR1" align=left width="15%">
              <input name=cp Id=cp type=edit size=15 maxlength=250 value="<%Response.Write(oRec.Fields.Item ("cpostal").Value + "")%>">
            </td>
            <td  Class="MessageTR1" align=right width="35%">Tel&eacute;fono:</td>
            <td  Class="MessageTR1" align=left width="15%">
              <input name=tel Id=tel type=edit size=15 maxlength=250 value="<%Response.Write(oRec.Fields.Item ("telefono").Value + "")%>">
            </td>
          </tr>
        </table>
      </tr>
      </tr>      
      <tr width="100%">
      <table border="0" cellspacing="1" cellpadding="1" align="center" width="100%">
      <tr width="100%">
        <td align=right class="MessageTR" width="40%">Pa�s:</td>        
        <td class="MessageTR" width="60%">
          <select id=pais name=pais onchange=mostrar()>
            <%generapais()%>
          </select>
        </td>
      </tr> 
      <tr width="100%"> 
        <td align=right class="MessageTR1" width="20%">Provincia:</td>        
        <td class="MessageTR1" width="40%">
          <select  name=prov id=prov>
            <%generaprovincia()%>
	  </select>        
        </td>      
      </tr>
      
      <tr width="100%">
        <td  Class="MessageTR" align=right width="40%">T&iacute;tulo:</td>
        <td  Class="MessageTR" align=left width="60%">
          <input name=titulo Id=titulo type=edit size=30 maxlength=250 value="<%=oRec.Fields.Item("titulo").Value%>">
        </td>
      </tr>
      
      <tr width="100%">
        <td  Class="MessageTR1" align=right width="40%">Profesi&oacute;n:</td>
        <td  Class="MessageTR1" align=left width="60%">
          <input name=profesion Id=profesion type=edit size=30 maxlength=250 value="<%=oRec.Fields.Item("profesion").Value%>">
        </td>
      </tr>
      <tr width="100%">
        <td  Class="MessageTR" align=right width="40%">Especialidad:</td>
        <td  Class="MessageTR" align=left width="60%">
          <input name=especialidad Id=especialidad type=edit size=30 maxlength=250 value="<%=oRec.Fields.Item("especialidad").Value%>">
        </td>
      </tr>
      <tr width="100%">
        <td  Class="MessageTR1" align=right width="40%">A�o&nbsp;de&nbsp;graduaci�n:</td>
        <td  Class="MessageTR1" align=left width="60%">
          <input name=annosgrad Id=annosgrad type=edit size=30 maxlength=250 value="<%=oRec.Fields.Item("annosgrad").Value%>">
        </td>
      </tr>
      <tr width="100%">
        <td  Class="MessageTR" align=right width="40%">Instituci&oacute;n:</td>
        <td  Class="MessageTR" align=left width="60%">
          <input name=institucion Id=institucion type=edit size=30 maxlength=250 value="<%=oRec.Fields.Item("institucion").Value%>">
        </td>
      </tr>
      <tr width="100%">
        <table border="0" cellspacing="1" cellpadding="1" align="center" width="100%">
          <tr width="100%">
            <td Class=ToolBar align=left width="100%">Perfiles&nbsp;Laborales</td>
          </tr>
        </table>  
        <table border="0" cellspacing="1" cellpadding="1" align="center" width="100%">
          <tr width="100%">
            <td  Class="MessageTR1" align=right width="40%">Trabajo&nbsp;Actual:</td>
            <td  Class="MessageTR1" align=left width="60%">
              <input name=trabactual Id=trabactual type=edit size=30 maxlength=250 value="<%=oRec.Fields.Item("trabactual").Value%>">
            </td>
          </tr>
          <tr width="100%">
            <td  Class="MessageTR" align=right width="40%">Cargo&nbsp;Actual:</td>
            <td  Class="MessageTR" align=left width="60%">
              <input name=cargoactual Id=cargoactual type=edit size=30 maxlength=250 value="<%=oRec.Fields.Item("cargoactual").Value%>">
            </td>
          </tr>
          <tr width="100%">
            <td  Class="MessageTR1" align=right width="40%">Otros:</td>
            <td  Class="MessageTR1" align=left width="60%">
              <input name=otros Id=otros type=edit size=30 maxlength=250 value="<%=oRec.Fields.Item("otros").Value%>">
            </td>
          </tr> 
        </table>
      </tr>
      <tr width="100%">
        <table border="0" cellspacing="1" cellpadding="1" align="center" width="100%">
          <tr width="100%">
            <td Class=ToolBar align=left width="100%">Perfiles&nbsp;Investigativos</td>
          </tr>
        </table>  
        <table border="0" cellspacing="1" cellpadding="1" align="center" width="100%">
          <tr width="100%">
            <td  Class="MessageTR" align=right width="40%">Categor&iacute;a&nbsp;Docente:</td>
            <td  Class="MessageTR" align=left width="60%">
              <input name=catDoc Id=catDoc type=edit size=30 maxlength=250 value="<%=oRec.Fields.Item("catDoc").Value%>">
            </td>
          </tr>
          <tr width="100%">
            <td  Class="MessageTR1" align=right width="40%">Categor&iacute;a&nbsp;Investigativa:</td>
            <td  Class="MessageTR1" align=left width="60%">
              <input name=catInv Id=catInv type=edit size=30 maxlength=250 value="<%=oRec.Fields.Item("catInv").Value%>">
            </td>
          </tr>
          <tr width="100%">
            <td  Class="MessageTR" align=right width="40%">Categor&iacute;a&nbsp;Cient&iacute;fica:</td>
            <td  Class="MessageTR" align=left width="60%">
              <input name=catCient Id=catCient type=edit size=30 maxlength=250 value="<%=oRec.Fields.Item("catCient").Value%>">
            </td>
          </tr> 
        </table>
      </tr>         
      </table>
      </tr>
      <tr width=100%>
        <table border="0" cellspacing="1" cellpadding="1" align="center" width="100%">
          <tr width="100%">
            <td Class=ToolBar align=left width="100%">Preferencias</td>
          </tr>
        </table>
      </tr>
      
      </table>
      </tr>

      <tr width="100%">
        <table border="0" cellspacing="1" cellpadding="1" align="center" width="100%">
          <tr width="100%">
            <td  Class="MessageTR" align=right width="1%">
             Estilo
            </td>
            <td  Class="MessageTR" align=left width="1%">
               <select id="skin" name="skin" onchange="preview()">
                 <option value="aqua" <%if (oRec.Fields.Item("skin") == "aqua") {%>selected<%}%> > Aqua</option>
                 <option value="fresh" <%if (oRec.Fields.Item("skin") == "fresh") {%>selected<%}%> > Fresh</option>
                 <option value="gold" <%if (oRec.Fields.Item("skin") == "gold") {%>selected<%}%> > Gold</option>
               </select>
            </td>
          </tr>
          <tr>
            <td colspan=2 Class="MessageTR" align=center width="50%">
               <img id="pr" name="pr" border="1px" src='../images/Preview/<%=oRec.Fields.Item("skin")%>.jpg' >
            </td>
            
          </tr>
          
        </table>
      </tr>
      <tr width="100%">
        <table border="0" cellspacing="1" cellpadding="1" align="center" width="100%">
          <tr width="100%">
            <td  Class="MessageTR1" align=left width="5">
             <input name=forw Id=forw type=checkbox size=40 maxlength=250 height="15" <%if (oRec.Fields.Item("forward") == true) {%>checked="on"<%}%>>          
            </td>
            <td  Class="MessageTR1" align=left width="95%">Enviar mensajes internos a su cuenta de correo externa.</td>
          </tr>
        </table>
      </tr>
            
      <tr width="100%">
        <table border="0" cellspacing="1" cellpadding="1" align="center" width="100%">
          <tr width="100%">
            <td width="100%" Class="Toolbar" colspan=2 align=center>
              <input type=button value=Aceptar name=gogo onclick="go()"> 
            </td>      
          </tr>
        </table>
      </tr>
    </table>


</form>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>

function preview() {
   modif.pr.src = '../images/Preview/' + modif.skin.value + '.jpg';
}
</SCRIPT>
     
</body>
<%
}
else
{
  Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
}
%>
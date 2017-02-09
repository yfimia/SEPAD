<%@ Language=JScript %>

<%
  Response.Expires = -1;
  var counter = 0;
%>

<!-- #include file='../js/user.inc' -->


<%
function CheckMod() {
   res = false;
   for (i=1;i<=Request.Form("cuenta");i++) 
    { 
      if (Request.Form("Check" + i) == "on") {
        return true;
      }
    }  
   return res; 
 }                        

function poncursos()
{
  var  filePath = Application("dataPath");   
  var  oConn    = Server.CreateObject("ADODB.Connection");
  var  oComm    = Server.CreateObject("ADODB.Command");
  var  oRec     = Server.CreateObject("ADODB.Recordset");
  oConn.Open( filePath );
  oRec.Open("select ID,Name from Modulo where ((state= " + MOD_ACA_ININSCRIPTION + ") or (state= " + MOD_ACA_INCOURINSC + ")) order by Name",oConn,3,3);
  var clase = 0;
  var i = 0;
  while (oRec.EOF != true)
  {
    if (clase == 0){
      i = i + 1;%>
      <tr width="100%">
        <td Class="MessageTR1" width="5%"><input type=checkbox id="Check<%=i%>" name="Check<%=i%>"></td>
        <td Class="MessageTR1" width="95%">
          <%=oRec.Fields.Item("Name").Value%>
          <input id="Hidden<%=i%>" name="Hidden<%=i%>" type=hidden value="<%=oRec.Fields.Item("ID").Value%>">
        <td>
      </tr>
    <%
      clase = 1;
    }
    else {
      i = i + 1;%>
      <tr width="100%">
        <td Class="MessageTR" width="5%"><input type=checkbox id="Check<%=i%>" name="Check<%=i%>"></td>
        <td Class="MessageTR" width="95%">
          <%=oRec.Fields.Item("Name").Value%>
          <input id="Hidden<%=i%>" name="Hidden<%=i%>" type=hidden value="<%=oRec.Fields.Item("ID").Value%>">
        <td>
      </tr>
    <%
      clase = 0;
    }
    oRec.move(1);
  }
  oConn.Close();
  counter = i;
}

function generapais()
{
  var Paises = new Array(
"Afganistán",
"Albania",
"Alemania",
"Algeria",
"Andorra",
"Angola",
"Anguila",
"Antártida",
"Antigua" + " " + "y" + " " + "Barbuda",
"Antillas" + " " + "Holandesas",
"Arabia" + " " + "Saudita",
"Argentina",
"Armenia",
"Aruba",
"Australia",
"Austria",
"Azerbaiyán",
"Bahamas",
"Bahrein",
"Bangladesh",
"Barbados",
"Belarusia",
"Bélgica",
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
"Bután",
"Cabo" + " " + "Verde",
"Cambodia",
"Camerún",
"Canadá",
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
"El" + " " + "Líbano",
"El" + " " + "Salvador",
"Emiratos" + " " + "Arabes" + " " + "Unidos",
"Eslovaquia",
"Eslovenia",
"España",
"Estados" + " " + "Unidos",
"Estonia",
"Etiopía",
"Faroe" + " " + "Islands",
"Fiji",
"Filipinas",
"Finlandia",
"Francia",
"Gabón",
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
"Haití",
"Heard" + " " + "and" + " " + "McDonald" + " " + "Islands",
"Holanda",
"Honduras",
"Hong" + " " + "Kong",
"Hungría",
"India",
"Indonesia",
"Irak",
"Irán",
"Irlanda",
"Isla" + " " + "Navidad",
"Isla" + " " + "Norfolk",
"Islandia",
"Islas" + " " + "Caimán",
"Islas" + " " + "Cook",
"Islas" + " " + "Malvinas",
"Islas" + " " + "Marshall",
"Islas" + " " + "Salomón",
"Islas" + " " + "Vírgenes",
"Israel",
"Italia",
"Jamaica",
"Japón",
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
"Malí",
"Malta",
"Marruecos",
"Martinica",
"Mauricios",
"Mauritania",
"México",
"Micronesia",
"Moldovia",
"Mónaco",
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
"Omán",
"Pakistán",
"Palau",
"Panamá",
"Paraguay",
"Perú",
"Pitcairn",
"Polinesia" + " " + "Francesa",
"Polonia",
"Portugal",
"Puerto" + " " + "Rico",
"Qatar",
"Reino" + " " + "Unido",
"República" + " " + "Arabe" + " " + "Siria",
"República" + " " + "Central" + " " + "Africana",
"República" + " " + "Checa",
"República" + " " + "Dominicana",
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
"Santa" + " " + "Lucía",
"Santo" + " " + "Tomás" + " " + "y" + " " + "Príncipe",
"Senegal",
"Seychelles",
"Sierra" + " " + "Leona",
"Singapur",
"Somalia",
"Srilanka",
"St." + " " + "Pierre" + " " + "y" + " " + "Miquelon",
"Suazilandia",
"Sudáfrica",
"Sudán",
"Suecia",
"Suiza",
"Surinam",
"Svalbard" + " " + "y" + " " + "Jan" + " " + "Mayen" + " " + "Islands",
"Tailandia",
"Taiwán",
"Tanzania",
"Tayikistán",
"Togo",
"Tokelau",
"Tonga",
"Trinidad" + " " + "and" + " " + "Tobago",
"Túnez",
"Turkmenistán",
"Turks" + " " + "y" + " " + "Caicos" + " " + "Islands",
"Turquía",
"Tuvalu",
"Ucrania",
"Uganda",
"Uruguay",
"Uzbekistán",
"Vanuatu",
"Venezuela",
"Vietnam",
"Wallis" + " " + "y" + " " + "Futuna" + " " + "Islands",
"Yemen",
"Yugoslavia",
"Zaire",
"Zona" + " " + "Neutral",
"");

for (i=0; i<230; i++)
{
  if (Paises[i] == "") {%><option selected value=""><%=Paises[i]%></option><%}
  else {%><option><%=Paises[i]%></option><%}
}
}


function generaprovincia()
{
  var Provincias = new Array("Pinar" + " " + "del" + " " + "Río",
  			     "Ciudad" + " " + "de" + " " + "la" + " " + "Habana",
  			     "La" + " " + "Habana",
  			     "Isla" + " " + "de" + " " + "la" + " " + "Juventud",
  			     "Matanzas",
  			     "Cienfuegos",
  			     "Villa" + " " + "Clara",
  			     "Sancti" + " " + "Spíritus",
  			     "Ciego" + " " + "de" + " " + "Avila",
  			     "Camagüey",
  			     "Las" + " " + "Tunas",
  			     "Holguín","Granma",
  			     "Santiago" + " " + "de" + " " + "Cuba",
  			     "Guantánamo",
  			     "");
  for (i=0; i<16; i++)
  {
    if (Provincias[i] == "") {%><option selected value=""><%=Provincias[i]%></option><%}
    else {%><option value="<%=Provincias[i]%>"><%=Provincias[i]%></option><%}
  }  
}

function genera()
{
  %><option selected value="1">1</option><%
  for (i=2;i<=31;i++)
  {
    %><option value="<%=i%>"><%=i%></option><%
  }
}

function genera2()
{

  var today = new Date();
  
  for (i=today.getYear() - 100;i<=today.getYear();i++)
  {
    %><option selected value="<%=i%>"><%=i%></option><%
  }


}

      function notIsScript(cad)
        {  //alert(cad);
          if (cad != '')
            {
              if ((cad.indexOf('>') != -1) || (cad.indexOf('&#62;') != -1) || (cad.indexOf('&gt;') != -1))
                {
                  return (false); 
                }
              else
                return true;  
            } 
          else
            return true; 
        }

if ((Request.Form("Gname").Count == 0) || 
    (Request.Form("nick").Count == 0)  ||
    (Request.Form("passw").Count == 0) ||
    (Request.Form("Gname") == "")      ||
    (Request.Form("nick") == "")      ||
    (Request.Form("passw") == ""))    
{
%>

<html>
<head>
<title>Solicitud de registro</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<style type="text/css">
<!--
td {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 10px;
	text-decoration: none;
}
-->
</style>
<SCRIPT LANGUAGE=javascript src="../js/md5.js"></script>
<SCRIPT LANGUAGE=javascript src="../js/user.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function go() 
  {
     if (!isNaN(parseInt(newgroup.annosgrad.value + "")))  { 
     
        if ((newgroup.passw.value != "") 
            && (newgroup.confirm.value != "") 
            && (newgroup.nick.value != "") 
            && (newgroup.Gname.value != "") 
            && (newgroup.email.value != ""))
        {   
        if (newgroup.passw.value == newgroup.confirm.value)
         {
           newgroup.passw.value = calcMD5(newgroup.passw.value);
           newgroup.confirm.value = newgroup.passw.value;
           newgroup.submit();
         }  
        else
          alert(DIFFERENT_PASSWORD_TEXT);
        } else alert("Faltan campos");  
     }
     else  alert("El campo Año de Graduación debe ser númerico.");  
        
  }

//-->
</SCRIPT>
</head>

<body bgcolor="#FFFFFF" text="#000000" style="overflow-y:scroll">
<form name=newgroup id=newgroup action="matriculas.asp" method=post LANGUAGE=javascript >
<table border="0" cellspacing="0" cellpadding="2" class="ToolBar" width="100%">
  <tr> 
    <td align=center><b>Solicitud de registro</b></td>
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
	<li>La cuenta de correo debe ser válida, pues de especificarse, será utilizada por el sistema para notificarle cuando sea aceptado.</li>
	<li>No se aceptarán usuarios con el mismo identificador. En caso que de un error reportando que ya existe un usuario en el sistema con el identificador especificado por usted, regrese a este formulario y escoja otro identificador.</li>
	<br>
	<br>
    </td>
  </tr>
</table>

    <table border="0" cellspacing="1" cellpadding="1" align="center" width="100%">
      <tr width="100%">
        <td Class="MessageTR1" align=right width="40%"><b>Nombre y apellidos:</b></td>
        <td Class="MessageTR1"  align=left width="60%">
          <input name=Gname Id=Gname type=text maxlength=250 size=30>
        </td>
      </tr>
      <tr width="100%">
        <td  Class="MessageTR" align=right width="40%"><b>Identificador:</b></td>
        <td  Class="MessageTR" align=left width="60%">
          <input name=nick Id=nick type=text maxlength=250 size=30>
        </td>
      </tr>
      <tr width="100%">
        <td  Class="MessageTR1" align=right width="40%"><b>Contraseña:</b></td>
        <td  Class="MessageTR1" align=left width="60%">
          <input name=passw Id=passw type=password size=30 maxlength=250>
        </td>
      </tr>
      <tr width="100%">
        <td  Class="MessageTR" align=right width="40%"><b>Confirme la Contraseña:</b></td>
        <td  Class="MessageTR" align=left width="60%">
          <input name=confirm Id=confirm type=password size=30 maxlength=250 >
        </td>
      </tr>      
      <tr width="100%">
        <td  Class="MessageTR1" align=right width="40%"><b>Email:</b></td>
        <td  Class="MessageTR1" align=left width="60%">
          <input name=email Id=email type=edit size=30 maxlength=250>
        </td>
      </tr> 
      <tr width="100%">
        <td  Class="MessageTR" align=right width="40%">N&uacute;mero&nbsp;de&nbsp;identidad&nbsp;permanente:</td>
        <td  Class="MessageTR" align=left width="60%">
          <input name=ci Id=ci type=edit size=30 maxlength=20>
        </td>
      </tr>     
      <tr width="100%">
        <table border="0" cellspacing="1" cellpadding="1" align="center" width="100%">
          <tr width="100%">
            <td align=left width="40%" class="MessageTR1" align="center">Fecha&nbsp;de&nbsp;Nacimiento:</td>
            <td Class="MessageTR1" align=right width="10%">D&iacute;a:</td>
            <td Class="MessageTR1" align=left width="10%">
              <select name=dia Id=dia type=text>
                <%genera()%>
              </select>
            </td>
            <td Class="MessageTR1" align=right width="10%">Mes:</td>
            <td Class="MessageTR1" align=left width="10%">
              <select name=mes Id=mes type=text>
                <option selected value=1>Enero</option>
                <option value=2>Febrero</option>
                <option value=3>Marzo</option>
                <option value=4>Abril</option>
                <option value=5>Mayo</option>
                <option value=6>Junio</option>
                <option value=7>Julio</option>
                <option value=8>Agosto</option>
                <option value=9>Septiembre</option>
                <option value=10>Octubre</option>
                <option value=11>Noviembre</option>
                <option value=12>Diciembre</option>
              </select>
            </td>
            <td  Class="MessageTR1" align=right width="10%">Año:</td>
            <td  Class="MessageTR1" align=left width="10%">
              <select name=anio Id=anio type=text>
                <%genera2()%>
              </select>
            </td>
          </tr>  
        </table>
      </tr>           
      <tr width="100%">
        <table border="0" cellspacing="1" cellpadding="1" align="center" width="100%">
          <tr width="100%">
            <td  Class="MessageTR" align=right width="40%">Sexo:</td>
            <td  Class="MessageTR" align=left width="60%">
              <select name=sexo Id=sexo type=text>
                <option selected value="m">Masculino</option>
                <option value="f">Femenino</option>
              </select>
            </td>
          </tr>
        </table>
      </tr>    
      <tr width="100%">
        <table border="0" cellspacing="1" cellpadding="1" align="center" width="100%">
          <tr width="100%"> 
            <td  Class="MessageTR1" align=right width="40%">Direcci&oacute;n:</td>
            <td  Class="MessageTR1" align=left width="60%">
              <input name=direccion Id=direccion type=edit height="38" size="30" maxlength="250">
            </td>
          </tr>
        </table>
      </tr>
      <tr width="100%">
        <table border="0" cellspacing="1" cellpadding="1" align="center" width="100%">
          <tr width="100%">
            <td  Class="MessageTR" align=right width="35%">C&oacute;digo Postal:</td>
            <td  Class="MessageTR" align=left width="15%">
              <input name=cp Id=cp type=edit size=15 maxlength=250>
            </td>
            <td  Class="MessageTR" align=right width="35%">Tel&eacute;fono:</td>
            <td  Class="MessageTR" align=left width="15%">
              <input name=tel Id=tel type=edit size=15 maxlength=250>
            </td>
          </tr>
        </table>
      </tr>      
      <tr width="100%">
      <table border="0" cellspacing="1" cellpadding="1" align="center" width="100%">
      <tr width="100%">
        <td align=right class="MessageTR1" width="40%">País:</td>        
        <td class="MessageTR1" width="60%">
          <select id=pais name=pais >
            <%generapais()%>
          </select>
        </td>
      </tr>
      <tr width="100%"> 
        <td align=right class="MessageTR" width="20%">Provincia:</td>        
        <td class="MessageTR" width="40%">
          <select id=prov name=prov>
            <%generaprovincia()%>
	  </select>        
        </td>      
      </tr>
      <tr width="100%">
        <td  Class="MessageTR1" align=right width="40%">T&iacute;tulo:</td>
        <td  Class="MessageTR1" align=left width="60%">
          <input name=titulo Id=titulo type=edit size=30 maxlength=250>
        </td>
      </tr>
      <tr width="100%">
        <td  Class="MessageTR" align=right width="40%">Profesi&oacute;n:</td>
        <td  Class="MessageTR" align=left width="60%">
          <input name=profesion Id=profesion type=edit size=30 maxlength=250>
        </td>
      </tr>
      <tr width="100%">
        <td  Class="MessageTR1" align=right width="40%">Especialidad:</td>
        <td  Class="MessageTR1" align=left width="60%">
          <input name=especialidad Id=especialidad type=edit size=30 maxlength=250>
        </td>
      </tr>
      <tr width="100%">
        <td  Class="MessageTR" align=right width="40%">Año&nbsp;de&nbsp;graduación:</td>
        <td  Class="MessageTR" align=left width="60%">
          <input name=annosgrad Id=annosgrad type=edit size=30 maxlength=250 value="0">
        </td>
      </tr>
      <tr width="100%">
        <td  Class="MessageTR1" align=right width="40%">Instituci&oacute;n:</td>
        <td  Class="MessageTR1" align=left width="60%">
          <input name=institucion Id=institucion type=edit size=30 maxlength=250>
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
            <td  Class="MessageTR" align=right width="40%">Trabajo&nbsp;Actual:</td>
            <td  Class="MessageTR" align=left width="60%">
              <input name=trabactual Id=trabactual type=edit size=30 maxlength=250>
            </td>
          </tr>
          <tr width="100%">
            <td  Class="MessageTR1" align=right width="40%">Cargo&nbsp;Actual:</td>
            <td  Class="MessageTR1" align=left width="60%">
              <input name=cargoactual Id=cargoactual type=edit size=30 maxlength=250>
            </td>
          </tr>
          <tr width="100%">
            <td  Class="MessageTR" align=right width="40%">Otros:</td>
            <td  Class="MessageTR" align=left width="60%">
              <input name=otros Id=otros type=edit size=30 maxlength=250>
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
            <td  Class="MessageTR1" align=right width="40%">Categor&iacute;a&nbsp;Docente:</td>
            <td  Class="MessageTR1" align=left width="60%">
              <input name=catDoc Id=catDoc type=edit size=30 maxlength=250>
            </td>
          </tr>
          <tr width="100%">
            <td  Class="MessageTR" align=right width="40%">Categor&iacute;a&nbsp;Investigativa:</td>
            <td  Class="MessageTR" align=left width="60%">
              <input name=catInv Id=catInv type=edit size=30 maxlength=250>
            </td>
          </tr>
          <tr width="100%">
            <td  Class="MessageTR1" align=right width="40%">Categor&iacute;a&nbsp;Cient&iacute;fica:</td>
            <td  Class="MessageTR1" align=left width="60%">
              <input name=catCient Id=catCient type=edit size=30 maxlength=250>
            </td>
          </tr> 
        </table>
      </tr>   
      </table>
      </tr>   
      </table>
      </tr>
    </table>
      <table border="0" cellspacing="1" cellpadding="1" align="center" width="100%">
        <tr width="100%"> 
          <td Class="MessageTR" align="right" valign="top" width="40%">Raz&oacute;n&nbsp;por la cual solicita la matrícula:</td>
          <td Class="MessageTR" align="left" width="60%"> 
            <textarea name="razon" Id="razon" rows="8" cols="52"></textarea>
          </td>
        </tr>
    </table>    
      </table>
    </td>
  </tr>
</table>
<table border="0" cellspacing="1" cellpadding="1" align="center" width="100%">
  <tr width="100%">
    <td width="100%" Class="ToolBar" align=center><b>Modalidades&nbsp;Acad&eacute;micas&nbsp;Disponibles</b></td>
  </tr>
  <table border="0" cellspacing="1" cellpadding="1" align="center" width="100%">
    <%poncursos()%>  
  </table>  
</table>

<table border="0" cellspacing="0" cellpadding="2" class="ToolBar" width="100%">
  <tr> 
    <td align="right"> 
      <table border="0" cellspacing="0" cellpadding="2">
        <tr> 
          <td width="50%"><a href="javascript:go()" class="ToolLink">&nbsp;Aceptar&nbsp;</a></td>
          <td width="50%"><a href="javascript:close()" class="ToolLink">&nbsp;Cerrar&nbsp;</a></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<input type=hidden id=cuenta name=cuenta value="<%=counter%>">
</form>

</body>
</html>

<%
        }
      else
        {
        
          if (CheckMod()) {
          if ((notIsScript(Request.Form("Gname") + "") == true) && 
              (notIsScript(Request.Form("nick") + "") == true) && 
              (notIsScript(Request.Form("email") + "") == true) && 
              (notIsScript(Request.Form("razon") + "") == true))  
            {  
              var  filePath = Application("dataPath");   
              var  oConn    = Server.CreateObject("ADODB.Connection");
              var  oComm    = Server.CreateObject("ADODB.Command");
              var  oRec     = Server.CreateObject("ADODB.Recordset");
              oConn.Open( filePath );
  
              if (Application("AutoMtr") == "0") {
                oRec.Open("select * from matricula where (name = '" + Request.Form("nick") + "')",oConn,3,3);
                if (oRec.EOF == false)
                  {
                    oRec.Close();
                    oConn.Close();
                    Response.Redirect("errorpage.asp?tipo=Error&short=" + LOGIN_EXIST_SHORT  + "&desc=" + LOGIN_EXIST_TEXT);     
                  }
                oRec.Close();  
              }

                {
                  oRec.Open("select * from usuarios where (name = '" + Request.Form("nick") + "')",oConn,3,3);
                  if (oRec.EOF == false)
                    {
                      oConn.Close();
                      Response.Redirect("errorpage.asp?tipo=Error&short=" + LOGIN_EXIST_SHORT  + "&desc=" + LOGIN_EXIST_TEXT);     
                    }
                  else  
                    {
                      if (Application("AutoMtr") == "0") {
                        oRec.Close();
                        oRec.Open("select * from matricula",oConn,3,3);
                      }  
                      
                      oRec.AddNew();
                      oRec.Fields.Item("fullName").Value = Request.Form("Gname") + "";
                      oRec.Fields.Item("name").Value = Request.Form("nick") + "";
                      oRec.Fields.Item("password").Value = Request.Form("passw") + "";
                      if ((Request.Form("email") == "") || (Request.Form("email").Count == 0)) 
                        {oRec.Fields.Item("email").Value = "";}
                      else
                        {oRec.Fields.Item("email").Value = Request.Form("email") + "";} 
                        
                        
                      var fecha = new Date (parseInt(Request.Form("anio")), parseInt(Request.Form("mes")), parseInt(Request.Form("dia")));
	              oRec.Fields.Item("fechaNac") = Request.Form("mes") + "/" + Request.Form("dia") + "/"  + Request.Form("anio");  
	              if (Request.Form("sexo") == "m") 
	                {oRec.Fields.Item("sexo").Value = true;}
                      else 
                        {oRec.Fields.Item("sexo").Value = false;}
                      if ((Request.Form("direccion") == "") || (Request.Form("direccion").Count == 0)) 
                        {oRec.Fields.Item("direccion").Value = "";}
                      else
                        {oRec.Fields.Item("direccion").Value = Request.Form("direccion") + "";}  
                      if ((Request.Form("cp") == "") || (Request.Form("cp").Count == 0)) 
                        {oRec.Fields.Item("cpostal").Value = "";}                         
                      else
                        {oRec.Fields.Item("cpostal").Value = Request.Form("cp") + "";}
                      if ((Request.Form("tel") == "") || (Request.Form("tel").Count == 0)) 
                        {oRec.Fields.Item("telefono").Value = "";}                         
                      else
                        {oRec.Fields.Item("telefono").Value = Request.Form("tel") + "";}
                      if ((Request.Form("pais") == "") || (Request.Form("pais").Count == 0)) 
                        {oRec.Fields.Item("Pais").Value = "";}                         
                      else
                        {oRec.Fields.Item("Pais").Value = Request.Form("pais") + "";}    
                      if ((Request.Form("prov") == "") || (Request.Form("prov").Count == 0)) 
                        {oRec.Fields.Item("Provincia").Value = "";}                         
                      else
                        {oRec.Fields.Item("Provincia").Value = Request.Form("prov") + "";}  
                      if ((Request.Form("titulo") == "") || (Request.Form("titulo").Count == 0)) 
                        {oRec.Fields.Item("titulo").Value = "";}                         
                      else
                        {oRec.Fields.Item("titulo").Value = Request.Form("titulo") + "";}
                      if ((Request.Form("institucion") == "") || (Request.Form("institucion").Count == 0)) 
                        {oRec.Fields.Item("Institucion").Value = "";}                         
                      else
                        {oRec.Fields.Item("Institucion").Value = Request.Form("institucion") + "";}
                        
                          
                      if (Application("AutoMtr") == "0") {
                        if ((Request.Form("razon") == "") || (Request.Form("razon").Count == 0)) 
                          {oRec.Fields.Item("razon").Value = "Ninguna";}
                        else
                         {oRec.Fields.Item("razon").Value = Request.Form("razon");}  
                      }   
                      oRec.Update(); 
                      oRec.MoveLast();
                      IDUsuario = oRec.Fields.Item("ID").Value;
                      
                      oRec.Close();

                      if (Application("AutoMtr") == "0") {
                        oRec.Open("select * from Mod_Matricula");
                      }
                      else
                        oRec.Open("select * from TMP_Matriculas_Cursos");  
                      
                      for (i=1;i<=Request.Form("cuenta");i++) 
                      { 
                        if (Request.Form("Check" + i) == "on")
                        {
                          oRec.AddNew();
                          if (Application("AutoMtr") == "0") {
                            oRec.Fields.Item("idmodulo").Value = Request.Form("Hidden" + i);
                            oRec.Fields.Item("idmatricula").Value = IDUsuario;                            
                          }  
                          else {
                            oRec.Fields.Item("Modulo").Value = Request.Form("Hidden" + i);
                            oRec.Fields.Item("User").Value = IDUsuario;                            
                          }  
                          
                          oRec.Update();
                        }
                      }                  
                      oConn.Close();
%>                        
<html>
<head>
<%
  if (Application("AutoMtr") == "1") {
%>                     
<title>Solicitud de registro enviada</title>
<%
  } 
  else {
%>
<title>Solicitud de registro ACEPTADA</title>        
<%
  }
%>             

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<SCRIPT LANGUAGE=javascript>
<!--
 function closewin() { close();}
//-->
</SCRIPT>
<%
                     if (Application("AutoMtr") == "0") {
%>                     

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="3">
  <tr> 
    <td width="1%" rowspan="3" valign="top"> <p align="center"><img src="../images/Grafico.gif" width="100" height="98"></p>
      <p align="center">&nbsp;</p></td>
    <td width="98%"> <p align="center"><font size="3"><strong>Su solicitud de 
        registro ha sido enviada satisfactoriamente. </strong></font></p> </td>
  </tr>
  <tr>
    <td><hr noshade></td>
  </tr>
  <tr> 
    <td width="98%"> <p align="justify">&nbsp;&nbsp;En caso de ser aceptado se 
        le notificará a través de la cuenta de correo electrónico que acaba de 
        especificar. </p>
      <p align="justify">&nbsp;&nbsp;Si la cuenta de correo especificada no es 
        válida usted será el encargado de revisar intentando conectarse al sistema 
        con el identificador y la contraseña especificada. </p></td>
  </tr>
  <tr> 
    <td colspan="2"> <hr noshade> </td>
  </tr>
  <tr> 
    <td colspan="2" align="center"> <input type=Button value=Cerrar onClick='closewin()' id=Button1 name=Button1></td>
  </tr>
</table>

<%
  } 
  else {
%> 

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="3">
  <tr> 
    <td width="1%" rowspan="3" valign="top"> <p align="center"><img src="../images/Grafico.gif" width="100" height="98"></p>
      <p align="center">&nbsp;</p></td>
    <td width="98%"> <p align="center"><font size="3"><strong>Su solicitud de registro ha sido aceptada.</strong></font></p> </td>
  </tr>
  <tr>
    <td><hr noshade></td>
  </tr>
  <tr> 
    <td width="98%"> <p align="justify">Para entrar en el sistema en la página principal en la esquina superior derecha introduzca el identificador y  la contraseña especificada y presiona Aceptar.</p></td>
  </tr>
  <tr> 
    <td colspan="2"> <hr noshade> </td>
  </tr>
  <tr> 
    <td colspan="2" align="center"> <input type=Button value=Cerrar onClick='closewin()' id=Button1 name=Button1></td>
  </tr>
</table>

                  
<%
  }
%>

  
</body>                      
</html>
<%                      
                    }  
                }  
            }    
          else
            { 
              Response.Redirect("errorpage.asp?tipo=Error&short=" + INVALID_CHARACTERS_SHORT  + "&desc=" + INVALID_CHARACTERS_TEXT);     
            } 
          }   
          else {
              Response.Redirect("errorpage.asp?tipo=Error&short=No ha seleccionado una modalidad académica&desc=Debe solicitar matrícula en al menos una modalidad académica.");     
          }  
        }
%></BODY></HTML>
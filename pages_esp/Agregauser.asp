<%@ Language=JScript%>

<%Response.Expires = -1;%>



<%
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
"Quatar",
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
  if (Paises[i] == "Cuba") {%><option selected value="<%=Paises[i]%>"><%=Paises[i]%></option><%}
  else {%><option value="<%=Paises[i]%>"><%=Paises[i]%></option><%}
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
%>
<SCRIPT language=jscript RUNAT=Server>
 function getAtualTime()
 {
  return Date();
 }

</SCRIPT>

<script language=JScript>

function pais_onchange()
{
  if (document.all.item("pais").selectedIndex==53)
  {
    document.all.item("prov").disabled=false;
  }
  else
  {
    document.all.item("prov").selectedIndex=15;
    document.all.item("prov").disabled=true;
  }
}

function extract_onclick(carnet)
{
  
  today = new Date();
  
  if (carnet < 10000000000) 
  {
    alert('El Número de Identidad debe tener 11 dígitos.');
  }
  else
  {
    aux=Math.round(carnet / 1000000000);
    temp = 1900 + aux;    
    document.all.item("anio").selectedIndex=temp-today.getYear()+100;
    
    if (carnet-aux*1000000000 < 0)
    {
      carnet=carnet-(aux-1)*1000000000;
    }
    else
    {
      carnet=carnet-aux*1000000000
    }    
    
    aux=Math.round(carnet / 10000000);
    document.all.item("mes").selectedIndex=aux-1;
    
    if (carnet-aux*10000000 < 0)
    {
      carnet=carnet-(aux-1)*10000000;
    }
    else
    {
      carnet=carnet-aux*10000000;
    }
    
    aux=Math.round(carnet / 100000);
    document.all.item("dia").selectedIndex=aux-1;
    
    if (carnet-aux*100000 < 0)
    {
      carnet=carnet-(aux-1)*100000;
    }
    else
    {
      carnet=carnet-aux*100000
    }
    aux=Math.round(carnet / 100);   
    if (carnet-aux*100 < 0)
    {
      carnet=carnet-(aux-1)*100;
    }
    else
    {
      carnet=carnet-aux*100;
    }    
    aux=Math.round(carnet / 10);
    if (carnet-aux*10 < 0)
    {
      carnet=aux-1;
    }
    else
    {
      carnet=aux;
    }
    if ((carnet==0)||(carnet==2)||(carnet==4)||(carnet==6)||(carnet==8))
    {
      document.all.item("sexo").selectedIndex=0;
    }
    else
    {
      document.all.item("sexo").selectedIndex=1;
    }
    
  }
}

</script>

<SCRIPT LANGUAGE=JScript RUNAT=Server> 
// This is a definition for the procedure PrintDate. 
function MomentDate() 
{ 
  var x;
  x = new Date();
  return x; 
}
</SCRIPT> 

<%
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
  
  for (var i=today.getYear() - 100;i<today.getYear();i++){%><option value="<%=i%>"><%=i%></option><%}
  %><option selected value="<%=i%>"><%=i%></option><%
    
}
%>


<%
  Response.Expires = -1;
%>
<!-- #include file='../js/adolibrary.inc' -->
<!-- #include file='../js/user.inc' -->

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

%>

<%
if (Session("PermissionType") == ADMINISTRATOR)
    {
      if ((Request.Form("Gname").Count == 0) || 
          (Request.Form("nick").Count == 0)  ||
          (Request.Form("passw").Count == 0) ||
	  (Request.Form("email").Count == 0) ||
	  (Request.Form("email") == "")      ||
          (Request.Form("Gname") == "")      ||
          (Request.Form("nick") == "")      ||
          (Request.Form("passw") == ""))     
        {
%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/EvalResult.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<SCRIPT LANGUAGE=javascript src="../js/md5.js"></script>
<SCRIPT LANGUAGE=javascript src="../js/user.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>


function newgroup_onsubmit() 
 {
     if (email_onblur(document.all.item("email").value))
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
     else  alert("El campo Años de Graduado debe ser númerico.");  
     }
 }
 
function email_onblur(correo)
{ 
  aux = new String(correo);
  var aux1=aux.indexOf("@",1);
  if (aux1!=-1)
  { 
    var aux2=aux.indexOf(".",1); 
    if (aux2==-1)
    {
      alert('La dirección de correo especificada no es válida.');
      document.all.item("email").focus;
      return false;
    }
    else return true;
  }
  else
  {
    alert('La dirección de correo especificada no es válida.');
    document.all.item("email").focus;
    return false;
  }
  
}

</SCRIPT>



 <form name=newgroup id=newgroup action="agregauser.asp?uid=<%=uid%>" method=post  LANGUAGE=javascript >
 
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td class="HeaderTable" align=center><b>Adicionar usuario</b></td>
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
	<li>No se aceptaran usuarios con el mismo identificador.</li>
	<br>
	<br>
    </td>
  </tr>
</table>

    <table border="0" cellspacing="1" cellpadding="1" align="center" width="100%">
      <tr width="100%">
        <td Class="MessageTR1" align=right width="40%"><b>Nombre y apellidos:</b></td>
        <td Class="MessageTR1"  align=left width="60%" colspan="2">
          <input name=Gname Id=Gname type=text maxlength=250 size=30>
        </td>
      </tr>
      <tr width="100%">
        <td  Class="MessageTR" align=right width="40%"><b>Identificador:</b></td>
        <td  Class="MessageTR" align=left width="60%" colspan="2">
          <input name=nick Id=nick type=text maxlength=16 size=30>
        </td>
      </tr>
      <tr width="100%">
        <td  Class="MessageTR1" align=right width="40%"><b>Contraseña:</b></td>
        <td  Class="MessageTR1" align=left width="60%" colspan="2">
          <input name=passw Id=passw type=password size=30 maxlength=250>
        </td>
      </tr>
      <tr width="100%">
        <td  Class="MessageTR" align=right width="40%"><b>Confirme la Contraseña:</b></td>
        <td  Class="MessageTR" align=left width="60%" colspan="2">
          <input name=confirm Id=confirm type=password size=30 maxlength=250 >
        </td>
      </tr>      
      <tr width="100%">
        <td  Class="MessageTR1" align=right width="40%"><b>Tipo de usuario:</b></td>
        <td  Class="MessageTR1" align=left width="60%" colspan="2">
          <Select Class="MCombo" name=TipoU Id=TipoU>
            <option selected>Usuario</option>
            <option>Administrador</option>
          </select>
        </td>
      </tr>
      <tr width="100%">
        <td  Class="MessageTR" align=right width="40%"><b>Email:</b></td>
        <td  Class="MessageTR" align=left width="60%" colspan="2">
          <input name=email Id=email type=edit size=30 maxlength=250 size="50" onblur="email_onblur(value)">
        </td>
      </tr>
      <tr width="100%">
        <td  Class="MessageTR1" align=right width="40%">N&uacute;mero&nbsp;de&nbsp;identidad&nbsp;permanente:</td>
        <td  Class="MessageTR1" align=left width="60%">
          <input name=ci Id=ci type=edit size=30 maxlength=11>
        </td>
        <td Class="MessageTR1" width="30%">
          <input name=extract Id=extract type=button value="Generar" onclick="extract_onclick(ci.value)">
        </td>
      </tr>      
      <tr width="100%">
        <table border="0" cellspacing="1" cellpadding="1" align="center" width="100%">
          <tr width="100%">
            <td align=left width="40%" class="MessageTR" align="center">Fecha de Nacimiento:</td>
            <td Class="MessageTR" align=right width="10%">D&iacute;a:</td>
            <td Class="MessageTR" align=left width="10%">
              <select name=dia Id=dia type=text>
                <%genera()%>
              </select>
            </td>
            <td Class="MessageTR" align=right width="10%">Mes:</td>
            <td Class="MessageTR" align=left width="10%">
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
            <td  Class="MessageTR" align=right width="10%">Año:</td>
            <td  Class="MessageTR" align=left width="10%">
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
            <td  Class="MessageTR1" align=right width="40%">Sexo:</td>
            <td  Class="MessageTR1" align=left width="60%">
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
            <td  Class="MessageTR" align=right width="40%">Direcci&oacute;n:</td>
            <td  Class="MessageTR" align=left width="60%">
              <input name=direccion Id=direccion type=edit height="38" size="30" maxlength="250">
            </td>
          </tr>
        </table>
      </tr>
      <tr width="100%">
        <table border="0" cellspacing="1" cellpadding="1" align="center" width="100%">
          <tr width="100%">
            <td  Class="MessageTR1" align=right width="35%">C&oacute;digo Postal:</td>
            <td  Class="MessageTR1" align=left width="15%">
              <input name=cp Id=cp type=edit size=15 maxlength=250>
            </td>
            <td  Class="MessageTR1" align=right width="35%">Tel&eacute;fono:</td>
            <td  Class="MessageTR1" align=left width="15%">
              <input name=tel Id=tel type=edit size=15 maxlength=250>
            </td>
          </tr>
        </table>
      </tr>      
      <tr width="100%">
      <table border="0" cellspacing="1" cellpadding="1" align="center" width="100%">
      <tr width="100%">
        <td align=right class="MessageTR" width="40%">País:</td>        
        <td class="MessageTR" width="60%">
          <select id=pais name=pais onchange="pais_onchange()">
            <%generapais()%>
          </select>
        </td>
      </tr>
      <tr width="100%"> 
        <td align=right class="MessageTR1" width="20%">Provincia:</td>        
        <td class="MessageTR1" width="40%">
          <select id=prov name=prov>
            <%generaprovincia()%>
	  </select>        
        </td>      
      </tr>
      <tr width="100%">
        <td  Class="MessageTR" align=right width="40%">T&iacute;tulo:</td>
        <td  Class="MessageTR" align=left width="60%">
          <input name=titulo Id=titulo type=edit size=30 maxlength=250>
        </td>
      </tr>
      
      <tr width="100%">
        <td  Class="MessageTR1" align=right width="40%">Profesi&oacute;n:</td>
        <td  Class="MessageTR1" align=left width="60%">
          <input name=profesion Id=profesion type=edit size=30 maxlength=250>
        </td>
      </tr>
      <tr width="100%">
        <td  Class="MessageTR" align=right width="40%">Especialidad:</td>
        <td  Class="MessageTR" align=left width="60%">
          <input name=especialidad Id=especialidad type=edit size=30 maxlength=250>
        </td>
      </tr>
      <tr width="100%">
        <td  Class="MessageTR1" align=right width="40%">Año&nbsp;de&nbsp;graduación:</td>
        <td  Class="MessageTR1" align=left width="60%">
          <input name=annosgrad Id=annosgrad type=edit size=30 maxlength=250 value="0">
        </td>
      </tr>
      <tr width="100%">
        <td  Class="MessageTR" align=right width="40%">Instituci&oacute;n:</td>
        <td  Class="MessageTR" align=left width="60%">
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
            <td  Class="MessageTR1" align=right width="40%">Trabajo&nbsp;Actual:</td>
            <td  Class="MessageTR1" align=left width="60%">
              <input name=trabactual Id=trabactual type=edit size=30 maxlength=250>
            </td>
          </tr>
          <tr width="100%">
            <td  Class="MessageTR" align=right width="40%">Cargo&nbsp;Actual:</td>
            <td  Class="MessageTR" align=left width="60%">
              <input name=cargoactual Id=cargoactual type=edit size=30 maxlength=250>
            </td>
          </tr>
          <tr width="100%">
            <td  Class="MessageTR1" align=right width="40%">Otros:</td>
            <td  Class="MessageTR1" align=left width="60%">
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
            <td  Class="MessageTR" align=right width="40%">Categor&iacute;a&nbsp;Docente:</td>
            <td  Class="MessageTR" align=left width="60%">
              <input name=catDoc Id=catDoc type=edit size=30 maxlength=250>
            </td>
          </tr>
          <tr width="100%">
            <td  Class="MessageTR1" align=right width="40%">Categor&iacute;a&nbsp;Investigativa:</td>
            <td  Class="MessageTR1" align=left width="60%">
              <input name=catInv Id=catInv type=edit size=30 maxlength=250>
            </td>
          </tr>
          <tr width="100%">
            <td  Class="MessageTR" align=right width="40%">Categor&iacute;a&nbsp;Cient&iacute;fica:</td>
            <td  Class="MessageTR" align=left width="60%">
              <input name=catCient Id=catCient type=edit size=30 maxlength=250>
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
      <tr width="100%">
        <table border="0" cellspacing="1" cellpadding="1" align="center" width="100%">
          <tr width="100%">
            <td  Class="MessageTR1" align=left width="5">
             <input name=forw Id=forw type=checkbox size=40 maxlength=250 height="15">          
            </td>
            <td  Class="MessageTR1" align=left width="95%">Enviar mensajes internos a su cuenta de correo externa.</td>
          </tr>
        </table>
      </tr>
      <tr width="100%">
        <table border="0" cellspacing="1" cellpadding="1" align="center" width="100%">
          <tr width="100%">
            <td  Class="Toolbar" colspan=2 align=center width="100%">
              <input type=button value="Agregar" id=submit1 name=submit1 onclick="newgroup_onsubmit()">
            </td>
          </tr>
        </table>
      </tr>    
      </table>
      </tr>
    </table>
  </form>

  
<%
        }
      else
        {
          var  filePath = Application("dataPath");   
          var  oConn    = Server.CreateObject("ADODB.Connection");
          var  oComm    = Server.CreateObject("ADODB.Command");
          var  oRec     = Server.CreateObject("ADODB.Recordset");
          oConn.Open( filePath );
          oRec.Open("select * from Usuarios where (name = '" + Request.Form("nick") + "')",oConn,3,3);
          if (oRec.EOF == false)
            {
              oConn.Close();
              
              Response.Redirect("errorpage.asp?tipo=Error&short=" + LOGIN_EXIST_SHORT  + "&desc=" + LOGIN_EXIST_TEXT);
            }  
          else if (Request.Form("passw") + "" != Request.Form("confirm") + "")
            {
                oConn.Close();
                Response.Redirect("errorpage.asp?tipo=Error&short=" + DIFFERENT_PASSWORD_SHORT  + "&desc=" + DIFFERENT_PASSWORD_TEXT);
            }                       
          else
            {
              oConn.Close();
              oConn.Errors.Clear();
              oConn.Open( filePath );
              oRec.Open("select * from Usuarios",oConn,3,3);
              oRec.AddNew();
              oRec.Fields.Item("fullName").Value = Request.Form("Gname") + "";
              oRec.Fields.Item("name").Value = Request.Form("nick") + "";
              oRec.Fields.Item("password").Value = Request.Form("passw") + "";
              oRec.Fields.Item("email").Value = Request.Form("email") + "";
              oRec.Fields.Item("Institucion").Value = Request.Form("institucion") + "";                
              oRec.Fields.Item("telefono").Value = Request.Form("tel") + "";
              oRec.Fields.Item("direccion").Value = Request.Form("direccion") + "";
              oRec.Fields.Item("cpostal").Value = Request.Form("cp") + "";
              oRec.Fields.Item("titulo").Value = Request.Form("titulo") + "";   
              oRec.Fields.Item("Pais").Value = Request.Form("pais") + "";
              oRec.Fields.Item("Provincia").Value = Request.Form("prov") + "";
              oRec.Fields.Item("ci").Value = Request.Form("ci") + "";
              oRec.Fields.Item("profesion").Value = Request.Form("profesion") + "";
              oRec.Fields.Item("especialidad").Value = Request.Form("especialidad") + "";
              
              if (!isNaN(parseInt(Request.Form("annosgrad") + ""))) 
              {
                oRec.Fields.Item("annosgrad").Value = Request.Form("annosgrad") + "";
              }
              else
              {
                Response.Redirect("errorpage.asp?tipo=Error&short=Campo númerico no válido&desc=El campo Años de Graduado debe ser númerico.");
              }  
              
              oRec.Fields.Item("trabactual").Value = Request.Form("trabactual") + "";
              oRec.Fields.Item("cargoactual").Value = Request.Form("cargoactual") + "";
              oRec.Fields.Item("otros").Value = Request.Form("otros") + "";
              oRec.Fields.Item("catDoc").Value = Request.Form("catDoc") + "";
              oRec.Fields.Item("catInv").Value = Request.Form("catInv") + "";
              oRec.Fields.Item("catCient").Value = Request.Form("catCient") + "";
              
              oRec.Fields.Item("Logins").Value = 0;
              
 	      //oRec.Fields.Item("fechaIngreso").Value = getActualTime() + "";
	      //oRec.Fields.Item("lastLogin").Value = getActualTime() + "";
	      var fecha = new Date (parseInt(Request.Form("anio")), parseInt(Request.Form("mes")), parseInt(Request.Form("dia")));
	      oRec.Fields.Item("fechaNac") = Request.Form("mes") + "/" + Request.Form("dia") + "/"  + Request.Form("anio");
	      //Response.Write(fecha.toLocaleString());
	      //fecha.toLocaleString();
	      //z = fecha.toLocaleString();
              var group = -1;
              
             if (Request.Form("sexo") == "m") {oRec.Fields.Item("sexo").Value = true;}
             else {oRec.Fields.Item("sexo").Value = false;}
             if (Request.Form("forw") == "on") 
             {
               oRec.Fields.Item("forward").Value = true;
             }
             else 
             {
               oRec.Fields.Item("forward").Value = false;
             }
             
              if (Request.Form("TipoU") == "Administrador")
                {
                  oRec.Fields.Item("PermissionType").Value = ADMINISTRATOR;
                  group = ADMIN_GROUP;
                }
                else  
                  { 
                    oRec.Fields.Item("PermissionType").Value = USER;
                  }  

              oRec.Update(); 
              oConn.Close();   
              var IID = 0;
              
              oConn.Open( filePath );
              oRec.Open("select * from Usuarios where (Name = '" + Request.Form("nick") + "')",oConn,3,3);
              if (oRec.EOF == false)
                { IID = oRec.Fields.Item("ID").Value; } 
              oConn.Close();   

              
              AddUserToGroup(IID, SEPAD_GROUP);
              if (group != -1) AddUserToGroup(IID, group);
              
              Response.Redirect("AgregaUser.asp?uid=" + uid);      
              
            }
        }
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

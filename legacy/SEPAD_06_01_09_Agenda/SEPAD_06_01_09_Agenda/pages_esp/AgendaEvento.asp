<%@ Language=jScript %>
<%
Response.Expires = -1;
%>
<!-- #include file='inc/AgendaEvento.inc' -->
<!-- #include file='../js/adolibrary.inc' -->
<!-- #include file='../js/user.inc' -->
<!-- #include file='../js/date_funtions.inc' -->
<%
	var sAgendaFecha = new String(Session("sAgendaFecha"));
	var dia1 = dia2 = sAgendaFecha.substr (6,2);
	var mes1 = mes2 = sAgendaFecha.substr (4,2);
	var anho1 = anho2 = sAgendaFecha.substr (0,4);
	var hora1 = hora2 = "08";
	var minutos1 = minutos2 = "00";
	
	var anho_act = anho1 - 0;
	
	var Title = NuevaActividad;

if (Request.QueryString.Item("uid") + "" == Session("uid"))
{    
	var uid = Request.QueryString.Item("uid") + "";       

	var oConn = Server.CreateObject("ADODB.Connection");  
	var filePath = Application("filePath");
	oConn = MakeConnection( oConn, filePath );

	if (Request.QueryString.Item("ID").Count != 0)
	{
		var rsEvento = Server.CreateObject("ADODB.Recordset");
		var sSQL = "SELECT " + 
				"SUBSTRING(CONVERT(char(10), DateBegin, 120), 9, 2) AS dia1, SUBSTRING(CONVERT(char(10), DateBegin, 120), 6, 2) AS mes1, CONVERT(char(4), " +
				"DateBegin, 120) AS anho1, CONVERT(char(2), DateBegin, 114) AS hora1, SUBSTRING(CONVERT(char(10), DateBegin, 114), 4, 2) AS minutos1, " + 
				"SUBSTRING(CONVERT(char(10), DateEnd, 120), 9, 2) AS dia2, SUBSTRING(CONVERT(char(10), DateEnd, 120), 6, 2) AS mes2, CONVERT(char(4), " + 
                "DateEnd, 120) AS anho2, CONVERT(char(2), DateEnd, 114) AS hora2, SUBSTRING(CONVERT(char(10), DateEnd, 114), 4, 2) AS minutos2, * FROM AgendaEventos WHERE AgendaEventos.[ID]='" + 
				Request.QueryString.Item("ID") + "'";
						
		rsEvento = Query( sSQL, rsEvento, oConn  );

		var ID = Request.QueryString("ID") + "";
		var OwnerId = rsEvento.Fields("OwnerId").Value ;
		var OwnerName = rsEvento.Fields ("OwnerName").Value;
		var ToName = rsEvento.Fields ("ToName").Value;
		var Subject = Title = rsEvento.Fields ("Subject").Value;
		var Content = rsEvento.Fields ("Content").Value;
		var Priority = rsEvento.Fields ("Priority").Value;
		var DateBegin = rsEvento.Fields ("DateBegin").Value;
		var DateEnd = rsEvento.Fields ("DateEnd").Value ;
		var EventType = rsEvento.Fields ("EventType").Value ;
	
		if (DateBegin + "" != 'null')
		{
			var dia1 = rsEvento.Fields ("dia1").Value;
			var mes1 = rsEvento.Fields ("mes1").Value;
			var anho1 = rsEvento.Fields ("anho1").Value;
			var hora1 = rsEvento.Fields ("hora1").Value;
			var minutos1 = rsEvento.Fields ("minutos1").Value;
			
			var anho_act = anho1 - 0;
		}
		
		if (DateEnd + "" != 'null')
		{
			var dia2 = rsEvento.Fields ("dia2").Value;
			var mes2 = rsEvento.Fields ("mes2").Value;
			var anho2 = rsEvento.Fields ("anho2").Value;
			var hora2 = rsEvento.Fields ("hora2").Value;
			var minutos2 = rsEvento.Fields ("minutos2").Value;
		}
	}
	else
	{
		var ID = "0";
		var OwnerId = Session("userID");
		var OwnerName = Session("fullName");
		var ToName = "";
		var Subject = "";
		var Content = "";
		var Priority = "0";
		var DateBegin = "";
		var DateEnd = "";
		var EventType = "1";
	}
	    
var EventoOwner = (OwnerId == Session("userID"));
function PropiedadNoOwner(asProperty)
{
	return ( EventoOwner ? "" : asProperty);
}
function TextOff()
{	
	return PropiedadNoOwner ("readonly");
}
function ControlOff()
{	
	return PropiedadNoOwner ("disabled");
}
function Selected(sField)
{	
	if (ID != "0")
		return (oRec.Fields("ID").Value == rsEvento.Fields(sField).Value 
			? "selected" : "");
	else 
		return (oRec.Fields("ID").Value == Session("userID") ? "selected" : "");
}
function _put0(aI)
{
	return ((aI < 10) ? ("0" + aI ) : ("" + aI ));
}
function _selected_option(aExpr)
{
	return (aExpr ? " SELECTED " : "" );
}
	
%>
<html>
<head>
<title><%=Title%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<script language=jscript src ="../js/library.js"></script>
<script language=jscript>

function validafecha(dia, mes, anho)
{
	var ianho = anho * 1;
	var imes = mes * 1;
	var idia = dia * 1;
	
	var diasmes = new Array (0,31,28,31,30,31,30,31,31,30,31,30,31);
	if ((ianho % 4 == 0) && (ianho % 4 != 400))
		diasmes[2] = 29 ;

	return (idia <= diasmes[imes]); 
}

function _save()
{  
	salir = true;
	if (true)
    {   
		if (document.evento.EventType.value + '' == '1')   
		{
			if (validafecha(document.evento.fecha_dia_1.value,
					document.evento.fecha_mes_1.value,
					document.evento.fecha_anho_1.value))
			{
				document.evento.DateBegin.value = 
					document.evento.fecha_dia_1.value  + '/' 
					+ document.evento.fecha_mes_1.value + '/'
					+ document.evento.fecha_anho_1.value + ' '
					+ document.evento.fecha_hora_1.value + ':'
					+ document.evento.fecha_minutos_1.value + ':00';
				var DateBeginFecha = new Date(
						document.evento.fecha_anho_1.value, document.evento.fecha_mes_1.value, document.evento.fecha_dia_1.value, 
						document.evento.fecha_hora_1.value, document.evento.fecha_minutos_1.value) 
			} 
			else
			{
				alert("Fecha de Inicio Incorrecta");
				document.evento.fecha_dia_1.focus();
				salir = false;
			}
			if (validafecha(document.evento.fecha_dia_2.value,
					document.evento.fecha_mes_2.value,
					document.evento.fecha_anho_2.value))
			{
				document.evento.DateEnd.value = 
					document.evento.fecha_dia_2.value  + '/' 
					+ document.evento.fecha_mes_2.value + '/'
					+ document.evento.fecha_anho_2.value + ' '
					+ document.evento.fecha_hora_2.value + ':'
					+ document.evento.fecha_minutos_2.value + ':00';
				var DateEndFecha = new Date(
						document.evento.fecha_anho_2.value, document.evento.fecha_mes_2.value, document.evento.fecha_dia_2.value, 
						document.evento.fecha_hora_2.value, document.evento.fecha_minutos_2.value) 
				if (DateEndFecha < DateBeginFecha )
				{
					alert("La Fecha de Conclusión no puede ser anterior a la Fecha de Inicio");
					document.evento.fecha_dia_2.focus();
					salir = false;
				}
			}
			else
			{
				alert("Fecha de Conclusión Incorrecta");
				document.evento.fecha_dia_2.focus();
				salir = false;
			}
		}
		else
		{
			document.evento.DateBegin.value = 'null';
			document.evento.DateEnd.value = 'null';
		}
	/*	if (document.evento.aPriority.value)
		{
			document.evento.aPriority.value = 1;
		}*/
		if (salir) evento.submit();      
	}
	else
		alert('Titulo no válido, revise la presencia de caráteres especiales');
} 

function Hab_Des_Fechas(aSeleccion)
{
	if (aSeleccion)
	{
		document.getElementById("fecha1").style.display  = "block";
		document.getElementById("fecha2").style.display  = "block";
		document.getElementById("fecha1o").style.display  = "none";
		document.getElementById("fecha2o").style.display  = "none";
	}
	else 
	{
		document.getElementById("fecha1").style.display  = "none";
		document.getElementById("fecha2").style.display  = "none";
		document.getElementById("fecha1o").style.display  = "block";
		document.getElementById("fecha2o").style.display  = "block";
	}

}  

function ChangePriority()
{
	if ((document.evento.Priority.value + "") == '1' )
		document.evento.Priority.value = '0'
	else
		document.evento.Priority.value = '1'
}

//-->
</script>

</head>

<body bgcolor="#FFFFFF" text="#000000" onLoad="Hab_Des_Fechas(<%=EventType%> == 1);">
<FORM ID=evento NAME=evento METHOD=POST ACTION="AgendaEventoSalvar.asp?uid=<%=uid%>" >
<table width="100%" border="0" cellspacing="2" cellpadding="0">
  <tr> 
    <td class="MessageBarTD"> 
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
		<% if (EventoOwner)
			{%>
				<td><a href="javascript:_save()" class="ToolLink"><%=Guardar%></a></td>
				<td>|</td>
		<%	}%>
		  <td><a href="javascript:close()" class="ToolLink"><%=Cerrar%></a></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">

        <tr> 
          <td class="MessageTR1"> 
            <table border="0" cellspacing="0" cellpadding="0" width="100%">
              <tr><table border="0" cellspacing="0" cellpadding="0" width="100%"><tr><td width="1%" > 
                  <%=Prioridad%>
                  <input <%=ControlOff()%> type="checkbox" 
                  ID="aPriority" name="aPriority" 
                  <%if (Priority != "0") Response.Write ("checked");%>
                  onClick = "ChangePriority();">
                </td>
                <td width="1%" class="PriorTD">!</td></tr>
                </table></td>

        <tr> 
          <td class="MessageTR1"> <%=Titulo%>
            <input <%=TextOff()%> type="text" id="Subject" name="Subject" class="Edit" size="50" value="<%=Subject%>">
          </td>
        </tr>
        <tr> 
          <td class="MessageTR1"> 
            <textarea <%=TextOff()%>  name="Content" ID="Content" rows="12" class="TextArea"><%=Content %></textarea>
          </td>
        </tr>
         <tr> 
          <td class="MessageTR1"> 
            <table border="0" cellspacing="0" cellpadding="0" width="100%">
              <tr> <td><%=Para%></td> </tr>
                <tr>
                <td> 
                  <select <%=ControlOff()%> id="Select" name="Select" class="ComboBox">

<%

  var oRec     = Server.CreateObject("ADODB.Recordset");
                 
  //var Sql      = "SELECT * FROM Grupos WHERE (NOT ID IN (" + GUEST_GROUP + ")) ORDER BY Name";
	var UserId = Session ("UserId");
	
   
  //********** Patch ( 30/6/2004 ) **************
  //Se codifican los valores de los option de la forma [ToGroup grupo/ToUsr usuario/ToSubGroup Subgrupo] + indice      
  
  Sql  = "SELECT B.* FROM Usuarios A INNER JOIN Subgrupo B ON B.Tutor = A.ID  " +
			" WHERE (B.Tutor =" + UserId + ") ORDER BY B.Name ";
	//Response.Write (Sql+"AQUI!!!");
	//Response.End;
  oRec = Query( Sql, oRec, oConn  );		
  
  while (oRec.EoF == false) 
  {     
    
%>      
  <OPTION <%=Selected ("ToSubGroup")%> VALUE='ToSubGroup_<%=oRec.Fields("ID").Value%>'>
    SubGrupo : <%=oRec.Fields("Name").Value%>
  </OPTION>
<%  
    oRec.Move(1);
   }
 
  
   
  oRec.Close();


// Esta consulta selecciona todos los grupos con que tiene que ver una persona
// Ya sea porque es miembro del Grupo o le da clases a el grupo 

/*
	var Sql      = "SELECT     Grupos.[ID], Grupos.[Name] " +
					"FROM  Grupos INNER JOIN " +
					"Grupos_de_Usuarios ON Grupos.ID = Grupos_de_Usuarios.[Group] " +
					"WHERE (Grupos_de_Usuarios.[User] = " + UserId + ") AND (Grupos_de_Usuarios.[Group] <> 60)" +
					"UNION " +
					"(SELECT     Grupos_1.[ID], Grupos_1.[Name] " +
					"FROM Grupos INNER JOIN " +
					"Grupos_de_Usuarios ON Grupos.ID = Grupos_de_Usuarios.[Group] INNER JOIN " +
					"Modulo ON Grupos.ID = Modulo.claustro INNER JOIN " +
					"Grupos Grupos_1 ON Modulo.grupo = Grupos_1.ID " +
					"WHERE (Grupos_de_Usuarios.[User] = " + UserId + ") AND (Grupos_de_Usuarios.[Group] <> 60))";
*/

// Consulta de Grupos para La Agenda

// Selecciona los grupos de alumnos que tiene el profesor y 
// los de claustro si es cordinador de la modalidad

	var Sql = "SELECT     Modulo.grupo as [id], Grupos.[Name] " +
				"FROM         Usuarios INNER JOIN " +
                "Grupos_de_Usuarios ON Usuarios.[id] = Grupos_de_Usuarios.[User] INNER JOIN " +
                "Modulo ON Grupos_de_Usuarios.[Group] = Modulo.claustro INNER JOIN " +
                "Grupos ON Modulo.grupo = Grupos.[id] " +
				"WHERE     (Usuarios.ID = " + UserId + ") " +
				"UNION ( " +
				"SELECT     Grupos.[id], Grupos.[Name] " +
				"FROM         Usuarios INNER JOIN " +
                "Grupos_de_Usuarios ON Usuarios.[id] = Grupos_de_Usuarios.[User] INNER JOIN " +
                "Modulo ON Grupos_de_Usuarios.[Group] = Modulo.claustro AND Usuarios.[id] = Modulo.cordinador INNER JOIN " +
                "Grupos ON Modulo.claustro = Grupos.[id] " +
				"WHERE     (Usuarios.[id] = " + UserId + "))";
  
 /*
  Response.Write(Sql+"AQUI!!!");
  Response.End;
 */ 
  
  oRec = Query( Sql, oRec, oConn  );		
  while (oRec.EoF == false) 
   {     
    //se codifican los valores de los option de la forma [ToGroup grupo/ToUsr usuario] + indice   
%>      
  <OPTION <%=Selected ("ToGroup")%> VALUE='ToGroup_<%=oRec.Fields("ID").Value%>'>
    Grupo : <%=oRec.Fields("Name").Value%>
  </OPTION>
<%  
    oRec.Move(1);
   }

  oRec.Close();
%>

<%
  //var Sql      = "SELECT * FROM Usuarios WHERE (NOT ID IN (" + GUEST_USER + ")) ORDER BY FullName";
  /*
  
  var Sql = "SELECT     Usuarios.ID, Usuarios.FullName " +
			"FROM         Grupos INNER JOIN " +
			"Grupos_de_Usuarios ON Grupos.ID = Grupos_de_Usuarios.[Group] INNER JOIN " +
			"Grupos_de_Usuarios Grupos_de_Usuarios_1 ON Grupos.ID = Grupos_de_Usuarios_1.[Group] INNER JOIN " +
			"Usuarios ON Grupos_de_Usuarios_1.[User] = Usuarios.ID " +
			"WHERE     (Grupos_de_Usuarios.[User] = "+ Session("userID") + " )	AND (Grupos_de_Usuarios_1.[Group] <> 60) " +
			"GROUP BY Usuarios.ID, Usuarios.FullName " +
			"UNION( " +
			"SELECT DISTINCT Usuarios_2.ID, Usuarios_2.FullName " +
			"FROM         Grupos_de_Usuarios Grupos_de_Usuarios_1 INNER JOIN " +
			"Usuarios Usuarios_1 ON Grupos_de_Usuarios_1.[User] = Usuarios_1.ID INNER JOIN " +
			"Usuarios Usuarios_2 INNER JOIN " +
			"Grupos_de_Usuarios Grupos_de_Usuarios_2 ON Usuarios_2.ID = Grupos_de_Usuarios_2.[User] ON  " +
			"Grupos_de_Usuarios_1.[Group] = Grupos_de_Usuarios_2.[Group] " +
			"WHERE     (Usuarios_1.ID = "+ Session("userID") + ") AND (Grupos_de_Usuarios_1.[Group] <> 60))";
  */
  
  var Sql = "SELECT     Usuarios.ID, Usuarios.FullName " +
			"FROM         Usuarios " +
			"WHERE     (Usuarios.ID = " + UserId + ")" +
			"UNION (" +
  			"SELECT     Usuarios_1.ID, Usuarios_1.FullName " +
			"FROM         Usuarios INNER JOIN " +
            "Grupos_de_Usuarios ON Usuarios.ID = Grupos_de_Usuarios.[User] INNER JOIN " +
            "Modulo ON Grupos_de_Usuarios.[Group] = Modulo.claustro AND Usuarios.ID = Modulo.cordinador INNER JOIN " +
            "Grupos ON Grupos_de_Usuarios.[Group] = Grupos.ID INNER JOIN " +
            "Grupos_de_Usuarios Grupos_de_Usuarios_1 ON Grupos.ID = Grupos_de_Usuarios_1.[Group] INNER JOIN " +
            "Usuarios Usuarios_1 ON Grupos_de_Usuarios_1.[User] = Usuarios_1.ID " +
			"WHERE     (Usuarios.ID = " + UserId + ") AND (Usuarios_1.ID <> " + UserId + ") " +
			")";

Sql = Sql + " UNION ( " +
			"SELECT     Usuarios_1.ID, Usuarios_1.FullName " +
			"FROM         Modulo INNER JOIN " +
		    "Usuarios INNER JOIN " +
			"Grupos_de_Usuarios ON Usuarios.ID = Grupos_de_Usuarios.[User] ON Modulo.claustro = Grupos_de_Usuarios.[Group] INNER JOIN " +
			"Usuarios Usuarios_1 INNER JOIN " +
			"Grupos_de_Usuarios Grupos_de_Usuarios_1 ON Usuarios_1.ID = Grupos_de_Usuarios_1.[User] ON  " +
			"Modulo.grupo = Grupos_de_Usuarios_1.[Group] " +
            "WHERE     (Usuarios.ID = " + UserId + ") AND (Usuarios_1.ID <> " + UserId + ") " +
			") " +
			"ORDER BY FullName ";

 //Response.Write(Sql);
 //Response.End();

  oRec = Query( Sql, oRec, oConn  );		
  while (oRec.EoF == false) 
   {     
    //se codifican los valores de los option de la forma [ToGroup grupo/ToUsr usuario/ToSubGroup Subgrupo]  + indice   
%>      
  <OPTION <%=Selected ("ToUsr")%> VALUE='ToUsr_<%=oRec.Fields("ID").Value%>'>
    <%=oRec.Fields("FullName").Value%>
  </OPTION>
<%  
    oRec.Move(1);
   }   
   //oRec.Close();
 
  //********** Patch ( 30/6/2004 ) **************
%>  
<%
  DestroyAdoObjects( oConn, oRec );  
%>

                  </select>
                </td>
			   </tr>
            </table>
          </td>
        </tr>       
      </table>
    </td>
  </tr>
  <tr>
	<td class="MessageTR1">
      <select <%=ControlOff()%> id="EventType" name="EventType" onChange="Hab_Des_Fechas(this.value == '1');">
      <% 
      var EventTypeOptions = new Array ("","Con Fecha","Permanente");
      for (var i = 1; i < EventTypeOptions.length  ; i++) { %>
		<option <%=_selected_option(EventType == i)%> value = '<%=i%>'><%=EventTypeOptions[i]%> </option>
      <%}%>
      </select>
	</td>
  </tr>
  <tr id= "fecha1o" height="18" class ="messagetr1"><td>&nbsp;</td></tr>
  <tr id= "fecha2o" height="18" class ="messagetr1"><td>&nbsp;</td></tr>
  <tr id= "fecha1" class ="messagetr1">  
                <td> 
                  <select <%=ControlOff()%> id="fecha_dia_1" name="fecha_dia_1">
                  <% for ( var i = 1; i <= 31; i++) 
                  {%>
                  <option <%=_selected_option(dia1 == _put0(i))%> value = '<%=_put0(i)%>'><%=_put0(i)%> </option>
                  <%
                  }%>
                  </select>
                  <select <%=ControlOff()%> id="fecha_mes_1" name="fecha_mes_1">
                  <%
                  for ( var i = 0; i < 12; i++) 
                  {%>
                  <option <%=_selected_option(mes1 == _put0(i+1))%> value = '<%=_put0(i+1)%>'><%=MONTH_NAME[i]%> </option>
                  <%}%>
                  </select>
				  <select <%=ControlOff()%> id="fecha_anho_1" name="fecha_anho_1">
				  <%
				  
				  for ( var i = (anho_act - 3); (i < (anho_act + 4));(i++)) 
                  {%>
                  <option <%=_selected_option(anho1 == i)%> value = '<%=i%>'><%=i%> </option>
                  <% } %>
                  </select>&nbsp;&nbsp;&nbsp;
                  <select <%=ControlOff()%> id="fecha_hora_1" name="fecha_hora_1">
					<%
					for ( var i = 0; i < 24; i++) 
					{%>
					<option <%=_selected_option(hora1 == _put0(i))%> value = '<%=_put0(i)%>'><%=_put0(i)%> </option>
					<%
					}%>
					</select>
	                <select <%=ControlOff()%> id="fecha_minutos_1" name="fecha_minutos_1">
					<%
					for ( var i = 0; i < 12; i++) 
					{%>
					<option <%=_selected_option(minutos1 == _put0(i*5))%> value = '<%=_put0(i*5)%>'><%=_put0(i*5)%> </option>
					<%
					}%>
					</select>  				
	</td>
              </tr>
                <tr class ="messagetr1" id= "fecha2">  
                <td> 
                  <select <%=ControlOff()%> id="fecha_dia_2" name="fecha_dia_2">
                  <% for ( var i = 1; i <= 31; i++) 
                  {%>
                  <option <%=_selected_option(dia2 == _put0(i))%> value = '<%=_put0(i)%>'><%=_put0(i)%> </option>
                  <%
                  }%>
                  </select>
                  <select <%=ControlOff()%> id="fecha_mes_2" name="fecha_mes_2">
                  <%
                  for ( var i = 0; i < 12; i++) 
                  {%>
                  <option <%=_selected_option(mes2 == _put0(i+1))%> value = '<%=_put0(i+1)%>'><%=MONTH_NAME[i]%> </option>
                  <%}%>
                  </select>
				  <select <%=ControlOff()%> id="fecha_anho_2" name="fecha_anho_2">
				  <%
				  
				  for ( var i = (anho_act - 3); (i < (anho_act + 4));(i++)) 
                  {%>
                  <option <%=_selected_option(anho2 == i)%> value = '<%=i%>'><%=i%> </option>
                  <% } %>
                  </select>&nbsp;&nbsp;&nbsp;
                  <select <%=ControlOff()%> id="fecha_hora_2" name="fecha_hora_2">
					<%
					for ( var i = 0; i < 24; i++) 
					{%>
					<option <%=_selected_option(hora2 == _put0(i))%> value = '<%=_put0(i)%>'><%=_put0(i)%> </option>
					<%
					}%>
					</select>
	                <select <%=ControlOff()%> id="fecha_minutos_2" name="fecha_minutos_2">
					<%
					for ( var i = 0; i < 12; i++) 
					{%>
					<option <%=_selected_option(minutos2 == _put0(i*5))%> value = '<%=_put0(i*5)%>'><%=_put0(i*5)%> </option>
					<%
					}%>
					</select>  				
	</td>
              </tr>
  <tr> 
    <td class="MessageBarTD"> 
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td>&nbsp;&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
  <INPUT TYPE=hidden NAME=ID ID=ID VALUE="<%=ID%>">
  <INPUT TYPE=hidden NAME=OwnerID ID=OwnerID VALUE="<%=OwnerId%>">
  <INPUT TYPE=hidden NAME=OwnerName ID=OwnerName VALUE="<%=OwnerName%>">
  <INPUT TYPE=hidden NAME=ToName ID=ToName VALUE="<%=ToName%>">  
  <INPUT TYPE=hidden NAME=DateBegin ID=DateBegin VALUE="<%=DateBegin%>">
  <INPUT TYPE=hidden NAME=DateEnd ID=DateEnd VALUE="<%=DateEnd%>">
  <INPUT TYPE=hidden NAME=Priority VALUE="<%=Priority%>">
</FORM>
</body>
</html>
<%
 }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>

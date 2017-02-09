<%

// Constantes necesarias para el trabajo con fechas

MONTH_NAME = Array ("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre");
MONTH_NAME_SHORT = Array ("Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic");

DAY_NAME_SHRT = Array ("D","L","M","M","J","V","S");
DAY_NAME_LONG = Array ("Domingo","Lunes","Martes","Miércoles","Jueves","Viernes","Sábado");

// Funciones para construir el Candario de la Agenda

function DateToStringShort (aDate)
{
	var sDateShort = ""; 
	sDateShort = (aDate.getFullYear()+"") + 
	((aDate.getMonth() < (10 - 1)? "0":"") + (aDate.getMonth() + 1)) + 
	((aDate.getDate() < 10 ? "0":"") + aDate.getDate());	
	return sDateShort;
}

function sShortDateToPrettyDate (asDate, year)
{
	var sPrettyDate = ""; 
	sPrettyDate = asDate.substring(6,8)+ "." + MONTH_NAME_SHORT[asDate.substring (4,6)-1] + ".";
	
	if (year == 1)
		sPrettyDate = sPrettyDate + asDate.substring (0,4);
	
	return sPrettyDate;
}

function DateToStringLong (aDate)
{
	var sDateLong = ""; 
	sDateLong = ("" + aDate.getDate())
		 + " de " + MONTH_NAME [aDate.getMonth()] 
		 + " del " + (aDate.getFullYear()+"");	
	return sDateLong;
}

function sDateShortTosDateLong (sDateShort)
{
	var sDateLong = ""; 
	sDateLong = ("" + sDateShort.substr(6, 2)) + " de " +
				("" + MONTH_NAME [sDateShort.substr(4, 2) - 1]) + " de " +
				("" + sDateShort.substr(0, 4));
	return sDateLong;
}

function AgendaBuildCalendario (iaMonth, iaYear, iaDay)
{
    var fHoy = new Date (); 
    iaMonth = ((iaMonth == 0) ? fHoy.getMonth () : (iaMonth - 1));
    iaYear = ((iaYear == 0) ? fHoy.getFullYear () : iaYear);
    iaDay = ((iaDay == 0) ? fHoy.getDate () : iaDay);
    var fAgenda = ((iaMonth == 0) ? fHoy : new Date(iaYear, iaMonth, 1));    

	AgendaBuildMonth (iaMonth, iaYear, iaDay);
}

function HayEventoDia(asUser, asDate)
{
	/*
	var asDate = Session("sAgendaFecha");
	var asUser = Session ("userID");
	*/
	
	dbwhereToGroup = "(AgendaEventos.ToGroup > 0) ";
	dbwhereToSubGroup = "(AgendaEventos.ToSubGroup > 0) ";
	dbwhereToUsr = "(AgendaEventos.ToUsr > 0)";

	dbwhere = " AND (Usuarios.[ID]=" + asUser +
		") AND ('" + asDate + "'" + 
		" BETWEEN CONVERT(char(10), DateBegin, 112) AND " + 
		" CONVERT(char(10), DateEnd, 112))  ";

    sSql = "SELECT CONVERT(char(8), DateBegin, 112) as sDateBegin " + 
				"FROM Grupos_de_Usuarios INNER JOIN " + 
				"Usuarios ON Grupos_de_Usuarios.[User] = Usuarios.ID RIGHT OUTER JOIN " +
				"AgendaEventos ON Grupos_de_Usuarios.[Group] = AgendaEventos.ToGroup " + 
				"WHERE " 
					+ dbwhereToGroup + dbwhere + 
				"UNION ( " + 
				"SELECT CONVERT(char(8), DateBegin, 112) as sDateBegin " + 
				"FROM AgendaEventos INNER JOIN " +
				"Usuarios ON AgendaEventos.ToUsr = Usuarios.ID " +
				"WHERE " 
					+ dbwhereToUsr + dbwhere + 
				") " +
				"UNION ( " + 
				"SELECT CONVERT(char(8), DateBegin, 112) as sDateBegin " + 
				"FROM UsuariosSubGrupo INNER JOIN " +
				"Usuarios ON UsuariosSubGrupo.Usuario = Usuarios.ID INNER JOIN " +
				"AgendaEventos ON UsuariosSubGrupo.Subgrupo = AgendaEventos.ToSubGroup " +
				"WHERE " 
					+ dbwhereToSubGroup + dbwhere + 
				") " ;				
	
	var oConn = Server.CreateObject("ADODB.Connection");  
	var oRec = Server.CreateObject("ADODB.Recordset");  
		
	var filePath = Application("filePath");
	oConn = MakeConnection( oConn, filePath );
	
	oRec = Query(sSql, oRec, oConn);
	return !(oRec.EOF);
}

function BuscaEventosPeriodo(asUser, asDate_i, asDate_f)
{
	
	dbwhereToGroup = "(AgendaEventos.ToGroup > 0) ";
	dbwhereToSubGroup = "(AgendaEventos.ToSubGroup > 0) ";
	dbwhereToUsr = "(AgendaEventos.ToUsr > 0)";
	
	dbwhere = " AND (Usuarios.[ID]=" + asUser +	") AND (CONVERT(char(10), DateBegin, 112) BETWEEN ('" + asDate_i + "') AND ('" +  asDate_f +"'))";

    sSql = "SELECT CONVERT(char(8), DateBegin, 112) as sDateBegin " + 
				"FROM Grupos_de_Usuarios INNER JOIN " + 
				"Usuarios ON Grupos_de_Usuarios.[User] = Usuarios.ID RIGHT OUTER JOIN " +
				"AgendaEventos ON Grupos_de_Usuarios.[Group] = AgendaEventos.ToGroup " + 
				"WHERE " 
					+ dbwhereToGroup + dbwhere + 
				"UNION ( " + 
				"SELECT CONVERT(char(8), DateBegin, 112) as sDateBegin " + 
				"FROM AgendaEventos INNER JOIN " +
				"Usuarios ON AgendaEventos.ToUsr = Usuarios.ID " +
				"WHERE " 
					+ dbwhereToUsr + dbwhere + 
				") " +
				"UNION ( " + 
				"SELECT CONVERT(char(8), DateBegin, 112) as sDateBegin " + 
				"FROM UsuariosSubGrupo INNER JOIN " +
				"Usuarios ON UsuariosSubGrupo.Usuario = Usuarios.ID INNER JOIN " +
				"AgendaEventos ON UsuariosSubGrupo.Subgrupo = AgendaEventos.ToSubGroup " +
				"WHERE " 
					+ dbwhereToSubGroup + dbwhere + 
				") " ;				
	
	var oConn = Server.CreateObject("ADODB.Connection");  
	var oRec = Server.CreateObject("ADODB.Recordset");  
 
	var filePath = Application("filePath");
	oConn = MakeConnection( oConn, filePath );
	
	
	oRec = Query(sSql, oRec, oConn);
	var DiaEvento = new Array;
	for (vi = 1; !oRec.EOF; vi++ ){
		DiaEvento [vi] = oRec.Fields("sDateBegin").Value;
    	oRec.MoveNext();
	}
	return (DiaEvento);
}

function BuscaFecha(aFecha, aDiaEvento, l){
	var aresult = false;
	for (i = 1; (i < l && !aresult); i++){
		if (aDiaEvento [i] == aFecha ){
			aresult = true;
		}
	}
	return (aresult);
}

function AgendaBuildMonth (iaMonth, iaYear, iaDay)
{
    var fInicioMes = new Date(iaYear, iaMonth, 1);
	var iDiaInicioMes = fInicioMes.getDay();
	var iAnnoInicioMes = fInicioMes.getYear();
	var fFechaDia = new Date (fInicioMes);
	var asUser = new String(Session("UserID"));
    %>
    <table>
		<!-- <tr><td class = "MessageTR1">&nbsp;<%=(MONTH_NAME [iaMonth])%> de <%=iAnnoInicioMes%></td></tr> -->
		<tr><td  class = "MessageTR1">
			<table>
    <%    
    %> <tr>  <%
   	for (i = 0; (i < 7 ); i++ )
	{
		%> <td align = "center" class = "MessageTR"> <% Response.Write (DAY_NAME_SHRT[i] + "") %> </td> <%
	}
	%> </tr> <%
	
	var fDate_i = new Date(fFechaDia.getYear() + 0, fFechaDia.getMonth() + 0, 1 );
	var sDate_i = DateToStringShort (fDate_i );
	
	var fDate_f = new Date( fFechaDia.getYear() + 0, fFechaDia.getMonth() + 0,31 );
	var sDate_f = DateToStringShort (fDate_f );
	
	var EventosPeriodo = BuscaEventosPeriodo(asUser, sDate_i, sDate_f);
	l = EventosPeriodo.length;
	var sAgendaFecha = "";

    for (i = 1; (i <= 42); i++)
    {
		if (i-1 % 7 == 0) %> <tr><%    
        %> <td align = "center"  <%
        iMesFecha = fFechaDia.getMonth() + 0;
        if (((i - iDiaInicioMes) > 0) && (iaMonth == iMesFecha))
		{
			fAgendaFecha = fFechaDia;
			sAgendaFecha = DateToStringShort (fAgendaFecha);
			aresult = false

			for (vi = 1; (vi < l && !aresult); vi++){
				if (EventosPeriodo [vi] == sAgendaFecha){
					aresult = true;
				}
			} 
			//Response.Write(EventosPeriodo.length + aresult + "");
			if (aresult ){
				astyle = " Style='font-weight:bolder ' ";
				tdclass = "class =  'MessageTR1' Style='BACKGROUND-COLOR: #FFFF99 '>";
			}
			else
			{
				astyle = "";
				tdclass = "class =  'MessageTR'>";
			}
        
			fFechaDia.setDate (i - iDiaInicioMes + 1);	
			sVinculo = tdclass +'<a '+astyle +' href = "AgendaDia.asp?sAgendaFecha=' + 
				sAgendaFecha + '&uid=' + uid + '" target="mainFrame">';
			sNumeroDia = (i - iDiaInicioMes) + ""
			Response.Write ( sVinculo + sNumeroDia + "</a>" );
		  		 
		} 
		else 
		  Response.Write ("class =  'MessageTR'>&nbsp;");
		%> </td> <%
        if (i % 7 == 0) 
        {
			%> </tr><%
			if (iaMonth != iMesFecha) break ;
        } 
    }
    %>
			</table>
		</td></tr>
    </table>
    <%
}

%>
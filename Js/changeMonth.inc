<%
 function CambiarMes(m) {
 	Meses = new Array();
 	Meses[0] = "Enero";
 	Meses[1] = "Febrero";
 	Meses[2] = "Marzo";
 	Meses[3] = "Abril";
 	Meses[4] = "Mayo";
 	Meses[5] = "Junio";
 	Meses[6] = "Julio";
 	Meses[7] = "Agosto";
 	Meses[8] = "Septiembre";
 	Meses[9] = "Octubre";
 	Meses[10] = "Noviembre";
 	Meses[11] = "Diciembre";
 	
 	cont = 0;
 	while ((m <= 11) && (m != Meses[cont])) {
 		cont++;
 	}
 	cont++;
 	return(cont);
}
%>
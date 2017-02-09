<%
function CogerFecha() {
 var dd, mm, aa;
 
 fech = new Date();
 
 dd = fech.getDate();
 mm = fech.getMonth();
 aa = fech.getFullYear();
 
 fecha = dd + "/" + mm + "/" + aa;
 
 return(fecha);
}
%>
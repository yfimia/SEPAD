<%@ Language=JScript %>
<%    Response.Expires = -1;%>

<!-- #include file='../js/user.inc' -->
<!-- #include file='inc/uestadisticas.inc' -->
<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       
    first = "-1";
    where = "";
    if (Request.QueryString.Count >= 2) {
      first = Request.QueryString.Item("first") + "";
      
      if (first != "-1") {
        //where = " where (Date < " + Application("dchar") +  first + + Application("dchar") +")";
      } 
    } 

%>

<html>
<head>
<title><%= TITLE1 %>...</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/Reports.css" rel="stylesheet" type="text/css" />
</head>

<body >
<%
  var  filePath = Application("dataPath");   
  var  oConn    = Server.CreateObject("ADODB.Connection");
  var  oComm    = Server.CreateObject("ADODB.Command");
  var  oRec     = Server.CreateObject("ADODB.Recordset");
  oRec.CursorLocation = 3;
  oRec.CursorType     = 3;
  oConn.Open( filePath );
  oComm.ActiveConnection = oConn;

  function DateOk(Day,Month,Year)
    {
      var meses =  new Array(31,28,31,30,31,30,31,31,30,31,30,31);
      var Error = new Boolean(false);
      if (((Year % 100 == 0) && ((Year / 100) % 4 == 0)) || ((Year % 4 == 0) && (Year % 100 != 0)))
        meses[1]++; 
      if ((Month <= 12) && (Month > 0))
        {
          if (Day > meses[Month - 1]) 
            Error = true;  
          if ((Day <= 0) || (Year <= 0)) 
            Error = true;
          if (isNaN(parseInt(Year)))
            Error = true;
        } 
       else
         Error = true;
      return Error;   
    }


 if (Request.Form("evaluaciones").Count != 0) 
   {
    /* Response.Write("gr " + Request.Form("grupos") + "<br>");
     Response.Write("us " + Request.Form("usuarios") + "<br>");
     Response.Write("cu " + Request.Form("cursos") + "<br>");
     Response.Write("ev " + Request.Form("evaluaciones") + "<br>");
     Response.Write("d1 " + Request.Form("dia1")+ "<br>");
     Response.Write("m1 " + Request.Form("dia2")+ "<br>");
     Response.Write("y1 " + Request.Form("mes1")+ "<br>");
     Response.Write("d2 " + Request.Form("mes2")+ "<br>");
     Response.Write("m2 " + Request.Form("year1")+ "<br>");
     Response.Write("y2 " + Request.Form("year2")+ "<br>");*/
     
     var Ce = 0;
     var Promedio = 0;
     var Sql1 =   "SELECT  Top " + SHOW_CANT + " Lecciones.Name as LName, Puntuation, Date";
     var Sql2 = " FROM evaluaciones, Lecciones";
     var Sql3 = " WHERE (Lesson = Lecciones.ID) and ([User] = " + Session("userID") + ") and (evaluaciones.Course = " + Session("course") + ")";
     
     var Sql  = Sql1 + Sql2 + Sql3;
     
     if (Request.Form("evaluaciones") != 1) 
       {
         if (Request.Form("evaluaciones") == 2)
           {Sql = Sql + " and (Puntuation = 2)";}
         else
           if (Request.Form("evaluaciones") == 3)
             {Sql = Sql + " and (Puntuation = 3)";}
           else
             if (Request.Form("evaluaciones") == 4)
               {Sql = Sql + " and (Puntuation = 4)";}
             else
               {Sql = Sql + " and (Puntuation = 5)";}
       }
       
     if ((Request.Form("dia1") != "") && (Request.Form("dia2") != "") &&
         (Request.Form("mes1") != "") && (Request.Form("mes2") != "") &&
         (Request.Form("year1")!= "") && (Request.Form("year2")!= ""))
       {
         
         var d1 = Request.Form("dia1");
         var d2 = Request.Form("dia2");
         var m1 = Request.Form("mes1");
         var m2 = Request.Form("mes2")
         var y1 = Request.Form("year1")
         var y2 = Request.Form("year2")
         if ((DateOk(d1,m1,y1) == false) && (DateOk(d2,m2,y2) == false))
           {
             var PFFecha  = Application("dchar")  + m1 + "/" + d1 + "/" + y1 + " 12:00:00 AM" + Application("dchar");
             var PLFecha  = Application("dchar") + m2 + "/" + d2 + "/" + y2 + " 11:59:59 PM" + Application("dchar");
             Sql = Sql + " and ((Date >= " + PFFecha + ") and (Date <= " + PLFecha + "))";
           }
       }  
      
      
     
     oComm.CommandText = Sql + where + " Order By Date";
     //Response.Write ("<font color=green>" + Sql + "</font>");  
             
     oRec = oComm.Execute();
%>
<center>
<H3><%= TITLE1 %></H3>
<H4><%= TEXT16 %><I><U><%=Session("fullName")%></U></I></H4>
<H4><%= TEXT17 %><I><U><%=Session("courseName")%></U></I></H4>
</center>
<br>
<br>
<br>
<table width="100%" border="0" cellspacing="1" cellpadding="1">
  <tr>
    <td class="Main" width="55%" align=center><%= TEXT11 %></td>
    <td class="Main" width="40"  align=center><%= TEXT12 %></td>
    <td class="Main" width="5%" align=center><%= TEXT13 %></td>
 </tr>
<%     
     
     var DAYS            = new Array("Domingo","Lunes","Martes","Miércoles","Jueves","Viernes","Sábado");
     var MONTHS	      = new Array("enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre","octubre", "noviembre", "diciembre");
     var clase = "StatiscticTD1";     

     var last = "-1";          
     
     while (oRec.EOF == false)
       {
          last = oRec.Fields.Item("Date").value + "";          
        
         var fecha = new Date (oRec.Fields.Item("Date").Value );
                 
         fch = DAYS[fecha.getDay()] + " " + 
               fecha.getDate() + ", de " + 
               MONTHS[ fecha.getMonth()] + ", " + 
               fecha.getYear() + " Hora " + 
               fecha.getHours() + ":" + fecha.getMinutes() + ":" + fecha.getSeconds();

%>
  <tr> 
    <td width="55%" class="First"><%=oRec.Fields.Item("LName")%></td>
    <td width="40%" class="First"><%=fch%></td>
    <td width="5%" class="Last" align="center"><%=oRec.Fields.Item("Puntuation")%></td>
  </tr>
<%
         Ce = Ce + 1;
         Promedio = Promedio + oRec.Fields.Item("Puntuation");
         oRec.Move(1);
       }

     if (Ce > 0)
       {Promedio = Promedio / Ce;}
     else
       {Promedio = 0;}  
     var res = new String("" + Promedio + "");
     if (res.indexOf(".") != -1) 
       {
         res = res.substr(0,res.indexOf(".") + 2);
       }       
     
     if (clase == "StatiscticTD1") clase = "StatiscticTD2"; else clase = "StatiscticTD1"         
%>
  <tr> 
    <td colspan=2 class="First"><b><%= TEXT14 %></b></td>
    <td  class="Last" align="center"><b><%=Ce%></b></td>
  </tr>
<!--  <tr> 
    <td colspan=3 class="<%=clase%>"><b>Promedio</b></td>
    <td  class="<%=clase%>" align="center"><b><%=res%></b></td>
  </tr>
-->  
</table>  
<%       
   }
    
  
  
  
  //Cierro la coneccion con la base de datos
  oConn.Close();   
  //Response.Write("" + Request.Form("grupos") + "");
  
%>   
<!--table border="0" cellspacing="0" cellpadding="0" class="ToolBar" width="100%">
  <tr>
    <td>
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td nowrap><a href="javascript:Notas.submit()" class="ToolLink">&nbsp;Consultar&nbsp;</a></td>
          <td>|</td>
          <td><a href="Festadisticas.asp?uid=<%=uid%>&first=<%=last%>" class="ToolLink">&nbsp;Próximos&nbsp;<%=SHOW_CANT%>&nbsp;</a></td>
        </tr>
      </table>
    </td>
  </tr>
</table-->

 </body>
</HTML>
<%  }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
  
%>
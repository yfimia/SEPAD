<%@ Language=JScript %>
<%    Response.Expires = -1;%>

<!-- #include file='../js/user.inc' -->
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
<title>Estad&iacute;sticas acad&eacute;micas...</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/EvalResult.css" type="text/css">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000" class="MessageTR1">
<%//  Response.Write("<div align=center><h1><font color=white>Consulte y descubrirá cosas inimaginables,  por favor, espere un momento y no se arrepentirá</font></h1></div><br><br>");    
  //Abro la coneccion con la base de datos   
  var  filePath = Application("dataPath");   
  var  oConn    = Server.CreateObject("ADODB.Connection");
  var  oComm    = Server.CreateObject("ADODB.Command");
  var  oRec     = Server.CreateObject("ADODB.Recordset");
  oRec.CursorLocation = 3;
  oRec.CursorType     = 3;
  oConn.Open( filePath );
  oComm.ActiveConnection = oConn;
  
  function Cursos()
    {
      oComm.CommandText = "Select Name from Cursos";      
      oRec = oComm.Execute();
      while (oRec.EOF == false)
        {
          Response.Write ("<option>" + oRec.Fields.Item ("Name").Value + "</option>");
          oRec.Move(1);
        }
    }

  function Grupos()
    {
      oComm.CommandText = "Select Name from Grupos";      
      oRec = oComm.Execute();
      while (oRec.EOF == false)
        {
          Response.Write ("<option>" + oRec.Fields.Item ("Name").Value + "</option>");
          oRec.Move(1);
        }
    }
  
  function Usuarios()
    {
      oComm.CommandText = "Select FullName from Usuarios";      
      oRec = oComm.Execute();
      while (oRec.EOF == false)
        {
          Response.Write ("<option>" + oRec.Fields.Item ("FullName").Value + "</option>");
          oRec.Move(1);
        }
    }
  
%>
<table width="100%" border="0" cellspacing="0" cellpadding="2">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td><img src="../images/<%=Session("skin")%>/ReportesIMG.gif" width="80" height="54"></td>
          <td class="HeaderTable">Estad&iacute;sticas acad&eacute;micas</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
 <form name=Notas Id=Notas action="Festadisticas.asp?uid=<%=uid%>" method=post>
     <table width="100%" border="0" cellspacing="0" cellpadding="0" class="StatisticTable">
       <tr>
         <td  colspan=3 align=center>
             <b class="StCTD">Formulario de consultas</b>
         </td>
       </tr>
       <tr>
         <td>
           <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
             <tr>
               <td colspan=3 align=center>
                   Desde
               </td>
             </tr>
             <tr>
               <td align=center>
                   Día:
               </td>
               <td align=center>
                   Mes:
               </td>
               <td align=center>
                   Año:
               </td>
             </tr> 
             <tr>
               <td align=center class="MarqueeTD">
                  <select name=dia1 id=dia1 >
                    <option value="">-</option>
                    <option value="1">1</option>
                    <option value="2">2</option>
                    <option value="3">3</option>
                    <option value="4">4</option>
                    <option value="5">5</option>
                    <option value="6">6</option>
                    <option value="7">7</option>
                    <option value="8">8</option>
                    <option value="9">9</option>
                    <option value="10">10</option>
                    <option value="11">11</option>
                    <option value="12">12</option>
                    <option value="13">13</option>
                    <option value="14">14</option>
                    <option value="15">15</option>
                    <option value="16">16</option>
                    <option value="17">17</option>
                    <option value="18">18</option>
                    <option value="19">19</option>
                    <option value="20">20</option>
                    <option value="21">21</option>
                    <option value="22">22</option>
                    <option value="23">23</option>
                    <option value="24">24</option>
                    <option value="25">25</option>
                    <option value="26">26</option>
                    <option value="27">27</option>
                    <option value="28">28</option>
                    <option value="29">29</option>
                    <option value="30">30</option>
                    <option value="31">31</option>
                  </select>
                 
               </td>
               <td align=center class="MarqueeTD">
                  <select name=mes1 id=mes1 >
                    <option value="">-</option>
                    <option value="1">Enero</option>
                    <option value="2">Febrero</option>
                    <option value="3">Marzo</option>
                    <option value="4">Abril</option>
                    <option value="5">Mayo</option>
                    <option value="6">Junio</option>
                    <option value="7">Julio</option>
                    <option value="8">Agosto</option>
                    <option value="9">Septiembre</option>
                    <option value="10">Octubre</option>
                    <option value="11">Noviembre</option>
                    <option value="12">Diciembre</option>
                  </select>
                 
               </td>
               <td align=center class="MarqueeTD">
                 <input name=year1 id=year1 type=text size=4 maxlength=4>
               </td>
             </tr>
           </table>
         </td>  
         <td>
           <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
             <tr>
               <td colspan=3 align=center>
                   Hasta
               </td>
             </tr>
             <tr>
               <td align=center>

                   Día:

               </td>
               <td align=center>

                   Mes:

               </td>
               <td align=center>

                   Año:

               </td>
             </tr> 
             <tr>
               <td align=center class="MarqueeTD">
                  <select name=dia2 id=dia2>
                    <option value="">-</option>
                    <option value="1">1</option>
                    <option value="2">2</option>
                    <option value="3">3</option>
                    <option value="4">4</option>
                    <option value="5">5</option>
                    <option value="6">6</option>
                    <option value="7">7</option>
                    <option value="8">8</option>
                    <option value="9">9</option>
                    <option value="10">10</option>
                    <option value="11">11</option>
                    <option value="12">12</option>
                    <option value="13">13</option>
                    <option value="14">14</option>
                    <option value="15">15</option>
                    <option value="16">16</option>
                    <option value="17">17</option>
                    <option value="18">18</option>
                    <option value="19">19</option>
                    <option value="20">20</option>
                    <option value="21">21</option>
                    <option value="22">22</option>
                    <option value="23">23</option>
                    <option value="24">24</option>
                    <option value="25">25</option>
                    <option value="26">26</option>
                    <option value="27">27</option>
                    <option value="28">28</option>
                    <option value="29">29</option>
                    <option value="30">30</option>
                    <option value="31">31</option>
                  </select>
                 
               </td>
               <td align=center class="MarqueeTD">
                  <select name=mes2 id=mes2 >
                    <option value="">-</option>
                    <option value="1">Enero</option>
                    <option value="2">Febrero</option>
                    <option value="3">Marzo</option>
                    <option value="4">Abril</option>
                    <option value="5">Mayo</option>
                    <option value="6">Junio</option>
                    <option value="7">Julio</option>
                    <option value="8">Agosto</option>
                    <option value="9">Septiembre</option>
                    <option value="10">Octubre</option>
                    <option value="11">Noviembre</option>
                    <option value="12">Diciembre</option>
                  </select>
                 
               </td>
               <td align=center class="MarqueeTD">
                 <input name=year2 id=year2 type=text size=4 maxlength=4>
               </td>
             </tr>
           </table>
         </td>         
       </tr>

       <tr>
         <td colspan=2>
           <table border="0" cellspacing="0" cellpadding="0"  class="ToolBar">
             <tr>
               <td>

                   Grupo:

               </td>
               <td >
                 <select  id=grupos name=grupos LANGUAGE=javascript >
                   <option selected>
                     Todos los grupos
                   </option>
                   <%=Grupos()%>
                 </select>
               </td>
             </tr>
             <tr>
               <td>

                   Usuario:

               </td>
               <td >
                 <select  id=usuarios name=usuarios LANGUAGE=javascript  >
                   <option selected>
                     Todos los usuarios
                   </option>
                   <%=Usuarios()%>
                 </select>
               </td>
             </tr>
             <tr>
               <td>

                   Curso:

               </td>
               <td >
                 <select id=cursos name=cursos >
                   <option selected>
                     Todos los cursos
                   </option>
                   <%=Cursos()%>
                 </select>
               </td>
             </tr>
             <tr>
               <td>

                   Calificación:

               </td>
               <td >
                 <select id=evaluaciones name=evaluaciones>
                   <option selected>
                     Todas las Calificaciones
                   </option>
                   <option>
                     Calificaciones de 2 puntos
                   </option>
                   <option>
                     Calificaciones de 3 puntos
                   </option>
                   <option>
                     Calificaciones de 4 puntos
                   </option>
                   <option>
                     Calificaciones de 5 puntos
                   </option>
                 </select>
               </td>
             </tr>
           </table>
         </td>
       
       </tr>
       <tr>
         <td colspan=3 align=center>
           <input type=submit value="Consultar">
         <td>
       </tr>
     </table>
   </form>  
<%

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
        } 
       else
         Error = true;
      return Error;   
    }


 if (Request.Form("grupos").Count != 0) 
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
     var Sql1 =   "SELECT  Top " + SHOW_CANT + "  C.fullname, B.Name AS Curso, A.Puntuation, A.Date";
     var Sql2 = " FROM evaluaciones AS A, Cursos AS B, Usuarios AS C";
     var Sql3 = " WHERE (A.[User] = C.ID) and (A.Course = B.ID)";
     
     if (Request.Form("grupos") != "Todos los grupos") 
       {
         Sql1 = Sql1 + ",D.Name AS Grupo";
         Sql2 = Sql2 + ", Grupos AS D, Grupos_de_Usuarios AS E";
         Sql3 = Sql3 + "and (E.[Group] = D.ID) and  (A.[User] = E.[User]) and (D.Name = '" + Request.Form("grupos") + "')";
       }
      
     var Sql  = Sql1 + Sql2 + Sql3;
     
     if (Request.Form("usuarios") != "Todos los usuarios") 
       {
         Sql = Sql + " and (C.fullname = '" + Request.Form("usuarios") + "')";
       } 
     if (Request.Form("cursos") != "Todos los cursos") 
       {
         Sql = Sql + " and (B.Name = '" + Request.Form("cursos") + "')";
       }
     if (Request.Form("evaluaciones") != "Todas las Calificaciones") 
       {
         if (Request.Form("evaluaciones") == "Calificaciones de 2 puntos")
           {Sql = Sql + " and (A.Puntuation = 2)";}
         else
           if (Request.Form("evaluaciones") == "Calificaciones de 3 puntos")
             {Sql = Sql + " and (A.Puntuation = 3)";}
           else
             if (Request.Form("evaluaciones") == "Calificaciones de 4 puntos")
               {Sql = Sql + " and (A.Puntuation = 4)";}
             else
               {Sql = Sql + " and (A.Puntuation = 5)";}
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
             Sql = Sql + " and ((A.Date >= " + PFFecha + ") and (A.Date <= " + PLFecha + "))";
           }
       }  
      
      
     
     oComm.CommandText = Sql + where + " Order By Date";
     //Response.Write ("<font color=green>" + Sql + "</font>");  
             
     oRec = oComm.Execute();
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" class="StatisticTable">
  <tr>
    <td width="30%" align=center>Nombre</td>
    <td width="30%" align=center>Curso</td>
    <td width="35%"  align=center>Fecha</td>
    <td width="5%" align=center>Puntuación</td>
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

          if (clase == "StatiscticTD1") clase = "StatiscticTD2"; else clase = "StatiscticTD1"        
%>
  <tr> 
    <td width="30%" class="<%=clase%>"><%=oRec.Fields.Item("fullname")%></td>
    <td width="30%" class="<%=clase%>" align="center"><%=oRec.Fields.Item("Curso")%></td>
    <td width="35%" class="<%=clase%>"><%=fch%></td>
    <td width="5%" class="<%=clase%>" align="center"><%=oRec.Fields.Item("Puntuation")%></td>
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
    <td colspan=3 class="<%=clase%>"><b>Total</b></td>
    <td  class="<%=clase%>" align="center"><b><%=Ce%></b></td>
  </tr>
<%
    if (clase == "StatiscticTD1") clase = "StatiscticTD2"; else clase = "StatiscticTD1"        
%>  
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
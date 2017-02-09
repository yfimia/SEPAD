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
<link rel="stylesheet" href="../css/<%=Session("skin")%>/EvalResult.css" type="text/css">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<script>
 function goPrint()
   {
     Notas.target = "_blank";
     Notas.action = "printableEstadisticas.asp?uid=<%=uid%>";
     Notas.submit();
   }
</script>

</head>

<body >
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

  
    
%>
<table width="100%" border="0" cellspacing="0" cellpadding="2">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td><img src="../images/<%=Session("skin")%>/ReportesIMG.gif" width="80" height="54"></td>
          <td class="HeaderTable"><%= TITLE1 %></td>
        </tr>
      </table>
    </td>
    <td class="ToolBar" width="1%" valign="bottom">
      <a href="javascript:goPrint()"><img height=20 width=20 src="../images/<%=Session("skin")%>/imprimir.gif" title="<%=TEXT15%>" border=0></a>
    </td>
  </tr>
</table>
 <form name=Notas Id=Notas action="uestadisticas.asp?uid=<%=uid%>" method=post>
     <table width="100%" border="0" cellspacing="0" cellpadding="0" class="StatisticTable">
       <tr>
         <td  colspan=3 align=center>
             <b class="StCTD"><%= TITLE2 %></b>
         </td>
       </tr>
       <tr>
         <td>
           <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
             <tr>
               <td colspan=3 align=center><%= TEXT1 %></td>
             </tr>
             <tr>
               <td align=center><%= TEXT2 %>:</td>
               <td align=center><%= TEXT3 %>:</td>
               <td align=center><%= TEXT4 %>:</td>
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
                    <option value="1"><%= TEXT19 %></option>
                    <option value="2"><%= TEXT20 %></option>
                    <option value="3"><%= TEXT21 %></option>
                    <option value="4"><%= TEXT22 %></option>
                    <option value="5"><%= TEXT23 %></option>
                    <option value="6"><%= TEXT24 %></option>
                    <option value="7"><%= TEXT25 %></option>
                    <option value="8"><%= TEXT26 %></option>
                    <option value="9"><%= TEXT27 %></option>
                    <option value="10"><%= TEXT28 %></option>
                    <option value="11"><%= TEXT29 %></option>
                    <option value="12"><%= TEXT30 %></option>
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
               <td colspan=3 align=center><%= TEXT5 %></td>
             </tr>
             <tr>
               <td align=center><%= TEXT2 %>:</td>
               <td align=center><%= TEXT3 %>:</td>
               <td align=center><%= TEXT4 %>:</td>
               <td align=center><%= TEXT6 %>:</td>
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
                    <option value="1"><%= TEXT19 %></option>
                    <option value="2"><%= TEXT20 %></option>
                    <option value="3"><%= TEXT21 %></option>
                    <option value="4"><%= TEXT22 %></option>
                    <option value="5"><%= TEXT23 %></option>
                    <option value="6"><%= TEXT24 %></option>
                    <option value="7"><%= TEXT25 %></option>
                    <option value="8"><%= TEXT26 %></option>
                    <option value="9"><%= TEXT27 %></option>
                    <option value="10"><%= TEXT28 %></option>
                    <option value="11"><%= TEXT29 %></option>
                    <option value="12"><%= TEXT30 %></option>
                  </select>
                 
               </td>
               <td align=center class="MarqueeTD">
                 <input name=year2 id=year2 type=text size=4 maxlength=4>
               </td>

               <td >
                 <select id=evaluaciones name=evaluaciones>
                   <option selected value="1"><%= TEXT9 %></option>
                   <option value="2"><%= TEXT7 %> 2 <%= TEXT8 %></option>
                   <option value="3"><%= TEXT7 %> 3 <%= TEXT8 %></option>
                   <option value="4"><%= TEXT7 %> 4 <%= TEXT8 %></option>
                   <option value="5"><%= TEXT7 %> 5 <%= TEXT8 %></option>
                 </select>
               </td>
               
             </tr>
           </table>
         </td>         
       </tr>

       <tr>
       
       </tr>
       <tr>
         <td colspan=3 align=center>
           <input type=submit value="<%= TEXT10 %>">
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
     var Sql1 =   "SELECT  Top " + SHOW_CANT + " evaluaciones.Id, evaluaciones.exercise, ejercicios.[key], Lecciones.Name as LName, Puntuation, Date";
     var Sql2 = " FROM evaluaciones, Lecciones, Ejercicios ";
     var Sql3 = " WHERE (evaluaciones.Lesson = Lecciones.ID) and (Ejercicios.id = evaluaciones.exercise) and ([User] = " + Session("userID") + ") and (evaluaciones.Course = " + Session("course") + ")";
     
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
<table width="100%" border="0" cellspacing="1" cellpadding="1" class="StatisticTable">
  <tr>
    <td width="25%" align=center>Ejercicio</td>
    <td width="40%" align=center><%= TEXT11 %></td>
    <td width="30"  align=center><%= TEXT12 %></td>
    <td width="5%" align=center><%= TEXT13 %></td>
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
    <td width="25%" class="<%=clase%>"><a href="preview.asp?uid=<%=uid%>&id=<%=oRec.Fields.Item("id")%>"><%=oRec.Fields.Item("key")%></a></td>
    <td width="40%" class="<%=clase%>"><%=oRec.Fields.Item("LName")%></td>
    <td width="30%" class="<%=clase%>"><%=fch%></td>
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
    <td colspan=3 class="<%=clase%>"><b><%= TEXT14 %></b></td>
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
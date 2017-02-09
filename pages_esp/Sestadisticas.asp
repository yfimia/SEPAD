<%@ Language=JScript %>

<%Response.Expires = -1;%>
<!-- #include file='../js/user.inc' -->
<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

%>



<html>
<head>
<title>Estad&iacute;sticas del Sistema...</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/EvalResult.css" type="text/css">
<SCRIPT src="../Js/user.js" language="JSCRIPT"></SCRIPT>
</head>

<body bgcolor="#FFFFFF" text="#000000" class="MessageTR1">
<table width="100%" border="0" cellspacing="0" cellpadding="2">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td><img src="../images/<%=Session("skin")%>/ReportesIMG.gif" width="80" height="54"></td>
          <td class="HeaderTable">Datos Generales del Sistema</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<%
  var  filePath = Application("dataPath");   
  var  oConn    = Server.CreateObject("ADODB.Connection");
  var  oComm    = Server.CreateObject("ADODB.Command");
  var  oRec     = Server.CreateObject("ADODB.Recordset");
  var Cnames = new Array(); 
  var CMark = new Array(); 
  oRec.CursorLocation = 3;
  oRec.CursorType     = 3;

  
  oConn.Open(filePath);
  oComm.ActiveConnection = oConn;
  
  oComm.CommandText = "SELECT Count(*) FROM Usuarios";
  oRec = oComm.Execute();
  var cu = oRec.Fields.Item(0).Value;
  oRec.Close();
    
  oComm.CommandText = "SELECT Name FROM Modulo";
  oRec = oComm.Execute();
  
  var cc = 0;
  while (oRec.EOF == false)
    {
      cc++;
      Cnames[cc] = oRec.Fields.Item("Name").Value;
      CMark[cc] = 0;
      oRec.Move(1); 
    }

  oComm.CommandText = "SELECT Sum(Logins) as Cantv FROM Usuarios";
  oRec = oComm.Execute();
  var cv = oRec.Fields.Item("Cantv").Value;
  
%>
<table border="0" cellspacing="3" cellpadding="0" class="BorderedTable">
  <tr> 
    <td width="80%"><b>Cantidad &nbsp;de &nbsp;modalidades:</b></td>
    <td width="20%"><%=cc%></td>
  </tr>
  <tr> 
    <td width="80%"><b>Cantidad&nbsp;de&nbsp;usuarios:</b></td>
    <td width="20%"><%=cu%></td>
  </tr>
  <tr>
    <td width="80%"><b>Cantidad&nbsp;de&nbsp;visitas:</b></td>
    <td width="20%"><%=cv%></td>
  </tr>
</table>

<table width="100%" border="0" cellspacing="0" cellpadding="0" class="StatisticTable">
  <tr> 
    <td width="99%" align="center"><b class="StCTD">Modalidades Académicas</b></td>
    <td width="1%"> 
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr valign="middle" align="center"> 
          <td colspan="2"><b>Cantidad de Visitas</b></td>
        </tr>
        <tr align="center"> 
          <td><b>DESDE</b></td>
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <td> 
            <table width="100" border="0" cellspacing="0" cellpadding="1" class="BorderedTable">
              <tr align="center"> 
                <td><b>D&iacute;a</b></td>
                <td><b>Mes</b></td>
                <td><b>A&ntilde;o</b></td>
              </tr>
              <tr valign="middle"> 
                <td class="MarqueeTD"> 
                  <select NAME=O1>
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
                <td class="MarqueeTD"> 
                  <select NAME=O2>
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
                <td class="MarqueeTD"> 
                  <input type="text" NAME=O3 size="4" maxlength="4">
                </td>
              </tr>
            </table>
          </td>
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <td align="center"><b>HASTA</b></td>
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <td align="center"> 
            <table width="100" border="0" cellspacing="0" cellpadding="1" class="BorderedTable">
              <tr align="center"> 
                <td><b>D&iacute;a</b></td>
                <td><b>Mes</b></td>
                <td><b>A&ntilde;o</b></td>
              </tr>
              <tr valign="middle"> 
                <td class="MarqueeTD"> 
                  <select NAME=O4>
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
                <td class="MarqueeTD"> 
                  <select NAME=O5>
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
                <td class="MarqueeTD"> 
                  <input type="text" NAME=O6 size="4" maxlength="4">
                </td>
              </tr>
            </table>
          </td>
          <td>&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>

<%  
  var total = 0;

  function marca(nombre)
   {
     var j = 1;
     while ((j <= cc) && (Cnames[j] != nombre))
      { j++; }
     if  (j <= cc) 
      { CMark[j] = 1; }
   }
  
  var clase = "StatiscticTD2";
  
  if (Request.QueryString.Count > 1) 
    {
      var PFDay   = Request.QueryString.Item(2) + "";  
      var PFMonth = Request.QueryString.Item(3) + "";  
      var PFYear  = Request.QueryString.Item(4) + "";  
      var PLDay   = Request.QueryString.Item(5) + "";  
      var PLMonth = Request.QueryString.Item(6) + "";  
      var PLYear  = Request.QueryString.Item(7) + "";  
      var PFFecha  = Application("dchar") + PFMonth + "/" + PFDay + "/" + PFYear + " 12:00:00 AM" +  Application("dchar");
      var PLFecha  = Application("dchar") + PLMonth + "/" + PLDay + "/" + PLYear + " 11:59:59 PM" + Application("dchar");

      var Sql = "SELECT A.Name,COUNT(*) as Cantidad FROM Cursos A, Conexiones_a_Cursos B WHERE (A.ID = B.Course)";      
      
      if ((PFDay!="") && (PFMonth!="") && (PFYear!="") && (PLDay!="") && (PLMonth!="") && (PLYear!="")) 
        Sql = Sql + " AND (B.Date >= " + PFFecha + ") AND (B.Date <= " + PLFecha + ")";
      Sql = Sql + " GROUP BY A.Name";
      oComm.CommandText = Sql;
      //oComm.CommandText = "SELECT A.Name,Sum(B.Course - B.Course + 1) as Cantidad FROM Cursos A, Conexiones_a_Cursos B WHERE (A.ID = B.Course) GROUP BY A.Name";
      oRec = oComm.Execute();
      
      while (oRec.EOF == false)
        {
          
          if (clase == "StatiscticTD1") clase = "StatiscticTD2"; else clase = "StatiscticTD1"
          total = total + oRec.Fields.Item("Cantidad");
%>
  <tr> 
    <td width="60%" class="<%=clase%>"><%=oRec.Fields.Item("Name")%></td>
    <td width="40%" class="<%=clase%>" align="center"><%=oRec.Fields.Item("Cantidad")%></td>
  </tr>
<%        
          marca(oRec.Fields.Item("Name"));
          oRec.Move(1); 
        }
    }    
   else
    {
      oComm.CommandText = "SELECT A.Name,Sum(B.Course - B.Course + 1) as Cantidad FROM Cursos A, Conexiones_a_Cursos B WHERE (A.ID = B.Course) GROUP BY A.Name";
      oRec = oComm.Execute();
      while (oRec.EOF == false)
        {
          if (clase == "StatiscticTD1") clase = "StatiscticTD2"; else clase = "StatiscticTD1"        
          total = total + oRec.Fields.Item("Cantidad");

%>
  <tr> 
    <td width="60%" class="<%=clase%>"><%=oRec.Fields.Item("Name")%></td>
    <td width="40%" class="<%=clase%>" align="center"><%=oRec.Fields.Item("Cantidad")%></td>
  </tr>
<%        
          marca(oRec.Fields.Item("Name"));
          oRec.Move(1); 
        }    
    } 
  
  
  for (i = 1;i <= cc;i++)
   {
     //Response.Write(Cnames[i] + " esta en " + CMark[i]);
     if (CMark[i] == 0)
      {
        if (clase == "StatiscticTD1") clase = "StatiscticTD2"; else clase = "StatiscticTD1"      
%>
  <tr> 
    <td width="60%" class="<%=clase%>"><%=Cnames[i]%></td>
    <td width="40%" class="<%=clase%>" align="center">0</td>
  </tr>

<%        
      }  

   }
  oConn.Close();
%>
  <tr> 
    <td width="60%" class="<%=clase%>"><b>Total</b></td>
    <td width="40%" class="<%=clase%>" align="center"><b><%=total%></b></td>
  </tr>

  <tr>
    <td colspan=2 class="ToolBar" align="right"> 
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td><a href="javascript:GetData()" class="ToolLink">&nbsp;Consultar&nbsp;</a></td>
          <td>|</td>
          <td><a href="javascript:close()" class="ToolLink" >&nbsp;Cerrar&nbsp;</a></td>
        </tr>
      </table>
    </td>
  </tr>
</table>


<script language=jscript>
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
   
   function Valid()
     {
       var Error = new Boolean(false);
       if (O1.value + O2.value + O3.value + O4.value + O5.value + O6.value + "" == "")  
         Error = false
       else
         if ((O1.value == "") || (O2.value == "")  || (O3.value == "")  || (O4.value == "")  || (O5.value == "")  || (O6.value == ""))
           Error = true
         else  
           {
             if (DateOk(parseInt(O1.value + ""),parseInt(O2.value + ""),parseInt(O3.value + "")) == true)
               Error = true;
             if (DateOk(parseInt(O4.value + ""),parseInt(O5.value + ""),parseInt(O6.value + "")) == true)  
               Error = true;
           }  
       return Error;  
     }
   
   function GetData()
     { 
       if (Valid() == false) 
         {   
           var res = "";      
           res = res + "&FDay="    + O1.value + "";
           res = res + "&FMonth="  + O2.value + "";
           res = res + "&FYear="   + O3.value + "";
           res = res + "&LDay="    + O4.value + "";
           res = res + "&LMonth="  + O5.value + "";
           res = res + "&LYear="   + O6.value + "";
           window.location = "SEstadisticas.asp?uid=<%=uid%>" + res;
         } 
       else
         alert(INVALID_DATE_TEXT);
     }


</script>

</body>
</html>


<%
  }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>



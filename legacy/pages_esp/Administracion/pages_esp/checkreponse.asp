<%@ Language=JavaScript %>
<%
 Response.Expires = -1;
%>
<!-- #include file='../js/adolibrary.inc' -->
<!-- #include file='../js/user.inc' -->
<!-- #include file='inc/checkresponse.inc' -->

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

%>


<SCRIPT language=vbscript RUNAT=Server>
 function getAtualTime
  getAtualTime = Now()
 End function 

</SCRIPT>

<%

// Revision de ejercicios...

 var i,j,k,ok,pe, bien;
 var cad,req;
 var rArr = new Array();
 var pArr = new Array();
 var rArrCant;
 var eval = new Array();
 var response,pesos;
 var posi;
 var cteOpt = 10; // constante de evaluacion optima...
 var resOpt;

function getResp(resp,pes,po)
  {
   var i,p,p1;
   var tmpOpt = 0;
   
   resp = resp + ";";
   pes = pes + ";";
   
   for (i = 0; i < po; i++)
   {
    p = resp.indexOf(";",0);
    p1 = pes.indexOf(";",0);
    resp = resp.substring(p + 1,resp.length);
    pes = pes.substring(p1 + 1,pes.length);
   }
   
   p = resp.indexOf(";",0);
   p1 = pes.indexOf(";",0);
   
   if (p != -1)
    {resp = resp.substring(0,p);}
   if (p1 != -1)
    {pes = pes.substring(0,p1);}
     
   resp = resp + ",";
   pes = pes + ",";
   
   i = 0;
   while (resp != "")
   {
    p = resp.indexOf(",");
    rArr[i] = resp.substring(0,p);
    resp = resp.slice(p + 1,resp.length);
    if (posi != -1)
     {
      p1 = pes.indexOf(",");
      pArr[i] = pes.substring(0,p1);
      pes = pes.slice(p1 + 1,pes.length);
     }
    else
     {
      pArr[i] = cteOpt + "";
     }  
     if (parseInt(pArr[i],10) > tmpOpt)
       {tmpOpt = pArr[i];} 

    i++;
   }
   rArrCant = i; 
   resOpt = resOpt + parseInt(tmpOpt,10);    
  
  } 
   
 function quitEsp( cad )
  {
    while (cad.indexOf(' ',1) > 0)
      {
       cad = cad.replace(' ','');
       
      }
      
    return cad;  
  }   
  
 function isokrArr(cad)
  { var i;
    var result = false;
    var tmp1,tmp2;
   
   for (i = 0;i < rArrCant;i++)
   {
    
    tmp1 = quitEsp(cad.toUpperCase());
    tmp2 = quitEsp(rArr[i].toUpperCase());
    
    

    if (tmp1 == tmp2)
     { 
      result = true;
      pe = pArr[i];
      return result;
      break;
     }
   }
    
  }
 
 function getPercent(p,total)
  { 
    var result = Math.round((p / total) * 5);
    return result;
  }
 
  
 
 
function insertEval(ev)
 { 
  var cad; 
    
  var  filePath = Application("dataPath"); 
  var  oConn    = Server.CreateObject("ADODB.Connection");
  var  oRec     = Server.CreateObject("ADODB.Recordset");
  var  oRec1    = Server.CreateObject("ADODB.Recordset");
       
       oRec.CursorLocation  = 3;
       oRec.CursorType      = 3;
       oRec1.CursorLocation = 3;
       oRec1.CursorType     = 3;

  
      oConn.Errors.Clear();
      oConn.Open( filePath );
      
          oRec.Open("select * from Evaluaciones",oConn,3,3);
          oRec1.Open("select * from Evaluaciones_Pendientes",oConn,3,3);
          
          for (i = 0;i <= Session("Cantidad");i++ )
           { //Response.Write(ev[i] + "<br>");
            if (ev[i] != 0) 
             {
              oRec.AddNew();
              oRec.Fields.Item("User").value      = Session("userID");
              oRec.Fields.Item("Course").value    = Session("course");           
              oRec.Fields.Item("Lesson").value    = Session("lastLesson");
              
              oRec.Fields.Item("Exercise").value    = Session("exer").items.id[Session("Cuales").items[i]];
              oRec.Fields.Item("Puntuation").value    = ev[i];
              oRec.Fields.Item("Date").value    = getAtualTime();
              oRec.Update();      
             }
            else
             if (Session("userID") != GUEST_USER)
             { 
              cad = "O" + Session("exer").items.id[Session("Cuales").items[i]] + "#" + 0;
              oRec1.AddNew();
              oRec1.Fields.Item("User").value      = Session("userID");
              oRec1.Fields.Item("Course").value    = Session("course");   
              oRec1.Fields.Item("Lesson").value    = Session("exer").items.lesson[Session("Cuales").items[i]];
              oRec1.Fields.Item("Url").value    = Session("exer").items.file[Session("Cuales").items[i]];
              oRec1.Fields.Item("Exercise").value    = Session("exer").items.id[Session("Cuales").items[i]];
              oRec1.Fields.Item("Response").value    = Request.Form.Item(cad);
//Response.Write(Request.Form.Item(cad) + "<br>");
//Response.Write(cad + "<br>");
  
              oRec1.Fields.Item("Date").value    = getAtualTime();
              oRec1.Update();      
             }  
           } 
              
      oConn.Close();     
  }
%>  

<html>
<head>
<title><%=title%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/EvalResult.css" type="text/css">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000">

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td><img src="../images/<%=Session("skin")%>/EvalResult.gif" width="80" height="54"></td>
          <td class="HeaderTable"><%=stmp1%></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td> 
      <table width="100%" border="0" cellspacing="1" cellpadding="0">
        <tr> 
                <td class="MessageTR1" width="20%" align="center"><b><%=stmp2%></b></td>
                <td class="MessageTR1" width="40%" align="center"><b><%=stmp3%></b></td>
                <td class="MessageTR1" width="40%" align="center"><b><%=stmp8%></b></td>
        </tr>
      
  
<%
 
 //Chequeo insiso por insiso...
   
 for (j = 0;j <= Session("Cantidad");j++)
  {
   i = 0; 
   ok = 0;
   resOpt = 0; 
   bien = 0;
   clase = "MessageTR1";
   
   response = Session("exer").items.response[Session("Cuales").items[j]];
   posi = response.indexOf("&",0);
   if (posi != -1)
    {pesos = response.substring(posi + 1,response.length);}
   response = response.substring(0,posi); 
   
   for (i = 0;i < Session("exer").items.cant[Session("Cuales").items[j]];i++)
    {
     cad = "O" + Session("exer").items.id[Session("Cuales").items[j]] + "#" + i;
     req = Request.Form.Item(cad) + "";
     //Response.Write(cad + "<br>");
     //Response.Write(req + "<br>");
     
     switch (Session("exer").items.type[Session("Cuales").items[j]])
      {
	    case "1":
	     { 
	      //Marcar... 
	      getResp(response,pesos,i);
          if (((rArr[0] == "T") && (req == 'true')) || ((rArr[0] == "F") && (req == 'false')))
           {ok = ok + parseInt(pArr[0],10);
            bien = bien + 1; 
           }
           else
            if ((rArr[0] == "T") && (req == 'false'))
              {ok = ok - parseInt(pArr[0],10);}     
           
	      break;
	     }
	    case "2":
	     {
	      //Enlazar...
	      getResp(response,pesos,i);
          
	      if (isokrArr(req.toUpperCase()) == true)               
           {
            ok = ok + parseInt(pe,10);
            bien = bien + 1;
           }
	      break;
	     }
	    case "3":
	     {
	     //Respuestas Multiples... 
	      getResp(response,pesos,i);
	      if (isokrArr(req.toUpperCase()) == true)               
           {ok = ok + parseInt(pe,10);
            bien = bien + 1;
           }
	      break;
	     }
	    case "4":
	     { //Ordenar...
	      getResp(response,pesos,i);

	      if (isokrArr(req.toUpperCase()) == true)               
           {ok = ok + parseInt(pe,10);
            bien = bien + 1;
           }
	      break;
	     } 
	    case "5":
	     { //Respuestas paralelas...
	      break;
	     }
	    case "6":
	     { //Ejercicio asitido...
	      
	      break;
	     }  
	   
      }
   
    } 
    
   if (clase ==  "MessageTR") clase = "MessageTR1"; else clase = "MessageTR";
   
   if (Session("exer").items.type[Session("Cuales").items[j]] != "6")
    {
     percent = getPercent(ok,resOpt); 
     if (percent >= 2)
      {eval[j] = percent;}
      else
       {eval[j] = 2;}
%>
        <tr> 
                <td class="<%=clase%>" width="20%" align="center"><%=(j + 1)%></td>
                <td class="<%=clase%>" width="40%" align="center"><%=bien%><%=stmp4%><%=Session("exer").items.cant[Session("Cuales").items[j]]%></td>
                <td class="<%=clase%>" width="40%" align="center"><%=eval[j]%></td>
        </tr>
    
<%     
    } 
   else
    {
     eval[j] = 0;

%>
        <tr> 
                <td class="<%=clase%>" width="20%" align="center"><%=(j + 1)%></td>
                <td class="<%=clase%>" width="40%" align="center"><%=stmp5%></td>
                <td class="<%=clase%>" width="40%" align="center"><%=stmp6%></td>
        </tr>
<%     
     
    } 
  }
  
%>
        <tr> 
          <td colspan="3" class="ToolBar"> 
            <table border="0" cellspacing="0" align="center" cellpadding="5">
              <tr> 
                <td  class="ToolBar"><a href="docs.asp?uid=<%=uid%>&lesson=<%=Session("lastLesson")%>&lessonurl=<%=Session("lastLessonUrl")%>&lessonindex=<%=Session("lastLessonIndex")%>&lessonname=<%=Session("lastLessonName")%>" class="ToolLink">&nbsp;<%=stmp7%>&nbsp;</a></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<%
 //insercion en la base de datos... 
 if (Session("userID") != GUEST_USER){insertEval(eval);   }
%>

</BODY>
</HTML>
<%
 }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>

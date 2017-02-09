<%@ Language=JavaScript %>


<%
 Response.Expires = -1;
%>
<!-- #include file="../../js/adolibrary.inc" -->
<!-- #include file="../../js/user.inc" -->

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       
      
 if ((Session("isTutor") > -1) || (Session("PermissionType") == ADMINISTRATOR) || (Session("userID") == Session("tutcursocordinador"))  || (Session("userID") == Request.QueryString.Item("courseOwner")))
    {
          var exer = Request.QueryString.Item("exer") + "";
          
          var dataPath = Application('dataPath');
    
          var oConn;
          var oRec;
 
          oConn = MakeConnection( oConn, dataPath );
          Sql ="SELECT Evaluaciones_Pendientes.Exercise, Evaluaciones_Pendientes.[User], Lecciones.dir, Lecciones.id,  Evaluaciones_Pendientes.URL, Evaluaciones_Pendientes.Response, Evaluaciones_Pendientes.Date" +
               " FROM Lecciones INNER JOIN Evaluaciones_Pendientes ON Lecciones.ID = Evaluaciones_Pendientes.Lesson " +
               " WHERE (((Evaluaciones_Pendientes.ID)= " + exer + ")) ";

          oRec = Query( Sql, oRec, oConn  );		
          
          if (!oRec.EOF) {
            var response = oRec.Fields.Item("Response").value;
            var url = oRec.Fields.Item("URL").value;
            var dir = "";
            if ((oRec.Fields.Item("dir").value != null) && (oRec.Fields.Item("dir").value != "")) dir = "/" + oRec.Fields.Item("dir").value;
            var user = oRec.Fields.Item("User").value;
            var date = oRec.Fields.Item("Date").value;
            var lesson = oRec.Fields.Item("id").value;
            var exercise = oRec.Fields.Item("Exercise").value;
            var username = Request.QueryString.Item("username") + "";
            //Response.Write(response);
          }
          else
            {
              DestroyAdoObjects( oConn, oRec );  
              Response.Redirect("../errorpage.asp?tipo=Error&short=" + NOT_EXER_EXIST_SHORT  + "&desc=" + NOT_EXER_EXIST_TEXT);              
            }    
          
          DestroyAdoObjects( oConn, oRec );  
      
%>

<html>
<head>
<title>Ejercicio pendiente de revisión</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/Main.css" type="text/css">
<script language=jscript src="../../js/ejercicios.js"></script>


</head>

<body bgcolor="#FFFFFF" text="#000000" >
  <table border="0" cellspacing="0" cellpadding="2" width="100%" class="ToolBar" >
    <tr> 
      <td align="center" width="100%" >Ejercicio pendiente</td>
    </tr>
  </table>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td align="center"> 
      <table width="100%" border="0" cellspacing="5" cellpadding="0" class="MessageTR1">
        <tr>
          <td width="70%" align="left" ><b>Realizado por:&nbsp;&nbsp </b><%=username%></td><td nowrap width="30%" align="right"><%=date%></td>
        </tr>
      </table>
      <table border="0" cellspacing="0" cellpadding="2" width="100%">
        <tr> 
          <td colspan="2" width="100%" class="MessageTR1">
            <hr noshade>
            <SCRIPT LANGUAGE=javascript src='../../Courses/course<%=Session("tutcurso") + "/" + url%>' ></SCRIPT>
            <textarea readonly name="FileDescription" rows="13" cols="78"><%=response%></textarea>
          </td>
        </tr>
      </table>
      <table border="0" cellspacing="0" cellpadding="2" width="100%">
        <tr> 
          <td  width="40%" class="MessageTR1">
          </td>
        
          <td  width="10%" class="MessageTR1">
            <b>Calificación: </b>
          </td>
          <td align="left" width="10%" class="MessageTR1">
            <select name="eval" id="eval" >
               <option value=2>2</option>
			   <option value=3>3</option>
               <option value=4>4</option>
               <option selected value=5>5</option>
            </select>
          </td>
          <td  width="40%" class="MessageTR1">
          </td>
          
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td class="ToolBar" align="right"> 
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td><a href="javascript:gogogo()" class="ToolLink">&nbsp;Registrar evaluación&nbsp;</a></td>
          <td>|</td>

          <td><a href="javascript:history.back()" class="ToolLink" >&nbsp;Regresar&nbsp;</a></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</form>

</body>
</html>
<SCRIPT LANGUAGE=javascript>
<!--
 
 
 function gogogo() {
   var sl = document.all.tags("SELECT").item(0);
   var p = sl.selectedIndex;
   var ev = sl.options(p).value;   
   location.href = "makevaluation.asp?uid=<%=uid%>&courseOwner=<%=Request.QueryString.Item("courseOwner")%>&userid=<%=user%>&lesson=<%=lesson%>&" + 
                   "exercise=<%=exercise%>&puntuation=" + ev + 
                   "&id=<%=exer%>&date=<%=date%>";
 }

//-->
</SCRIPT>
<SCRIPT LANGUAGE=javascript>
<!--

  document.body.getElementsByTagName("TEXTAREA").item(0).style.display  = "none";
  
//-->
</SCRIPT>

<%
      }
    else
     {  
      Response.Redirect("../errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);     
     }

  }
 else
   Response.Redirect("../errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
  //Response.Write(Request.QueryString.Item("uid"));
%>


<%@ Language=JavaScript%>
<%
Response.Expires = -1;
%>

<!-- #include file='../js/adolibrary.inc' -->
<!-- #include file='../js/user.inc' --> 

<% 

if (Request.QueryString.Item("uid") + "" == Session("uid"))
{    
  var uid = Request.QueryString.Item("uid") + "";       
    //se reciben los parametros courseID,courseName, lastlesson, lastlessonName
  if (Request.QueryString.Count != 3)   
  { 
   //Error en el pase de parametros.
    Response.Redirect("errorpage.asp?tipo=Error&short=" + BAD_PARAMS_SHORT  + "&desc=" + BAD_PARAMS_TEXT);
  }
  
  var oConn;
  var oRec;
 
  var filePath = Application('dataPath');
    
  oConn = MakeConnection( oConn, filePath );       
  
   
  Sql ="SELECT * FROM faq WHERE (course = " + Request.QueryString("IdCurso") + ") ORDER BY question";

  oRec = Query( Sql, oRec, oConn  );
         
%>    
<html>
<head>
<title>&nbsp;</title>
<link rel="stylesheet" href="../Css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../Css/<%=Session("skin")%>/EvalResult.css" type="text/css">
</head>
<body  class="MainBody" style="overflow-y:auto">
<table width="100%" border="0" cellspacing="0" cellpadding="2">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td><img src="../images/<%=Session("skin")%>/EvalResult.gif" width="80" height="54"></td>
          <td class="HeaderTable">Preguntas Frecuentes</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<%
  var i = 1;
  while(!oRec.EOF)
  {
%>
<table width="100%"  border="0" cellspacing="0" cellpadding="2">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td width="3%"><strong>Pregunta #<%=i%>: </strong><%=oRec.Fields.Item("question") %></td>
  </tr>
  <tr>
    <td><strong>Respuesta: </strong><%=oRec.Fields.Item("response") %></td>
  </tr>
</table>
<%
	oRec.Move(1);
	i++;
  }
%>
</body>
</html>
<%
}
%>
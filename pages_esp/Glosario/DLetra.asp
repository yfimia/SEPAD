<%@language=jscript%>
<%
  
  var IC      = Request.QueryString.Item("IdCurso") * 1;
  var Let     = Request.QueryString.Item("Letra");
  var CName   = Request.QueryString.Item("courseName");

  Response.Redirect("glosario.asp?IdCurso=" + IC + "&Letra=" + Let + "&IDP=0&courseName=" + CName);
  
%>
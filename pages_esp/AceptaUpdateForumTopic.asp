<%@ Language=Jscript%>
<%
/*
    Esta página acepta los temas modificados y los pone en la base de datos. Los 
    mensajes referentes a los foros de discusión.
    
    PARAMETROS
    ----------
    
    uid        -- El ID que genera la sesión cuando el usuario se conecta 
                  por primera vez.
                  
    Course     -- ID del curso al que se le piden los temas de discusión.
    
    courseName -- Nombre del curso
    
    Los demás parámetros vienen en una forma desde la pagina: NewForumTopic.asp

    pagenumber -- numero del bloque o pagina que se esta mostrando
    
    topicID    -- id del tema a modificar
    
  */

  Response.Expires = -1;
%>

<!-- #include file='../js/user.inc' -->
<!-- #include file= 'inc/AceptaNewForumTopic.inc' -->
<%

  var topicID = "";
  if (Request.QueryString.Item("topicID").Count != 0) {
    topicID = Request.QueryString.Item("topicID") + "";
  }
  
  var uid = "";
  if (Request.QueryString.Item("uid").Count != 0) {
    uid = Request.QueryString.Item("uid") + "";
  }

  if (uid == Session("uid"))
   {    
   
      var CourseInUse = 0;
      if (Request.QueryString.Item("Course").Count != 0) {
        CourseInUse = Request.QueryString.Item("Course")
      }  

      var pagenumber = 0;
      if (Request.QueryString.Item("pn").Count != 0) {
        pagenumber = Request.QueryString.Item("pn")
      }  

      var courseNameInUse = "";
      if (Request.QueryString.Item("courseName").Count != 0) {
        courseNameInUse = Request.QueryString.Item("courseName")
      }  
      
      if (Session("valid")){
%>

<SCRIPT language=vbscript RUNAT=Server>
 function getDate
  getDate = Now()
 End function 
</SCRIPT>

<% 
 
 MTitulo = Request.Form("Ttitulo") + "";
 MBody = Request.Form("TBody") + "";  
 MAcepta = Request.Form("TAcepta");
 
 MBody = MBody.replace(/\n/g,"<br>");
 MBody = MBody.replace(/  /g,"&nbsp;&nbsp;");

 
 var oConn    = Server.CreateObject("ADODB.Connection");  
 var oRec     = Server.CreateObject("ADODB.Recordset");
 var filePath = Application("filePath");  
 var Sql      = "UPDATE ForumTopics SET Titulo = '" + MTitulo + "', Descripcion = '" + MBody + "' WHERE id=" + topicID;
 
 oConn.Open( filePath );   
 oRec.Open(Sql, oConn, 3, 3);    
 
 oConn.Close(); 

%>


<HTML>
<HEAD>
</HEAD>

<BODY>

<SCRIPT LANGUAGE=JSCRIPT>
  opener.location = "ShowForumTopics.asp?uid=<%=uid%>&Course=<%=CourseInUse%>&pn=<%=pagenumber%>&courseName=<%=courseNameInUse%>";
  close();
</SCRIPT>

</BODY>
</HTML>
<%
      }    
      else {
        Response.Redirect("errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);                   
      }
    }  
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>


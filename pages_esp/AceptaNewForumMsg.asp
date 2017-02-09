<%@ Language=Jscript%>
<%
/*
    Esta página acepta los mensajes y los pone en la base de datos. Los 
    mensajes referentes a los foros de discusión.
    
    PARAMETROS
    ----------
    
    uid        -- El ID que genera la sesión cuando el usuario se conecta 
                  por primera vez.
                  
    Course     -- ID del curso al que se le piden los temas de discusión.
    
    TopicID    -- ID del Tema seleccionado del cual se desean los mensajes
    
    courseName -- Nombre del curso
    
    pagenumber -- numero del bloque o pagina que se esta mostrando
    
    Los demás parámetros vienen en una forma desde la pagina: NewForumMsg.asp

  */

  Response.Expires = -1;
%>

<!-- #include file='../js/user.inc' -->
<!-- #include file= 'inc/AceptaNewForumMsg.inc' -->
<%
   if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       
      var CourseInUse = 0;
      if (Request.QueryString.Item("Course").Count != 0) {
        CourseInUse = Request.QueryString.Item("Course")
      }  

      var TopicID = 0;
      if (Request.QueryString.Item("Topic").Count != 0) {
        TopicID = Request.QueryString.Item("Topic")
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
        var  filePath = Application("dataPath");   
        var  oConn    = Server.CreateObject("ADODB.Connection");
        var  oRec     = Server.CreateObject("ADODB.Recordset");
        oConn.Open(filePath);

        //Chequeo si esta habilitado el tema
        oRec.Open("SELECT * FROM ForumTopics WHERE (ID = " + TopicID + ") and (Activo = " + Application("dtrue") + ")",oConn,3,3);
        if (oRec.EOF == false) {
          oRec.Close();
      
      

%>

<SCRIPT language=vbscript RUNAT=Server>
 function getDate
  getDate = Now()
 End function 
</SCRIPT>

<% 
 
 MTopic = Request.Form("MsgTopic") + "";
 MBody = Request.Form("MsgBody") + "";  
 MAcepta = Request.Form("MsgAcepta");
 
 MBody = MBody.replace(/\n/g,"<br>");
 MBody = MBody.replace(/  /g,"&nbsp;&nbsp;");
 
 
 var oConn    = Server.CreateObject("ADODB.Connection");  
 var oRec     = Server.CreateObject("ADODB.Recordset");
 var filePath = Application("filePath");  
 var Sql      = "SELECT * FROM ForumMessages";
 
 oConn.Open( filePath );   
 oRec.Open(Sql, oConn, 3, 3);    
 
 //Entrada de datos del mensaje
 
 oRec.AddNew();
 oRec.Fields.Item("User").value             = Session("userid");
 oRec.Fields.Item("texto").value            = MBody;
 oRec.Fields.Item("ReadingCount").value     = 0;
 oRec.Fields.Item("AceptaRespuestas").value = MAcepta; 
 oRec.Fields.Item("Topic").value			  = MTopic; 
 oRec.Update();
 
 //*********************************
 
 oConn.Close(); 
 
%>


<HTML>
<HEAD>
</HEAD>

<BODY>

<SCRIPT LANGUAGE=JSCRIPT>
  opener.location = "ShowForumMessages.asp?uid=<%=uid%>&Course=<%=CourseInUse%>&Topic=<%=TopicID%>&pn=<%=pagenumber%>&courseName=<%=courseNameInUse%>";
  close();
</SCRIPT>

</BODY>
</HTML>
<%
        }
        else{
          Response.Redirect("errorpage.asp?tipo=Error&short=Mensaje inactivo" + "&desc=El mensaje seleccionado no está activado para el uso" );                   
        }
      }    
      else {
        Response.Redirect("errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);                   
      }
   }
   else 
     Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>


<%@ Language=Jscript%>
<%
/*
    Esta página acepta las opiniones y las pone en la base de datos. 
    Las opiniones son referentes a los mensajes de los foros de discusión.
    
    PARAMETROS
    ----------
    
    uid        -- El ID que genera la sesión cuando el usuario se conecta 
                  por primera vez.
                  
    Course     -- ID del curso al que se le piden los temas de discusión.
    
    courseName -- Nombre del curso
    
    Los demás parámetros vienen en una forma desde la pagina: NewForumMsg.asp

  */

  Response.Expires = -1;
%>

<!-- #include file='../js/user.inc' -->
<!-- #include file= 'inc/AceptaNewForumAnswer.inc' -->
<%
   if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       
      var CourseInUse = 0;
      if (Request.QueryString.Item("Course").Count != 0) {
        CourseInUse = Request.QueryString.Item("Course")
      }  

      var pagenumber = 0;
      if (Request.QueryString.Item("pn").Count != 0) {
        pagenumber = Request.QueryString.Item("pn")
      }  

      var msgid = 0;
      if (Request.QueryString.Item("msgid").Count != 0) {
        msgid = Request.QueryString.Item("msgid")
      }  
      
      var TopicID = 0;
      if (Request.QueryString.Item("Topic").Count != 0) {
        TopicID = Request.QueryString.Item("Topic");
      }  

      var courseNameInUse = "";
      if (Request.QueryString.Item("courseName").Count != 0) {
        courseNameInUse = Request.QueryString.Item("courseName")
      }  

      if (Session("valid")) {
        var  filePath = Application("dataPath");   
        var  oConn    = Server.CreateObject("ADODB.Connection");
        var  oRec     = Server.CreateObject("ADODB.Recordset");
        oConn.Open(filePath);
        
        oRec.Open("SELECT ForumAnswers.ID AS AID, ForumMessages.ID AS IDMSG1, ForumMessages.AceptaRespuestas MActivo, ForumTopics.ID AS TID, ForumTopics.Activo AS TActivo " +
                  "FROM ForumTopics INNER JOIN (ForumMessages INNER JOIN ForumAnswers ON ForumMessages.ID = ForumAnswers.Origen) ON ForumTopics.ID = ForumMessages.Topic " +
                  "WHERE (ForumMessages.ID = " + msgid + ")",oConn,3,3);
        
        if ((oRec.EOF == true) || ((oRec.Fields.Item("MActivo").value == true) && (oRec.Fields.Item("TActivo").value == true))){
          oRec.Close(); 

%>

<SCRIPT language=vbscript RUNAT=Server>
 function getDate
  getDate = Now()
 End function 
</SCRIPT>

<% 
 
 msgid = Request.Form("msgid") + "";
 MBody  = Request.Form("MsgBody") + "";  
 MAcepta = Request.Form("MsgAcepta");
 
 MBody = MBody.replace(/\n/g,"<br>");
 MBody = MBody.replace(/  /g,"&nbsp;&nbsp;");

 
 var oConn    = Server.CreateObject("ADODB.Connection");  
 var oRec     = Server.CreateObject("ADODB.Recordset");
 var filePath = Application("filePath");  
 var Sql      = "SELECT * FROM ForumAnswers";
 
 oConn.Open( filePath );   
 oRec.Open(Sql, oConn, 3, 3);    
 
 //Entrada de datos de la opinion
 
 oRec.AddNew();
 oRec.Fields.Item("User").value             = Session("userid");
 oRec.Fields.Item("texto").value            = MBody;
 oRec.Fields.Item("Origen").value			= msgid; 
 oRec.Update();
 
 //*********************************
 
 oConn.Close(); 
 
%>


<HTML>
<HEAD>
  <title>
   Tema <%=TopicID%>
  </title>
</HEAD>

<BODY>

<SCRIPT LANGUAGE=JSCRIPT>
  opener.location = "ShowForumanswer.asp?uid=<%=uid%>&Course=<%=CourseInUse%>&Topic=<%=TopicID%>&msgid=<%=msgid%>&pn=<%=pagenumber%>&courseName=<%=courseNameInUse%>";
  close();
</SCRIPT>

</BODY>
</HTML>
<%            
          }
        else
          {
            Response.Redirect("errorpage.asp?tipo=Error&short='Operación no válida'&desc='El mensaje escogido no permite enviarle opiniones'");                   
          }
      }    
      else {
        Response.Redirect("errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);                   
      }
    }  
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>

<%@ Language=JScript %>
<%
  /*
    Esta página administra el listado de los temas de discusión que 
    pertenece al curso dado.
    
    PARAMETROS
    ----------
    
    uid        -- El ID que genera la sesión cuando el usuario se conecta 
                  por primera vez.
                  
    Course     -- ID del curso al que se le piden los temas de discusión.
    
    courseName -- Nombre del curso
    
    lasttopic  -- ID del ultimo Tema a partir del cual se comenzará a 
                  mostrar las listas de temas
                  
    op         -- Contiene la operación a efectuar sobre el tema seleccionado
    
    tema	   -- Contiene el ID del tema seleccionado

  */
  Response.Expires = -1;
%>
<!-- #include file='../js/adolibrary.inc' -->
<!-- #include file='../js/user.inc' -->
<!-- #include file= 'inc/AdminForumTopics.inc' -->
<%    
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
      
      var courseNameInUse = "";
      if (Request.QueryString.Item("courseName").Count != 0) {
        courseNameInUse = Request.QueryString.Item("courseName")
      }  

      var MlastTopic = -1;
      if (Request.QueryString.Item("lasttopic").Count != 0) {
        MlastTopic = Request.QueryString.Item("lasttopic")
      }  

      var op = "none";
      if (Request.QueryString.Item("op").Count != 0) {
        op = Request.QueryString.Item("op")
      }  

      var tema = -1;
      if (Request.QueryString.Item("tema").Count != 0) {
        tema = Request.QueryString.Item("tema")
      }  
      
      var adcourse = ((isUserInGroupByID(Session("UserID"),ADMIN_GROUP)) || (Session("permissionType") == ADMINISTRATOR)); 
      //var valid = ((isUserInGroup(Session("UserID"),'Curso_' + courseNameInUse)) || adcourse);
      var valid = true;
      
      if (valid) {
        var  filePath = Application("dataPath");   
        var  oConn    = Server.CreateObject("ADODB.Connection");
        var  oComm    = Server.CreateObject("ADODB.Command");
        var  oRec     = Server.CreateObject("ADODB.Recordset");
        oRec.CursorLocation = 3;
        oRec.CursorType     = 3;
        oConn.Open(filePath);
        oComm.ActiveConnection = oConn;
        
        if ((op == "on") && (tema != -1)) {
          oComm.CommandText = "UPDATE ForumTopics SET Activo = " + Application("dtrue") + " WHERE ID = " + tema + "";
          oComm.execute();
        }
        
        if ((op == "off") && (tema != -1)) {
          oComm.CommandText = "UPDATE ForumTopics SET Activo = " + Application("dfalse") + " WHERE ID = " + tema + "";
          oComm.execute();
        }
        
        var manipulador = ((adcourse) || (isUserInGroup(Session("UserID"),'Tutores_Curso_' + courseNameInUse)));
        //Forum Topics
        if (manipulador) {
          oComm.CommandText = "SELECT Top " + SHOW_CANT + " ForumTopics.ID AS FID, ForumTopics.Curso, ForumTopics.Titulo, ForumTopics.Descripcion AS DForum, ForumTopics.Fecha, ForumTopics.Activo, Usuarios.FullName, Usuarios.ID " + 
                              "FROM Usuarios INNER JOIN ForumTopics ON Usuarios.ID = ForumTopics.[User] " + 
                              "WHERE (ForumTopics.curso = " + CourseInUse + ") and (ForumTopics.ID >= " + MlastTopic + ")";
        }                    
        else {
          oComm.CommandText = "SELECT Top " + SHOW_CANT + " ForumTopics.ID AS FID, ForumTopics.Curso, ForumTopics.Titulo, ForumTopics.Descripcion AS DForum, ForumTopics.Fecha, ForumTopics.Activo, Usuarios.FullName, Usuarios.ID " + 
                              "FROM Usuarios INNER JOIN ForumTopics ON Usuarios.ID = ForumTopics.[User] " + 
                              "WHERE (ForumTopics.curso = " + CourseInUse + ") and (ForumTopics.ID >= " + MlastTopic + ") and (ForumTopics.[User] = " + Session("UserID") + ")";
        }
        
        oRec = oComm.Execute();    
%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<script type="text/javascript" language="javascript" src="../js/Windows.js"></script>
</HEAD>

<body bgcolor="#ffffff" text="#000000">

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar" align="middle"><h1><b><%=titulo%></b></h1></td>
  </tr>
</table> 

<br>

<table width="98%" border="0" cellspacing="0" cellpadding="0" align=center>
  <tr> 
    <td> 
      <table width="100%" border="0" cellspacing="1" cellpadding="0">
        <tr> 
          <td width="1%" class="ToolBar" align="middle"></td>
          <td width="34%" class="ToolBar" align="middle"><b><%=tema%></b></td>
          <td width="15%" class="ToolBar" align="middle"><b><%=fecha%></b></td>
          <td width="15%" class="ToolBar" align="middle"><b><%=estado%></b></td>
        </tr>
<%    
        var clase = "MessageTR1";
        var i = 0;
        var last = "-1"
        var Fimage = "../images/<%=Session("skin")%>/icon_folder.gif";
        var msgNext = true;
        
        while (oRec.EOF == false)
          { 
            i = i + 1;
            if (clase == "MessageTR1") {clase = "MessageTR";} else {clase = "MessageTR1";}
                      
            if (oRec.Fields.Item("Activo").value == true){
              Fimage = "../images/<%=Session("skin")%>/icon_folder.gif";
            }
            else {
              Fimage = "../images/<%=Session("skin")%>/icon_folder_locked.gif";
            }
            
            var topicURL = "ShowForumMessages.asp?uid=" + uid + "&Course=" + CourseInUse + "&Topic=" + oRec.Fields.Item("FID").value + "&lastmsg=0&courseName=" + courseNameInUse;
               
%>
       
        <tr> 
          <td width="1%" class="<%=clase%>" align="middle">
            <img src="<%=Fimage%>" border=0>
          </td>
          <td class="<%=clase%>" align="middle">
            <b><%=oRec.Fields.Item("titulo").value%></b>
          </td>
          <td class="<%=clase%>" align="middle">
            <%=oRec.Fields.Item("Fecha").value%>
          </td>
          <td class="<%=clase%>" align="middle">
            <%if (oRec.Fields.Item("Activo").value == true){%>
              <a href="AdminForumTopics.asp?uid=<%=uid%>&Course=<%=CourseInUse%>&lasttopic=<%=MlastTopic%>&op=off&tema=<%=oRec.Fields.Item("FID").value%>&courseName=<%=courseNameInUse%>"> 
                <b><%=block%></b>
              </a>  
            <%}
              else {%>
              <a href="AdminForumTopics.asp?uid=<%=uid%>&Course=<%=CourseInUse%>&lasttopic=<%=MlastTopic%>&op=on&tema=<%=oRec.Fields.Item("FID").value%>&courseName=<%=courseNameInUse%>"> 
                <b><%=desblock%></b>
              </a>  
            <%}%>  
          </td>
        </tr> 
<%          lastTopic = oRec.Fields.Item("FID").value;
            oRec.Move(1);
          }  
        if (i > 0)
          {
            oRec.MoveFirst();
            if (i < SHOW_CANT) 
              msgNext = false
            else
              msgNext = true;
          }
        else
          {
            lastTopic = 0;
            msgNext = false;
%>        
            <tr><td align="middle" width="100%" colspan="5" class="MessageTR"><%=message%></td></tr>        
<%          
          }  
%>      
          <tr>
            <td class="ToolBar" colspan=5>
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td Width="10%" class="ToolBar" align=center>
<%                  if (msgNext){%>                    
                    <a href="AdminForumTopics.asp?uid=<%=uid%>&Course=<%=CourseInUse%>&lasttopic=<%=lastTopic%>&courseName=<%=courseNameInUse%>" class="ToolLink">&nbsp;<%=prox%>&nbsp;<%=SHOW_CANT%>&nbsp;</a>
<%                  }
                    else{%>
                    <a href="AdminForumTopics.asp?uid=<%=uid%>&Course=<%=CourseInUse%>&lasttopic=-1&courseName=<%=courseNameInUse%>" class="ToolLink">&nbsp;<%=init%>&nbsp;</a>
<%                  }%>                    
                  </td>
                  <td Width="99%" class="ToolBar">
                  </td>
                </tr> 
              </table>  
            </td>
          </tr>  
      </table>  
    </td>  
  </tr>
</TABLE>  
<%            
      }    
      else {
        Response.Redirect("errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);                   
      }
    }  
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>
</body>
</HTML>

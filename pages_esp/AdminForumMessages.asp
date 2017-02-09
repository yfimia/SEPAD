<%@ Language=JScript %>
<%
  /*
    Esta página muestra el listado de los mensajes relativos a un tema 
    de discusión que pertenece a un curso dado.
    
    PARAMETROS
    ----------
    
    uid        -- El ID que genera la sesión cuando el usuario se conecta 
                  por primera vez.
                  
    Course     -- ID del curso al que se le piden los temas de discusión.
    
    courseName -- Nombre del curso
    
    TopicID    -- ID del Tema seleccionado del cual se desean los mensajes
    
    lastmsg    -- ID del mensaje a partir del cual se comenzará a  mostrar 
                  la lista de mensajes

    op         -- Contiene la operación a efectuar sobre el mensaje seleccionado
    
    idmsg  	   -- Contiene el ID del mensaje seleccionado
                  
  */

  Response.Expires = -1;
%>
<!-- #include file='../js/adolibrary.inc' -->
<!-- #include file='../js/user.inc' -->
<!-- #include file='inc/AdminForumMessages.inc' -->
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

      var TopicID = 0;
      if (Request.QueryString.Item("Topic").Count != 0) {
        TopicID = Request.QueryString.Item("Topic")
      }  
      
      var courseNameInUse = "";
      if (Request.QueryString.Item("courseName").Count != 0) {
        courseNameInUse = Request.QueryString.Item("courseName")
      }  

      var op = "none";
      if (Request.QueryString.Item("op").Count != 0) {
        op = Request.QueryString.Item("op")
      }  

      var idmsg = -1;
      if (Request.QueryString.Item("idmsg").Count != 0) {
        idmsg = Request.QueryString.Item("idmsg")
      }  


      var lastmsg = -1;
      if (Request.QueryString.Item("lastmsg").Count != 0) {
        lastmsg = Request.QueryString.Item("lastmsg")
      } 
      var Mlastmsg = lastmsg; 
      
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

        if ((op == "on") && (idmsg != -1)) {
          oComm.CommandText = "UPDATE ForumMessages SET AceptaRespuestas = "+Application("dtrue")+" WHERE ID = " + idmsg + "";
          oComm.execute();
        }
        
        if ((op == "off") && (idmsg != -1)) {
          oComm.CommandText = "UPDATE ForumMessages SET AceptaRespuestas = "+Application("dfalse")+" WHERE ID = " + idmsg + "";
          oComm.execute();
        }

        var manipulador = ((adcourse) || (isUserInGroup(Session("UserID"),'Tutores_Curso_' + courseNameInUse)));
        //Forum messages
        if (manipulador) {
          oComm.CommandText = "SELECT Top " + SHOW_CANT + " ForumMessages.ID,ForumMessages.[User],ForumMessages.Texto, ForumMessages.Fecha, ForumMessages.ReadingCount, ForumMessages.AceptaRespuestas, ForumMessages.Topic, ForumTopics.Curso,ForumTopics.Activo,ForumTopics.titulo, Usuarios.FullName, Usuarios.Email " + 
                              "FROM Usuarios INNER JOIN (ForumTopics INNER JOIN ForumMessages ON ForumTopics.ID = ForumMessages.Topic) ON Usuarios.ID = ForumMessages.[User] " + 
                              "WHERE (ForumMessages.Topic = " + TopicID + ") and (ForumMessages.ID >= " + lastmsg + ") and (ForumTopics.Activo = " + Application("dtrue") + ")";
        } 
        else {                      
          oComm.CommandText = "SELECT Top " + SHOW_CANT + " ForumMessages.ID,ForumMessages.[User],ForumMessages.Texto, ForumMessages.Fecha, ForumMessages.ReadingCount, ForumMessages.AceptaRespuestas, ForumMessages.Topic, ForumTopics.Curso,ForumTopics.Activo,ForumTopics.titulo, Usuarios.FullName, Usuarios.Email " + 
                              "FROM Usuarios INNER JOIN (ForumTopics INNER JOIN ForumMessages ON ForumTopics.ID = ForumMessages.Topic) ON Usuarios.ID = ForumMessages.[User] " + 
                              "WHERE (ForumMessages.Topic = " + TopicID + ") and (ForumMessages.ID >= " + lastmsg + ") and (ForumTopics.Activo = " + Application("dtrue") + ") and (ForumMessages.[User] = " + Session("UserID") + ")";
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

<table width="98%" border="0" cellspacing="0" cellpadding="0" align=center>
  <tr>
    <td>
      <table border="0" cellspacing="3" cellpadding="0">
        <tr>      
          <td align=left width="1%">
            <IMG src="../images/<%=Session("skin")%>/icon_folder_open.gif" border=0>
          </td>
          <td align=left colspan=2>
            <%=infor1%>
          </td>
          <td>
          </td>
        </tr>
        <tr>      
          <td align=left width="1%">
            <IMG src="../images/<%=Session("skin")%>/icon_bar.gif" border=0>
          </td>
          <td align=left width="1%">
            <IMG src="../images/<%=Session("skin")%>/icon_folder_open_topic.gif" border=0>
          </td>
          <td align=left>
            <%=infor2%>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td> 
      <table width="100%" border="0" cellspacing="1" cellpadding="0">
        <tr> 
          <td width="1" class="ToolBar" align="middle"></td>       
          <td width="15%" class="ToolBar" align="middle"><b><%=autor%></b></td>
          <td width="84%" class="ToolBar" align="middle"><b><%=texto%></b></td>
        </tr>
<%    
        var clase = "MessageTR1";
        var i = 0;
        var Fimage = "../images/<%=Session("skin")%>/icon_folder.gif";
        var msgNext = true;
        
        while (oRec.EOF == false)
          { 
            
            if (clase == "MessageTR1") clase = "MessageTR"; else clase = "MessageTR1";
                      
            if ((oRec.Fields.Item("AceptaRespuestas").value == true) && (oRec.Fields.Item("Activo").value == true)) {
              Fimage = "../images/<%=Session("skin")%>/icon_folder.gif";
            }
            else {
              Fimage = "../images/<%=Session("skin")%>/icon_folder_locked.gif";
            }
            
            
            if (oRec.Fields.Item("Activo").value == true){
              i = i + 1;
              %>
              
        <tr> 
          <td width="1%" class="<%=clase%>" align="middle">
            <img src="<%=Fimage%>" border=0>
          </td>
          <td class="<%=clase%>" align="middle">
            <b><%=oRec.Fields.Item("Fullname").value%></b>
          </td>
          <td class="<%=clase%>" align=left>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="1%">
                  <IMG src="../images/<%=Session("skin")%>/icon_posticon.gif">
                </td> 
                <td width="100%" align=left valign=bottom nowrap>
                  <font size=2><%=enviado><%=oRec.Fields.Item("Fecha").value%></font>
                </td>  
                <td width="1%" align=left valign=middle nowrap>
                  &nbsp;
                  <%if (oRec.Fields.Item("Aceptarespuestas").value == true){%>
                    <a href="AdminForumMessages.asp?uid=<%=uid%>&Course=<%=CourseInUse%>&topic=<%=TopicID%>&lastmsg=<%=Mlastmsg%>&op=off&idmsg=<%=oRec.Fields.Item("ID").value%>&courseName=<%=courseNameInUse%>">
                      <font size=1><%=noacept%></font>
                    </a>
                  <%} else {%>
                    <a href="AdminForumMessages.asp?uid=<%=uid%>&Course=<%=CourseInUse%>&topic=<%=TopicID%>&lastmsg=<%=Mlastmsg%>&op=on&idmsg=<%=oRec.Fields.Item("ID").value%>&courseName=<%=courseNameInUse%>">
                      <font size=1><%%=siacept></font>
                    </a>
                  <%}%>
                </td>
              </tr>
            </table>  
            <HR noShade SIZE=1>
            <%=oRec.Fields.Item("Texto").value%>
          </td>
        </tr> 
<%          
             
           }
           lastmsg = oRec.Fields.Item("ID").value;
           oRec.Move(1);
         }  
        
        if (i > 0) 
          {
            //oRec.MoveFirst();
            if (i < SHOW_CANT)
              msgNext = false
            else
              msgNext = true;
          }
        else
          {
            lastmsg = 0;
            msgNext = false;
%>        
            <tr><td align="middle" width="100%" colspan="5" class="MessageTR"><%=noelem%></td></tr>        
<%          
          }  
%>      
          <tr>
            <td class="ToolBar" colspan=5>
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td Width="10%" class="ToolBar" align=middle>
<%                  if (msgNext){%>                    
                    <A class=ToolLink href="AdminForumMessages.asp?uid=<%=uid%>&Course=<%=CourseInUse%>&topic=<%=TopicID%>&lastmsg=<%=lastmsg%>&courseName=<%=courseNameInUse%>">&nbsp;<%=prox%>&nbsp;<%=SHOW_CANT%>&nbsp;</a>
<%                  }
                    else{%>
                    <A class=ToolLink href="AdminForumMessages.asp?uid=<%=uid%>&Course=<%=CourseInUse%>&topic=<%=TopicID%>&lastmsg=-1&courseName=<%=courseNameInUse%>">&nbsp;<%=init%>&nbsp;</a>
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
</table>  
<%            
      }    
      else {
        Response.Redirect("errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);                   
      }
    }  
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>
</TR></TABLE>
</body>
</HTML>

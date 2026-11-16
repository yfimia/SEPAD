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
    
    pagenumber -- numero del bloque o pagina que se esta mostrando

    op         -- Operacion extra a efectuar antes de cargar la pagina
                    Lista de operaciones
                    --------------------
                    0 - No operacion
   
                    1 - Eliminar mensaje
                          msgid - trae el ID del mensaje a borrar
              
                    2 - Activar mensaje
                          msgid - trae el ID del mensaje a poner activo
              
                    3 - Desactivar mensaje
                          msgid - trae el ID del mensaje a desactivar
                  
  */

  Response.Expires = -1;
%>
<!-- #include file='../js/adolibrary.inc' -->
<!-- #include file='../js/user.inc' -->
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

      var pagenumber = 0;
      if (Request.QueryString.Item("pn").Count != 0) {
        pagenumber = Request.QueryString.Item("pn")
      }  
      
      if (Session("valid")) {
        var  filePath = Application("dataPath");   
        var  oConn    = Server.CreateObject("ADODB.Connection");
        var  oRec     = Server.CreateObject("ADODB.Recordset");
        oConn.Open(filePath);

        //Chequeo si esta habilitado el tema
        oRec.Open("SELECT * FROM ForumTopics WHERE (ID = " + TopicID + ") and (Activo = " + Application("dtrue") + ")",oConn,3,3);
        if (oRec.EOF == false) {
          oRec.Close();

          operation = 0;
          if (Request.QueryString.Item("op").count != 0){
            operation = Request.QueryString.Item("op");
          }
  
          //Chequeo las operaciones
        
          var TopicId = -1;
          if (Request.QueryString.Item("Topic").Count != 0) {
            TopicId = Request.QueryString.Item("Topic")
          }  

          var msgid = -1;
          if (Request.QueryString.Item("msgid").Count != 0) {
            msgid = Request.QueryString.Item("msgid")
          }  

          if (operation == 1){
            if ( (Session("PermissionType") == ADMINISTRATOR) || (Session("cordinador") = Session("userID")) || (Session("courseOwner") = Session("userID")) || (oRec.Fields.Item("ID").value == Session("userID")) ){
              oRec.Open("DELETE FROM ForumMessages WHERE ID = " + msgid , oConn, 3, 3);    
            }  
          }
        
          if (operation == 2){
            if ( (Session("PermissionType") == ADMINISTRATOR) || (Session("cordinador") = Session("userID")) || (Session("courseOwner") = Session("userID")) || (oRec.Fields.Item("ID").value == Session("userID")) ){
              oRec.Open("UPDATE ForumMessages SET AceptaRespuestas = " + Application("dtrue") + " WHERE ID = " + msgid , oConn, 3, 3);
            }  
          }
        
          if (operation == 3){
            if ( (Session("PermissionType") == ADMINISTRATOR) || (Session("cordinador") = Session("userID")) || (Session("courseOwner") = Session("userID")) || (oRec.Fields.Item("ID").value == Session("userID")) ){
              oRec.Open("UPDATE ForumMessages SET AceptaRespuestas = " + Application("dfalse") + " WHERE ID = " + msgid , oConn, 3, 3);
            }      
          }
        
          //Forum messages
          oRec.Open("SELECT ForumMessages.ID,ForumMessages.[User],ForumMessages.Texto, ForumMessages.Fecha, ForumMessages.ReadingCount, ForumMessages.AceptaRespuestas, ForumMessages.Topic, ForumTopics.Curso,ForumTopics.Activo,ForumTopics.titulo, Usuarios.FullName, Usuarios.Email " + 
                    "FROM Usuarios INNER JOIN (ForumTopics INNER JOIN ForumMessages ON ForumTopics.ID = ForumMessages.Topic) ON Usuarios.ID = ForumMessages.[User] " + 
                    "WHERE (ForumMessages.Topic = " + TopicID + ") and (ForumTopics.Activo = " + Application("dtrue") + ")" ,oConn,3,3);
        
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
    <td class="ToolBar" align="middle"><h1><b>Listado de mensajes</b></h1></td>
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
            <a href="ShowForumTopics.asp?uid=<%=uid%>&Course=<%=CourseInUse%>&pn=0&courseName=<%=courseNameInUse%>">
              Todos los temas
            </a>  
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
            Lista de mensajes
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td> 
      <table width="100%" border="0" cellspacing="1" cellpadding="0">
        <tr>
          <td colspan=4 valign=bottom align=right>
            <table width="1%" border="0" cellspacing="1" cellpadding="0" align=right>
              <tr>
                <td width="1%" align=left>
                </td>
                <td width="1%" align=left valign=middle  nowrap>
                </td>
                <td width="1%" align=left>
                </td>
                <td width="1%" align=left>
                  <IMG src="../images/<%=Session("skin")%>/icon_send_topic.gif">
                </td>
                <td width="1%" align=left valign=middle  nowrap>
                  &nbsp;
                  <a href="javascript:abreVentana('Mensajeria', 600, 214, 'NewForumMsg.asp?uid=<%=uid%>&course=<%=CourseInUse%>&Topic=<%=TopicID%>&pn=<%=pagenumber%>&courseName=<%=courseNameInUse%>', 'yes')">
                    <font size=1>Nuevo Mensaje</font>
                  </a>
                </td>
                <td width="80%" align=left></td>
              </tr>
            </table>  
          </td>
        </tr>  
        <tr> 
          <td width="1" class="ToolBar" align="middle"></td>       
          <td width="1" class="ToolBar" align="middle"></td>       
          <td width="15%" class="ToolBar" align="middle"><b>Autor</b></td>
          <td width="83%" class="ToolBar" align="middle"><b>Texto</b></td>
        </tr>
        <%    
        var clase = "MessageTR1";

        creg = oRec.RecordCount;
      
        if (pagenumber * SHOW_CANT < creg){
          oRec.Move(pagenumber * SHOW_CANT); 
        }
        else{
          pagenumber = 0;
        }
      
        var i = 0;
        while ((oRec.EOF == false) && (i < SHOW_CANT)){ 
          i = i + 1;
          if (clase == "MessageTR1"){clase = "MessageTR";} else {clase = "MessageTR1";};

          var answerURL = "ShowForumAnswer.asp?uid=" + uid + "&Course=" + CourseInUse + "&Topic=" + TopicID + "&msgid=" + oRec.Fields.Item("ID").value + "&pn=0" + "&courseName=" + courseNameInUse;
        %>
        <tr> 
        
          <td width="1%" class="<%=clase%>" align="middle">
            <%if ((oRec.Fields.Item("AceptaRespuestas").value == true) && (oRec.Fields.Item("Activo").value == true)) {%>
            <a href="ShowForumMessages.asp?uid=<%=uid%>&Course=<%=CourseInUse%>&Topic=<%=TopicID%>&msgid=<%=oRec.Fields.Item("ID").value%>&pn=0&op=3&courseName=<%=courseNameInUse%>">
              <img title="Desactivar mensaje" src="../images/<%=Session("skin")%>/icon_folder.gif" border=0>
            </a>
            <%}
              else{%>  
            <a href="ShowForumMessages.asp?uid=<%=uid%>&Course=<%=CourseInUse%>&Topic=<%=TopicID%>&msgid=<%=oRec.Fields.Item("ID").value%>&pn=0&op=2&courseName=<%=courseNameInUse%>">
              <img title="Activar mensaje" src="../images/<%=Session("skin")%>/icon_folder_locked.gif" border=0>
            </a>
            <%}%>  
          </td>
        
        
          <td width="1%" class="<%=clase%>" align="middle">
            <%if ( (Session("PermissionType") == ADMINISTRATOR) || (Session("cordinador") = Session("userID")) || (Session("courseOwner") = Session("userID")) || (oRec.Fields.Item("User").value == Session("userID")) ){%>
            <a title="Eliminar mensaje" href="ShowForumMessages.asp?uid=<%=uid%>&Course=<%=CourseInUse%>&Topic=<%=TopicID%>&msgid=<%=oRec.Fields.Item("ID").value%>&pn=0&op=1&courseName=<%=courseNameInUse%>">
              <img src="../images/<%=Session("skin")%>/admintree/denied.gif" border=0>
            </a>  
            <%}%>
          </td>
          <td class="<%=clase%>" align="middle">
            <a href="<%=answerURL%>">
              <b><%=oRec.Fields.Item("Fullname").value%></b>
            </a>
          </td>
          <td class="<%=clase%>" align=left>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="1%">
                  <IMG src="../images/<%=Session("skin")%>/icon_posticon.gif">
                </td> 
                <td width="100%" align=left valign=bottom nowrap>
                  <font size=2>Enviado: <%=oRec.Fields.Item("Fecha").value%></font>
                </td>  
                <td width="1%" align=left>
                &nbsp;&nbsp;&nbsp;&nbsp;
                </td>
                <td width="1%" align=left>
                  <IMG src="../images/<%=Session("skin")%>/icon_profile.gif">
                </td>
                <td width="1%" align=left valign=middle  nowrap>
                  &nbsp;
                  <a href="javascript:abreVentana('Mensajería',600,192,'ShowForumAutorData.asp?uid=<%=uid%>&userid=<%=oRec.Fields.Item("User").value%>','yes')">
                    <font size=1>Autor</font>
                  </a>
                </td>
                <td width="1%" align=left>
                &nbsp;&nbsp;&nbsp;&nbsp;
                </td>
                <td width="1%" align=left>
                  <IMG src="../images/<%=Session("skin")%>/icon_email.gif">
                </td>
                <td width="1%" align=left valign=middle  nowrap>
                  &nbsp;
                  <a href="javascript:abreVentana('Mensajeria', 600, 252, 'NewForumEmail.asp?uid=<%=uid%>&course=<%=CourseInUse%>&UserTo=<%=oRec.Fields.Item("User").value%>&UserNameTo=<%=oRec.Fields.Item("Fullname").value%>&courseName=<%=courseNameInUse%>', 'yes')">
                    <font size=1>Email</font>
                  </a>
                </td>
                
                <td width="1%" align=left>
                &nbsp;&nbsp;&nbsp;&nbsp;
                </td>
                <td width="1%" align=left>
                  <IMG src="../images/<%=Session("skin")%>/icon_reply_topic.gif">
                </td>
                <td width="1%" align=left valign=middle  nowrap>
                  &nbsp;
                  <%if (oRec.Fields.Item("Aceptarespuestas").value == true){%>
                  <a href="<%=answerURL%>">
                  <%}%>
                  <font size=1>Opiniones</font>
                  <%if (oRec.Fields.Item("Aceptarespuestas").value == true){%>
                  </a>
                  <%}%>
                </td>
              </tr>
            </table>  
            <HR noShade SIZE=1>
            <%=oRec.Fields.Item("Texto").value%>
          </td>
        </tr> 
<%        oRec.Move(1);
        }  //While
        
        
%>      
        <tr>
          <td class="ToolBar" colspan=6>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td Width="10%" class="ToolBar" align=center>
                  <a href="ShowForumMessages.asp?uid=<%=uid%>&Course=<%=CourseInUse%>&Topic=<%=TopicID%>&pn=0&courseName=<%=courseNameInUse%>" class="ToolLink">&nbsp;Inicio&nbsp;</a>
                </td>
                <td Width="10%" class="ToolBar" align=center>
                  <a href="ShowForumMessages.asp?uid=<%=uid%>&Course=<%=CourseInUse%>&Topic=<%=TopicID%>&pn=<%=(pagenumber * 1) + 1%>&courseName=<%=courseNameInUse%>" class="ToolLink">&nbsp;Próximos&nbsp;<%=SHOW_CANT%>&nbsp;</a>
                </td>
                <td Width="99%" class="ToolBar">
                </td>
              </tr> 
              <tr>
                <td colspan=3  align=center>
                  <%
                    i = 0;
                    while (i * SHOW_CANT < creg){
                  %>    
                  <a class="numbers" href="ShowForumMessages.asp?uid=<%=uid%>&Course=<%=CourseInUse%>&Topic=<%=TopicID%>&pn=<%=i%>&courseName=<%=courseNameInUse%>">
                    <b><%=i + 1%>&nbsp;</b>
                  </a>  
                  <%
                      i = i + 1;
                    }
                  %>
                </td>
              </tr>
            </table>  
          </td>
        </tr>  
      </table>  
    </td>  
  </tr>
  <tr>
    <td colspan=4 valign=bottom align=right>
      <table width="1%" border="0" cellspacing="1" cellpadding="0" align=right>
        <tr>
          <td width="1%" align=left>
          </td>
          <td width="1%" align=left valign=middle  nowrap>
          </td>
          <td width="1%" align=left>
          </td>
          <td width="1%" align=left>
            <IMG src="../images/<%=Session("skin")%>/icon_send_topic.gif">
          </td>
          <td width="1%" align=left valign=middle  nowrap>
            &nbsp;
            <a href="javascript:abreVentana('Mensajeria', 600, 214, 'NewForumMsg.asp?uid=<%=uid%>&course=<%=CourseInUse%>&Topic=<%=TopicID%>&pn=<%=pagenumber%>&courseName=<%=courseNameInUse%>', 'yes')">
              <font size=1>Nuevo Mensaje</font>
            </a>
          </td>
          <td width="80%" align=left></td>
        </tr>
      </table>  
    </td>
  </tr>  
</table>  
</TR></TABLE>
</body>
</HTML>
<%
        }
        else{
          Response.Redirect("errorpage.asp?tipo=Error&short=Mensaje inactivo" + "&desc=El mensaje seleccionado no está activado para el uso" );                   
        }
        oRec.Close();
        oConn.Close();
      }    
      else{
        Response.Redirect("errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);                   
      }
  }  
  else
    Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>

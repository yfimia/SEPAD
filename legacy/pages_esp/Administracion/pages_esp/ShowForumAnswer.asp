<%@ Language=JScript %>
<%
/*
    Esta página muestra el listado de respuestas o comentarios enviados 
    referentes a un mensaje dado.
    
    PARAMETROS
    ----------
    
    uid        -- El ID que genera la sesión cuando el usuario se conecta 
                  por primera vez.
                  
    Course     -- ID del curso al que se le piden los temas de discusión.
    
    courseName -- Nombre del curso
    
    TopicID    -- ID del Tema seleccionado del cual se desean los mensajes
    
    msgid      -- ID del mensaje sobre el cual se pide el listado de las respuestas.
    
    pagenumber -- numero del bloque o pagina que se esta mostrando

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
        courseNameInUse = Request.QueryString.Item("courseName");
      }  

      var pagenumber = 0;
      if (Request.QueryString.Item("pn").Count != 0) {
        pagenumber = Request.QueryString.Item("pn");
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
          //Forum answer
        
          oRec.Open("SELECT ForumAnswers.ID,ForumAnswers.[User], ForumAnswers.Texto As Answer, ForumAnswers.fecha, ForumAnswers.Origen, ForumMessages.Texto, ForumTopics.Titulo, ForumTopics.Activo,Usuarios.FullName " + 
                    "FROM Usuarios INNER JOIN (ForumTopics INNER JOIN (ForumMessages INNER JOIN ForumAnswers ON ForumMessages.ID = ForumAnswers.Origen) ON ForumTopics.ID = ForumMessages.Topic) ON Usuarios.ID = ForumAnswers.[User] " + 
                    "WHERE (ForumAnswers.Origen = " + msgid + ") and (ForumTopics.Activo = " + Application("dtrue") + ")",oConn,3,3);

%>

<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<script src="../js/CheckBoxes.js" language="JavaScript"></script>
<script type="text/javascript" language="javascript" src="../js/Windows.js"></script>
</HEAD>

<body bgcolor="#ffffff" text="#000000">

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar" align="middle"><h1><b>Listado de respuestas</b></h1></td>
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
          <td align=left colspan=3>
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
            <IMG src="../images/<%=Session("skin")%>/icon_folder_open.gif" border=0>
          </td>
          <td align=left colspan=2>
            <a href="ShowForumMessages.asp?uid=<%=uid%>&Course=<%=CourseInUse%>&Topic=<%=TopicID%>&pn=0&courseName=<%=courseNameInUse%>">
              Lista de mensajes
            </a>  
          </td>
        </tr>
        <tr>      
          <td align=left width="1%">
          </td>
          <td align=left width="1%">
            <IMG src="../images/<%=Session("skin")%>/icon_bar.gif" border=0>
          </td>
          <td align=left width="1%">
            <IMG src="../images/<%=Session("skin")%>/icon_folder_open_topic.gif" border=0>
          </td>
          <td align=left>
            Respuestas del mensaje
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td> 
      <table width="100%" border="0" cellspacing="1" cellpadding="0">
        <tr>
          <td colspan=3 valign=bottom align=right>
            <table width="1%" border="0" cellspacing="0" cellpadding="0" align=right>
              <tr>
                <td width="1%" align=left>
                </td>
                <td width="1%" align=left valign=middle  nowrap>
                </td>
                <td width="1%" align=left>
                &nbsp;&nbsp;&nbsp;&nbsp;
                </td>
                <td width="1%" align=left>
                  <IMG src="../images/<%=Session("skin")%>/announcement_li.gif">
                </td>
                <td width="1%" align=left valign=middle  nowrap>
                  &nbsp;
                  <a href="javascript:abreVentana('Mensajería',600,192,'NewForumAnswer.asp?uid=<%=uid%>&course=<%=CourseInUse%>&Topic=<%=TopicID%>&msgid=<%=msgid%>&pn=<%=pagenumber%>&courseName=<%=courseNameInUse%>','yes')">
                    <font size=1>Nueva opinión</font>
                  </a>
                </td>
                <td width="80%" align=left></td>
              </tr>
            </table>  
          </td>
        </tr>  
        <tr> 
          <td width="15%" class="ToolBar" align="middle"><b>Autor</b></td>
          <td width="85%" class="ToolBar" align="middle"><b>Opinión o respuesta</b></td>
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

        %>
        <tr> 
          <td class="<%=clase%>" align="middle">
            <a title="Ver datos del autor" href="javascript:abreVentana('Mensajería',600,192,'ShowForumAutorData.asp?uid=<%=uid%>&userid=<%=oRec.Fields.Item("User").value%>','yes')">
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
              </tr>
            </table>  
            <HR noShade SIZE=1>
            <%=oRec.Fields.Item("Answer").value%>
          </td>
        </tr> 
<%        oRec.Move(1);
        }  //While
%>      
          <tr>
            <td class="ToolBar" colspan=5>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td Width="10%" class="ToolBar" align=center>
                  <a href="ShowForumAnswer.asp?uid=<%=uid%>&Course=<%=CourseInUse%>&Topic=<%=TopicID%>&msgid=<%=msgid%>&pn=0&courseName=<%=courseNameInUse%>" class="ToolLink">&nbsp;Inicio&nbsp;</a>
                </td>
                <td Width="10%" class="ToolBar" align=center>
                  <a href="ShowForumAnswer.asp?uid=<%=uid%>&Course=<%=CourseInUse%>&Topic=<%=TopicID%>&msgid=<%=msgid%>&pn=<%=pagenumber * 1 + 1%>&courseName=<%=courseNameInUse%>" class="ToolLink">&nbsp;Próximos&nbsp;<%=SHOW_CANT%>&nbsp;</a>
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
                  <a class="numbers" href="ShowForumAnswer.asp?uid=<%=uid%>&Course=<%=CourseInUse%>&Topic=<%=TopicID%>&msgid=<%=msgid%>&pn=<%=i%>&courseName=<%=courseNameInUse%>">
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
    <td colspan=3 valign=bottom align=right>
      <table width="1%" border="0" cellspacing="0" cellpadding="0" align=right>
        <tr>
          <td width="1%" align=left>
          </td>
          <td width="1%" align=left valign=middle  nowrap>
          </td>
          <td width="1%" align=left>
            &nbsp;&nbsp;&nbsp;&nbsp;
          </td>
          <td width="1%" align=left>
            <IMG src="../images/<%=Session("skin")%>/announcement_li.gif">
          </td>
          <td width="1%" align=left valign=middle  nowrap>
            &nbsp;
            <a href="javascript:abreVentana('Mensajería',600,192,'NewForumAnswer.asp?uid=<%=uid%>&course=<%=CourseInUse%>&Topic=<%=TopicID%>&msgid=<%=msgid%>&pn=<%=pagenumber%>&courseName=<%=courseNameInUse%>','yes')">
              <font size=1>Nueva opinión</font>
            </a>
          </td>
          <td width="80%" align=left></td>
        </tr>
      </table>  
    </td>
  </tr>  
</table>  
</TR></TBODY></TABLE>
</body>
</HTML>
<%            
        }
        else{
          Response.Redirect("errorpage.asp?tipo=Error&short=Mensaje inactivo" + "&desc=El mensaje seleccionado no está activado para el uso" );                   
        }
        oConn.Close();
      }    
      else {
        Response.Redirect("errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);                   
      }
    }  
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>

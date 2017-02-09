<%@ Language=JScript %>
<%
  /*
    Esta página muestra el listado de los temas de discusión que 
    pertenece al curso dado.
    
    PARAMETROS
    ----------
    
    uid        -- El ID que genera la sesión cuando el usuario se conecta 
                  por primera vez.
                  
    Course     -- ID del curso al que se le piden los temas de discusión.
    
    courseName -- Nombre del curso
    
    pagenumber -- numero del bloque o pagina que se esta mostrando
    
    op         -- Operacion extra a efectuar antes de cargar la pagina
                    Lista de operaciones
                    --------------------
                    0 - No operacion
   
                    1 - Eliminar tema
                          Topic - trae el ID del tema a borrar
              
                    2 - Activar tema
                          Topic - trae el ID del tema a poner activo
              
                    3 - Desactivar tema
                          Topic - trae el ID del tema a desactivar
              
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

  var  filePath = Application("dataPath");   
  var  oConn    = Server.CreateObject("ADODB.Connection");
  var  oRec     = Server.CreateObject("ADODB.Recordset");
  oConn.Open(filePath);

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

      if (Session("valid")) {
   
        operation = 0;
        if (Request.QueryString.Item("op").count != 0){
          operation = Request.QueryString.Item("op");
        }
  
        //Chequeo las operaciones
        
        var TopicId = -1;
        if (Request.QueryString.Item("Topic").Count != 0) {
          TopicId = Request.QueryString.Item("Topic")
        }  

        if (operation == 1){
          if ( (Session("PermissionType") == ADMINISTRATOR) || (Session("cordinador") = Session("userID")) || (Session("courseOwner") = Session("userID")) || (oRec.Fields.Item("ID").value == Session("userID")) ){
            oRec.Open("DELETE FROM ForumTopics WHERE ID = " + TopicId , oConn, 3, 3);    
          }  
        }
        
        if (operation == 2){
          if ( (Session("PermissionType") == ADMINISTRATOR) || (Session("cordinador") = Session("userID")) || (Session("courseOwner") = Session("userID")) || (oRec.Fields.Item("ID").value == Session("userID")) ){
            oRec.Open("UPDATE ForumTopics SET Activo = " + Application("dtrue") + " WHERE ID = " + TopicId , oConn, 3, 3);
          }  
        }
        
        if (operation == 3){
          if ( (Session("PermissionType") == ADMINISTRATOR) || (Session("cordinador") = Session("userID")) || (Session("courseOwner") = Session("userID")) || (oRec.Fields.Item("ID").value == Session("userID")) ){
            oRec.Open("UPDATE ForumTopics SET Activo = " + Application("dfalse") + " WHERE ID = " + TopicId , oConn, 3, 3);
          }      
        }
        
        
        //Forum Topics
        oRec.Open("SELECT ForumTopics.ID AS FID, ForumTopics.Curso, ForumTopics.Titulo, ForumTopics.Descripcion AS DForum, ForumTopics.Fecha, ForumTopics.Activo, Usuarios.FullName, Usuarios.ID " + 
                  "FROM Usuarios INNER JOIN ForumTopics ON Usuarios.ID = ForumTopics.[User] " + 
                  "WHERE (ForumTopics.curso = " + CourseInUse + ")", oConn, 3, 3);

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
    <td class="ToolBar" align="middle"><h1><b>Listado de temas</b></h1></td>
  </tr>
</table> 

<table width="98%" border="0" cellspacing="0" cellpadding="0" align=center>
  <tr>
    <td>
      <table border="0" cellspacing="3" cellpadding="0">
        <tr>      
          <td align=left width="1%">
            <IMG src="../images/<%=Session("skin")%>/icon_folder_open_topic.gif" border=0>
          </td>
          <td align=left width="99%">
            Todos los temas
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td align=right colspan=4>
      <table width="1%" border="0" cellspacing="0" cellpadding="0" align=right>
        <tr>
          <td width="1%" align=left>
            <a href="javascript:abreVentana('Mensajeria', 600, 214, 'NewForumTopic.asp?uid=<%=uid%>&course=<%=CourseInUse%>&courseName=<%=courseNameInUse%>', 'yes')">             
              <IMG border=0 src="../images/<%=Session("skin")%>/icon_folder_new_topic.gif">
            </a>  
          </td>
          <td width="1%" align=left valign=middle nowrap>
            &nbsp;
            <a href="javascript:abreVentana('Mensajeria', 600, 234, 'NewForumTopic.asp?uid=<%=uid%>&course=<%=CourseInUse%>&pn=<%=pagenumber%>&courseName=<%=courseNameInUse%>', 'yes')">             
              <font size=1>Nuevo tema</font>
            </a>  
          </td>
        </tr>
      </table>  
    </td>
  </tr>
  <tr> 
    <td> 
      <table width="100%" border="0" cellspacing="1" cellpadding="0">
        <tr> 
          <td width="1%" class="ToolBar" align="middle"></td>
          <td width="1%" class="ToolBar" align="middle"></td>
          <td width="34%" class="ToolBar" align="middle"><b>Tema</b></td>
          <td width="34%" class="ToolBar" align="middle"><b>Descripción</b></td>
          <td width="15%" class="ToolBar" align="middle"><b>Autor</b></td>
          <td width="15%" class="ToolBar" align="middle"><b>Fecha</b></td>
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
          var topicURL = "ShowForumMessages.asp?uid=" + uid + "&Course=" + CourseInUse + "&Topic=" + oRec.Fields.Item("FID").value + "&pn=0&courseName=" + courseNameInUse;
%>
        <tr> 
          <td width="1%" class="<%=clase%>" align="middle">
            <%if (oRec.Fields.Item("Activo").value == true){%>
            <a href="ShowForumTopics.asp?uid=<%=uid%>&Course=<%=CourseInUse%>&Topic=<%=oRec.Fields.Item("FID").value%>&pn=0&op=3&courseName=<%=courseNameInUse%>">
              <img title="Desactivar tema" src="../images/<%=Session("skin")%>/icon_folder.gif" border=0>
            </a>
            <%}
              else{%>  
            <a href="ShowForumTopics.asp?uid=<%=uid%>&Course=<%=CourseInUse%>&Topic=<%=oRec.Fields.Item("FID").value%>&pn=0&op=2&courseName=<%=courseNameInUse%>">
              <img title="Activar tema" src="../images/<%=Session("skin")%>/icon_folder_locked.gif" border=0>
            </a>
            <%}%>  
          </td>
          <td width="1%" class="<%=clase%>" align="middle">
            <%if ( (Session("PermissionType") == ADMINISTRATOR) || (Session("cordinador") = Session("userID")) || (Session("courseOwner") = Session("userID")) || (oRec.Fields.Item("ID").value == Session("userID")) ){%>
            <a title="Eliminar tema" href="ShowForumTopics.asp?uid=<%=uid%>&Course=<%=CourseInUse%>&Topic=<%=oRec.Fields.Item("FID").value%>&pn=0&op=1&courseName=<%=courseNameInUse%>">
              <img src="../images/<%=Session("skin")%>/admintree/denied.gif" border=0>
            </a>  
            <%}%>
          </td>
          <td class="<%=clase%>" align="middle">
            <%if (oRec.Fields.Item("Activo").value == true){%>
            <a href="<%=topicURL%>">
            <%}%>            
              <b><%=oRec.Fields.Item("titulo").value%></b>
            <%if (oRec.Fields.Item("Activo").value == true){%>              
            </a>
            <%}%>            
          </td>
          <td class="<%=clase%>" align="middle"><%=oRec.Fields.Item("DForum").value%></td>
          <td class="<%=clase%>" align="middle"><%=oRec.Fields.Item("Fullname").value%></td>
          <td class="<%=clase%>" align="middle"><%=oRec.Fields.Item("Fecha").value%></td>
        </tr> 
<%          
          oRec.Move(1);
        }//While
        
%>      
        <tr>
          <td class="ToolBar" colspan=6>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td Width="10%" class="ToolBar" align=center>
                  <a href="ShowForumTopics.asp?uid=<%=uid%>&Course=<%=CourseInUse%>&pn=0&courseName=<%=courseNameInUse%>" class="ToolLink">&nbsp;Inicio&nbsp;</a>
                </td>
                <td Width="10%" class="ToolBar" align=center>
                  <a href="ShowForumTopics.asp?uid=<%=uid%>&Course=<%=CourseInUse%>&pn=<%=((pagenumber * 1) + 1)%>&courseName=<%=courseNameInUse%>" class="ToolLink">&nbsp;Próximos&nbsp;<%=SHOW_CANT%>&nbsp;</a>
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
                  <a class="numbers" href="ShowForumTopics.asp?uid=<%=uid%>&Course=<%=CourseInUse%>&pn=<%=i%>&courseName=<%=courseNameInUse%>">
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
</TABLE>  
</body>
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

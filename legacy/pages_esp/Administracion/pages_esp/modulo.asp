<%@ Language=JavaScript %>

<!-- #include file='../js/adolibrary.inc' -->
<!-- #include file='../js/user.inc' -->
<!-- #include file='../js/course.inc' -->
<!-- #include file='../js/change.inc' -->

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

 Session("modulo")        = Request.QueryString("ID") + "";
 Session("moduloName")    = Request.QueryString("Name") + "";
 Session("state")        = Request.QueryString("state") + "";
 Session("cordinador")    = Request.QueryString("cordinador") + "";
 Session("grupo")        = Request.QueryString("grupo") + "";
 Session("claustro")    = Request.QueryString("claustro") + "";
 
 var state = Request.QueryString("state") + "";
 var cordinador = Request.QueryString("cordinador") + "";
 var grupo = Request.QueryString("grupo") + "";
 var claustro = Request.QueryString("claustro") + "";

%>



<%
 //Chequeos de los derechos del usuario actual respecto a este curso
 var isowner = ((Session("UserID") == cordinador) || (Session("permissionType") == ADMINISTRATOR));
 var valid = ( isowner || (isUserInGroupByID(Session("UserID"), grupo)) || (isUserInGroupByID(Session("UserID"), claustro)));
 Session("valid") = valid;
 
%>
<html>
<head>
<title>Generalidades académicas de <%=Session("moduloName")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">

<script src="../js/windows.js" language="JavaScript"></script>
<script language=JavaScript>
    function openwin(uid){
      window.open("grupo.asp?uid=" + uid,null,"height=200,width=340,toolbar=no,status=no,directories=no,menubar=no,scrollbars=yes,resizable=yes");
    }
</script>

</head>


<body bgcolor="#FFFFFF" text="#000000" style="overflow-y:scroll">
    <%
      var  filePath = Application("dataPath");
      var  oConn    = Server.CreateObject("ADODB.Connection");
      var  oRec     = Server.CreateObject("ADODB.Recordset");
      oConn.Open(filePath);
      sql = "SELECT Name FROM grupos where (ID = " + Session("grupo") + ")";
      
      
      oRec.Open(sql,oConn,3,3);
      var Name = oRec.Fields.Item('Name').value;  
      
       oRec.Close();
       oConn.Close();
     %>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar"> 
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td width="84%">
          </td>

<%
  if (Session("userID")==GUEST_USER) {  
       
%>  
          <td width="1%" ><a href='javascript:<%=MATRICULAS_PAGE%>' class="ToolLink">&nbsp;Solicitar&nbsp;Matrícula&nbsp;</a></td>
<% 
  
  
  }
  else {
  //Si no el el usuario invitado
  //Si no es matricula y el curso esta en incripcion
  if ((!valid) && ((state == MOD_ACA_ININSCRIPTION) || (state == MOD_ACA_INCOURINSC))) {       
%>  
          <td width="1%" ><a href="javascript:abreVentana('Matricular',250,200,'matriculatemp.asp?uid=<%=uid%>&modulo=<%=Session("modulo")%>&courseName=<%=Session("moduloName")%>', 'no', 'no')" class="ToolLink">&nbsp;Solicitar&nbsp;Matrícula&nbsp;</a></td>
            <td class="Separador" width="1%">|</td>
<% 
   }
%>
            <td width="1%"><a href="javascript:abreVentana('Anuncios', 480, 300, 'Anuncios.asp?uid=<%=uid%>&idmodulo=<%=Session("modulo")%>')" class="ToolLink">Anuncios</td>
<% 
   }
  
  if (isowner) {       
    if ((state == MOD_ACA_ININSCRIPTION) || (state == MOD_ACA_INCOURINSC)) {       
%>      
          <td class="Separador" width="1%">|</td>
<%  
    }
%>
          <td width="1%"><a href="javascript:abreVentana('ModalidadAdm',700,400,'ModAdm/main.asp?uid=<%=uid%>&modulo=<%=Session("modulo")%>&courseName=<%=Session("moduloName")%>', 'no', 'no')" class="ToolLink">&nbsp;Administrar&nbsp;</a></td>
<%
  }
%>
        <td width="10%">
        </td>

        </tr>
      </table>
    </td>
  </tr>
  <tr>


 <tr>
 <td>
 <table width="100%" border="0" cellspacing="5" cellpadding="0" class="CouseTable">
  <tr>
   <td width="73%">
    <table width="90%" border="0" cellspacing="3" cellpadding="0" class="BorderedTable" align="center">
     <tr>
      <td align="center">
       <h2>Modalidad Académica: <%=Session("moduloName")%></h2> <h6>Código: <%=Session("modulo")%></h6>
      </td>
      </tr>
      <tr>
       <td >
<%
  var corduser = GetUserData(cordinador);

          var dataPath = Application('dataPath');
          var oConn;
          var oRec;
          
          oConn = MakeConnection( oConn, dataPath );

          Sql = "SELECT Usuarios.* " + 
                "FROM Usuarios INNER JOIN (Grupos INNER JOIN Grupos_de_Usuarios ON Grupos.ID = Grupos_de_Usuarios.[Group]) ON Usuarios.ID = Grupos_de_Usuarios.[User] " +
                "WHERE (((Grupos.ID)= " + claustro + ") and (Usuarios.ID <> " + corduser.id + " ))";
          
          oRec = Query( Sql, oRec, oConn  );		
%>        
         <table width="100%" border="0" cellspacing="1" cellpadding="0" class="PrologueTable">
           <tr> 
             <td colspan=4 width="100%" class="ToolBar" align="center"></td>
           </tr>
         
           <tr> 
             <td width="5%" class="ToolBar" align="center"><b>Claustro:</b></td>
             <td width="20%" class="ToolBar" align="center"><b>Identificador</b></td>
             <td width="40%" class="ToolBar" align="center"><b>Nombre y apellidos</b></td>
             <td width="35%" class="ToolBar" align="center"><b>Correo eléctronico</b></td>
           </tr>

          <tr> 
            <td width="5%" class="MessageTR1">&nbsp;<b>Coordinador</b>&nbsp;</td>
            <td width="20%" class="MessageTR1">&nbsp;<a href="javascript:abreVentana('',350,280,'userDetails.asp?uid=<%=uid%>&id=<%=corduser.id%>', 'no', 'no')"><%=corduser.name%></a>&nbsp;</td>
            <td width="40%" class="MessageTR1">&nbsp;<%=corduser.fullName%>&nbsp;</td>
            <td width="35%" class="MessageTR1">&nbsp;<a href="mailto:<%=corduser.email%>" ><%=corduser.email%></a>&nbsp;</td>
          </tr> 



<%
          var clase = "MessageTR1";
          
          var i = 0;
          var last = "-1"
          while (oRec.EOF == false)
            { 
              i = i + 1;
              if (clase == "MessageTR1") clase = "MessageTR"; else clase = "MessageTR1";
              last = oRec.Fields.Item("Name").value;          
%>
          <tr> 
            <td width="5%" class="<%=clase%>"> </td>
            <td width="20%" class="<%=clase%>">&nbsp;<a href="javascript:abreVentana('',350,280,'userDetails.asp?uid=<%=uid%>&id=<%=oRec.Fields.Item("id").value%>', 'yes', 'yes')"><%=oRec.Fields.Item("name").value%></a>&nbsp;</td>
            <td width="40%" class="<%=clase%>">&nbsp;<%=oRec.Fields.Item("FullName").value%>&nbsp;</td>
            <td width="35%" class="<%=clase%>">&nbsp;<a href="mailto:<%=oRec.Fields.Item("email").value%>" ><%=oRec.Fields.Item("email").value%></a>&nbsp;</td>
          </tr> 
  
<%
           oRec.Move(1);
         }  
         oRec.Close;        
        if (clase == "MessageTR1") clase = "MessageTR"; else clase = "MessageTR1";
%>        

          <tr> 
            <td colspan="4" class="ToolBar"></td>
	  <tr>             
          <tr> 
            <td width="1%" class="<%=clase%>"><b>Grupo&nbsp;de&nbsp;alumnos:&nbsp;</b></td>      
            <td colspan="3" class="<%=clase%>">&nbsp;<a href="javascript:openwin(<%=uid%>)" ><%=Name%></a>&nbsp;</td>      
          </tr> 
         </table>
       </td>
     </tr>
    </table>
   </td>
   <td width="27%" height="70"><img id="logo" src="../images/<%=Session("skin")%>/logoSepad.gif" width="150" height="109" />
   <h5>Estado:
   <%  
      switch (parseInt(Session("state"))) {
       case MOD_ACA_NOVISIBLE: {%>No visible<% break;}
       case MOD_ACA_DISABLE  : {%>Deshabilitado<% break;}
       case MOD_ACA_ININSCRIPTION : {%>En inscripción<% break;}
       case MOD_ACA_INCOURINSC : {%>En inscripción en curso<% break;}
       case MOD_ACA_INCOURSE : {%>En curso<% break;}
       case MOD_ACA_FINISHED : {%>Finalizado<% break;}
       default : {%>Deshabilitado<% break;}
      }       
     %>        
   </h5>
   </td>
  </tr>
  <tr>
  <td colspan="2">
    
  </td>
  </tr>
   
  
     
   
   
</table>

<%
  
  var courses = GetModuloCourses( Session("modulo") );
  
  var i;
  if (courses.length > 0) {
%>
<table width="90%" cellpadding="2" cellspacing="1" align="center" class="PrologueTable">
  <tr> 
    <td colspan=8 width="100%" class="ToolBar" align="center"></td>
  </tr>

  <tr >
    <td colspan=8 align=center class="ToolBar"><b>Modulos</b>
    </td>
  </tr>
        
           <tr> 
             <td width="1%" class="ToolBar" align="center"></td>
             <td width="4%" class="ToolBar" align="center"><b>Código</b></td>
             <td width="48%" class="ToolBar" align="center"><b>Nombre</b></td>
             <td width="10" class="ToolBar" align="center"><b>Estado</b></td>
             <td width="36%" colspan=4 class="ToolBar" align="center"></td>
             
           </tr>

<%
  for (i = 0; i < courses.length; i++) {
    if ((courses[i].state != MOD_ACA_NOVISIBLE) || (isowner)) { 
%>  	
<tr>
        <td class="<%=clase%>" width="1%"><IMG SRC="../images/<%=Session("skin")%>/TreeViewImg.gif"   class="IconClass" /></td>
        <td class="<%=clase%>" width="4%"><%=courses[i].ID%></td>
        
        <td class="<%=clase%>"  width="48%"><%=courses[i].Name%> </td>
        <td width="10%" nowrap class="<%=clase%>" >
<%  
      switch (courses[i].state) {
       case MOD_ACA_NOVISIBLE: {%>No visible<% break;}
       case MOD_ACA_DISABLE  : {%>Deshabilitado<% break;}
       case MOD_ACA_INCOURSE : {%>En curso<% break;}
       case MOD_ACA_FINISHED : {%>Finalizado<% break;}
       default : {%>Deshabilitado<% break;}
      }       
%>        
        </td>
        
        <td class="ToolBar"  width="1%"><a href="Course.asp?uid=<%=uid%>&ID=<%=courses[i].ID%>&Name=<%=courses[i].Name%>">
           <img title="Generalidades del módulo" src="../images/<%=Session("skin")%>/Generalidades.gif"  width="20" height="20" border="0" ></a>   
        </td>
        
        <td nowrap  class="ToolBar" width="1%">
<%
  if ((courses[i].owner == Session("userID")) || (isowner) || (valid && (courses[i].state == MOD_ACA_INCOURSE) && ((state == MOD_ACA_INCOURSE) || (state == MOD_ACA_INCOURINSC)))) {
    if (!empty("select id from Lecciones where course = " + courses[i].ID)) {   
%>        
        
        <a href="lessonBrowser.asp?uid=<%=uid%>&init=0&ID=<%=courses[i].ID%>&Name=<%=courses[i].Name%>">
           <img title="Aula Virtual" src="../images/<%=Session("skin")%>/avirtual.gif"  width="20" height="20" border="0" ></a>   
<%
    }
  }
%>           
           
        </td>
        <td  class="ToolBar" width="1%">
<%
  if ( isowner || (courses[i].owner == Session("userID"))) {
%>        
        
        <a  href="javascript:abreVentana('CourseAdm',700,400,'moduloAdm/main.asp?uid=<%=uid%>&course=<%=courses[i].ID%>&courseName=<%=courses[i].Name%>&courseOwner=<%=courses[i].owner%>', 'no', 'no')" ><img title="Administrar" src="../images/<%=Session("skin")%>/admin.gif"  width="20" height="20" border="0" ></a>
<%
  }
%>           

        </td>
        
      </tr>
  </td>
</tr> 
<%
   }      	
  }

%>
</table>
<%
 }
%>


<%
  Sql = "SELECT * " + 
         "FROM Modulo " +
         "WHERE (ID= " + Session("modulo") + ")";
   
   oRec = Query( Sql, oRec, oConn  );		


%>
<br>
<%
  clase = "";
  if ((oRec.Fields.Item("diploma").value != " ") && (oRec.Fields.Item("org").value != null)) {
%>
<table width="90%" cellpadding="2" cellspacing="1" align="center" class="PrologueTable">
<tr >
  <td colspan=6 class="ToolBar"><b>Certificado:</b>
  </td>
</tr>
<tr>
  <td class="<%=clase%>" >
     <%= change(oRec.Fields.Item("diploma").value) %>
  </td>
</tr>
</table>
<br>
<%}
  if ((oRec.Fields.Item("Fundament").value != " ") && (oRec.Fields.Item("org").value != null)) {
%>
<table width="90%" cellpadding="2" cellspacing="1" align="center" class="PrologueTable">
<tr >
  <td colspan=6 class="ToolBar"><b>Fundamentación:</b>
  </td>
</tr>
<tr>
  <td class="<%=clase%>" >
     <%= change(oRec.Fields.Item("Fundament").value) %>
  </td>
</tr>
</table>
<br>

<%}
  if ((oRec.Fields.Item("Sede").value != " ") && (oRec.Fields.Item("org").value != null)) {
%>
<table width="90%" cellpadding="2" cellspacing="1" align="center" class="PrologueTable">
<tr >
  <td colspan=6 class="ToolBar"><b>Sede:</b>
  </td>
</tr>
<tr>
  <td class="<%=clase%>" >
     <%= change(oRec.Fields.Item("Sede").value) %>
  </td>
</tr>
</table>
<br>
<%}
  if ((oRec.Fields.Item("fprevia").value != " ") && (oRec.Fields.Item("org").value != null)) {
%>
<table width="90%" cellpadding="2" cellspacing="1" align="center" class="PrologueTable">
<tr >
  <td colspan=6 class="ToolBar"><b>Perfil del Egresado:</b>
  </td>
</tr>
<tr>
  <td class="<%=clase%>" >
     <%= change(oRec.Fields.Item("fprevia").value) %>
  </td>
</tr>
</table>
<br>
<%}
  if ((oRec.Fields.Item("fechas").value != " ") && (oRec.Fields.Item("org").value != null)) {
%>
<table width="90%" cellpadding="2" cellspacing="1" align="center" class="PrologueTable">
<tr >
  <td colspan=6 class="ToolBar"><b>Fechas y Duración:</b>
  </td>
</tr>
<tr>
  <td class="<%=clase%>" >
     <%= change(oRec.Fields.Item("fechas").value) %>
  </td>
</tr>
</table>
<br>
<%}
  if ((oRec.Fields.Item("objetivos").value != " ") && (oRec.Fields.Item("org").value != null)) {
%>
<table width="90%" cellpadding="2" cellspacing="1" align="center" class="PrologueTable">
<tr >
  <td colspan=6 class="ToolBar"><b>Objetivos Generales:</b>
  </td>
</tr>
<tr>
  <td class="<%=clase%>" >
     <%= change(oRec.Fields.Item("objetivos").value) %>
  </td>
</tr>
</table>
<br>

<%}
  if ((oRec.Fields.Item("programa").value != " ") && (oRec.Fields.Item("org").value != null)) {
%>
<table width="90%" cellpadding="2" cellspacing="1" align="center" class="PrologueTable">
<tr >
  <td colspan=6 class="ToolBar"><b>Programa:</b>
  </td>
</tr>
<tr>
  <td class="<%=clase%>" >
     <%= change(oRec.Fields.Item("programa").value) %>
  </td>
</tr>
</table>
<br>
<%}
  if ((oRec.Fields.Item("org").value != " ") && (oRec.Fields.Item("org").value != null)) {
%>
<table width="90%" cellpadding="2" cellspacing="1" align="center" class="PrologueTable">
<tr >
  <td colspan=6 class="ToolBar"><b>Sistema de Evaluación:</b>
  </td>
</tr>
<tr>
  <td class="<%=clase%>" >
     <%= change(oRec.Fields.Item("org").value) %>
  </td>
</tr>
</table>
<br>

<%}
  if ((oRec.Fields.Item("app").value != " ") && (oRec.Fields.Item("org").value != null)) {
%>
<table width="90%" cellpadding="2" cellspacing="1" align="center" class="PrologueTable">
<tr >
  <td colspan=6 class="ToolBar"><b>Sistema de Tutorías:</b>
  </td>
</tr>
<tr>
  <td class="<%=clase%>" >
     <%= change(oRec.Fields.Item("app").value) %>
  </td>
</tr>
</table>
<br>

<%}
  if ((oRec.Fields.Item("Bibliog").value != " ") && (oRec.Fields.Item("org").value != null)) {
%>
<table width="90%" cellpadding="2" cellspacing="1" align="center" class="PrologueTable">
<tr >
  <td colspan=6 class="ToolBar"><b>Bibliografía General:</b>
  </td>
</tr>
<tr>
  <td class="<%=clase%>" >
     <%= change(oRec.Fields.Item("Bibliog").value) %>
  </td>
</tr>
</table>
<br>

<%}
  if ((oRec.Fields.Item("doc").value != " ") && (oRec.Fields.Item("org").value != null)) {
%>
<table width="90%" cellpadding="2" cellspacing="1" align="center" class="PrologueTable">
<tr >
  <td colspan=6 class="ToolBar"><b>Recursos:</b>
  </td>
</tr>
<tr>
  <td class="<%=clase%>" >
     <%=change(oRec.Fields.Item("doc").value)%>
  </td>
</tr>
</table>
<br>
<%}
  if ((oRec.Fields.Item("coste").value != " ") && (oRec.Fields.Item("org").value != null)) {
%>
<table width="90%" cellpadding="2" cellspacing="1" align="center" class="PrologueTable">
<tr >
  <td colspan=6 class="ToolBar"><b>Coste y forma de pago:</b>
  </td>
</tr>
<tr>
  <td class="<%=clase%>" >
     <%= change(oRec.Fields.Item("coste").value) %>
  </td>
</tr>
</table>
<br>

<%
  }
  if ((oRec.Fields.Item("inst").value != " ") && (oRec.Fields.Item("org").value != null)) {
%>
<table width="90%" cellpadding="2" cellspacing="1" align="center" class="PrologueTable">
<tr >
  <td colspan=6 class="ToolBar"><b>Datos de la institución:</b>
  </td>
</tr>
<tr>
  <td class="<%=clase%>" >
     <%= change(oRec.Fields.Item("inst").value) %>
  </td>
</tr>
</table>
<%
 }
%>


</html>
<%
  DestroyAdoObjects( oConn, oRec );        
 }
 else
   {Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT)};
%>






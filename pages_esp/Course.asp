<%@ Language=JavaScript %>
<%
Response.Expires = -1;

%>
<!-- #include file='../js/adolibrary.inc' -->
<!-- #include file='../js/course.inc' -->
<!-- #include file='../js/user.inc' -->
<!-- #include file='inc/course.inc' -->

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

 Session("course")        = Request.QueryString("ID") + "";
 Session("courseName")    = Request.QueryString("Name") + "";
 Session("FullCourse")    = GetFullCourse(Session("course"));
 fullCourse = Session("FullCourse");

%>



<%
      var sXml = "../Courses/course" + Session("course") + "/prologo.xml";
      try {
	var sXml = "../Courses/course" + Session("course") + "/" + fullCourse.prologue;
      }
      catch(e) {		
        
      }	
	var sXsl = "../Xsl/prologo.xsl";

		
	var oXmlDoc = Server.CreateObject("MICROSOFT.XMLDOM");
	var oXslDoc = Server.CreateObject("MICROSOFT.XMLDOM");
	
	oXmlDoc.async = false;
	oXslDoc.async = false;
	
	oXmlDoc.load(Server.MapPath(sXml));
	oXslDoc.load(Server.MapPath(sXsl));
  
%>
<html>
<head>
<title><%=title%><%=Session("courseName")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">

<script src="../js/windows.js" language="JavaScript"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
function inicioClick()
 {//iniciar un nuevo curso......
    
     st	      = "?uid=<%=uid%>&init=1";
     //alert(st);
     location.href = "lessonBrowser.asp" + st;
 }
 
function contClick()
 {
   st = "?uid=<%=uid%>&init=0&ID=<%=Session("course")%>&Name=<%=Session("courseName")%>";
   location.href = "lessonBrowser.asp" + st; 
 }
 
 
//-->
</SCRIPT>
</head>

<body bgcolor="#FFFFFF" text="#000000"  style="overflow:auto">
<table width="100%" border="0" cellspacing="0" cellpadding="0">

<%
   //Si no es un usuario con derechos para ver el curso
    //Si es el usuario invitado que se matricule primero...
    if (Session("UserID") != GUEST_USER)
     {
%>
  <tr> 
    <td class="ToolBar"> 
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td  width="85%" ></td>        
<%        
    if ((Session("userID") == fullCourse.owner) || Session("valid"))  { 
      if ((fullCourse.state == MOD_ACA_INCOURSE) && ((Session("state") == MOD_ACA_INCOURSE) || (Session("state") == MOD_ACA_INCOURINSC))) {    
         if (!empty("select id from Lecciones where course = " + Session("course"))) {   
      
%>        
          <td  width="1%" ><a href="javascript:contClick()" class="ToolLink">&nbsp;<%=stmp2%>&nbsp;</a></td>
<%
         }
%>          
<%    } 
       if (
           (Session("userID") == fullCourse.owner) || 
           (Session("userID") == Session("cordinador")) || 
           (Session("PermissionType") == ADMINISTRATOR) || 
           (
            isUserInGroupByID(Session("userID"), Session("claustro"))  
            && (
                 (Session("state") == MOD_ACA_INCOURSE) || 
                 (Session("state") == MOD_ACA_INCOURINSC)
                )
           )
          ) { //si es tutor
%>              
          <!--td  width="1%" >|</td>
          <td  width="1%" ><a href="javascript:abreVentana(<%=stmp3%>,700,400,'Glosario/glosario.asp?uid=<%=uid%>&IdCurso=<%=Session("course")%>&Letra=A&IDP=0&courseName=<%=Session("courseName")%>', 'yes', 'yes')" class="ToolLink">&nbsp;Glosario&nbsp;</a></td-->
          
          <td  width="1%" >|</td>
          <td  width="1%" ><a href="javascript:abreVentana(<%=stmp3%>,700,400,'moduloTutor/main.asp?uid=<%=uid%>&course=<%=Session("course")%>&courseName=<%=Session("courseName")%>&courseOwner=<%=fullCourse.owner%>', 'no', 'no')" class="ToolLink">&nbsp;<%=stmp4%>&nbsp;</a></td>
<%     }
       if ((Session("userID") == fullCourse.owner) || (Session("userID") == Session("cordinador")) || (Session("PermissionType") == ADMINISTRATOR) )  { //si es profesor 
%>          
          <td  width="1%" >|</td>
          
          
          <td  width="1%" ><a href="javascript:abreVentana(<%=stmp5%>,700,400,'moduloAdm/main.asp?uid=<%=uid%>&course=<%=Session("course")%>&courseName=<%=Session("courseName")%>&courseOwner=<%=fullCourse.owner%>&cordinador=<%=Session("cordinador")%>&modulo=<%=Session("modulo")%>&moduloName=<%=Session("moduloName")%>&state=<%=Session("state")%>&grupo=<%=Session("grupo")%>&claustro=<%=Session("claustro")%>', 'no', 'no')" class="ToolLink">&nbsp;<%=stmp6%>&nbsp;</a></td>
<%
       }
    }
%>          
          <td  width="10%" ></td>

        </tr>
      </table>
    </td>
  </tr>
<%    
    
  }   
%> 


 <tr>
 <td>
 <table width="100%" border="0" cellspacing="5" cellpadding="0" class="CouseTable">
  <tr>
   <td width="73%">
    <table width="90%" border="0" cellspacing="3" cellpadding="0" class="BorderedTable" align="center">
     <tr>
      <td align="center">
       <h2><%=stmp11%><%=Session("courseName")%></h2> <h6><%=stmp12%><%=Session("course")%></h6>
      </td>
      </tr>
      <tr>
       <td >
<%
          var dataPath = Application('dataPath');
          var oConn;
          var oRec;
          
          oConn = MakeConnection( oConn, dataPath );

          Sql = "SELECT distinct Usuarios.* " +
                "FROM Usuarios INNER JOIN SubGrupo ON Usuarios.ID = SubGrupo.Tutor " +
                "WHERE (((SubGrupo.Curso)="  + Session("course") + ")) order by Usuarios.Name";
          
          oRec = Query( Sql, oRec, oConn  );		
%>        
         <table width="100%" border="0" cellspacing="1" cellpadding="0" class="PrologueTable">
           <tr> 
             <td colspan=4 width="100%" class="ToolBar" align="center"></td>
           </tr>


           <tr> 
             <td width="5%" class="ToolBar" align="center"><b><%=stmp13%></b></td>
             <td width="20%" class="ToolBar" align="center"><b><%=stmp14%></b></td>
             <td width="40%" class="ToolBar" align="center"><b><%=stmp15%></b></td>
             <td width="35%" class="ToolBar" align="center"><b><%=stmp16%></b></td>
           </tr>

<%
  var corduser = GetUserData(fullCourse.owner);
%>   
          <tr> 
            <td width="5%" nowrap align=left class="MessageTR1">&nbsp;<b><%=stmp17%></b>&nbsp;</td>
            <td width="20%" class="MessageTR1">&nbsp;<a href="javascript:abreVentana('',350,280,'userDetails.asp?uid=<%=uid%>&id=<%=corduser.id%>', 'yes', 'yes')"><%=corduser.name%></a>&nbsp;</td>
            <td width="40%" class="MessageTR1">&nbsp;<%=corduser.fullName%>&nbsp; </td>
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
            <td width="5%" align=left class="<%=clase%>">&nbsp;<%=stmp18%>&nbsp;</td>
            <td width="20%" class="<%=clase%>">&nbsp;<a href="javascript:abreVentana('',350,280,'userDetails.asp?uid=<%=uid%>&id=<%=oRec.Fields.Item("id").value%>', 'yes', 'yes')"><%=oRec.Fields.Item("name").value%></a>&nbsp;</td>
            <td width="40%" class="<%=clase%>">&nbsp;<%=oRec.Fields.Item("FullName").value%>&nbsp;</td>
            <td width="35%" class="<%=clase%>">&nbsp;<a href="mailto:<%=oRec.Fields.Item("email").value%>" ><%=oRec.Fields.Item("email").value%></a>&nbsp;</td>
          </tr> 
  
<%
           oRec.Move(1);
         }  
         oRec.Close;        
%>        

         </table>
       </td>
     </tr>
    </table>
   </td>
   <td width="27%" height="70"><img id="logo" src="../images/<%=Session("skin")%>/logoSepad.gif" width="150" height="109" />
   <h5><%=stmp19%> 
   <%  
      switch (parseInt(fullCourse.state)) {
       case MOD_ACA_NOVISIBLE: {%><%=stmp20%><% break;}
       case MOD_ACA_DISABLE  : {%><%=stmp21%><% break;}
       case MOD_ACA_INCOURSE : {%><%=stmp22%><% break;}
       case MOD_ACA_FINISHED : {%><%=stmp23%><% break;}
       default : {%><%=stmp24%><% break;}
      }       
     %>        
   </h5>
   </td>
  </tr>
  <tr>
  <td colspan="2">
    
  </td>
  </tr>  
  
  <tr>
<%
    //Aqui va el prologo... 
    Response.Write(oXmlDoc.transformNode(oXslDoc));
   // Response.Write(oXmlDoc.xml);
%>
  </tr>
  

</table>
<body>
</html>
<%
 }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>

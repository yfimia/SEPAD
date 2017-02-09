<%@ Language=JavaScript %>
<%
 Response.Expires = -1;


%>

<!-- #include file="../js/Adolibrary.inc" -->
<!-- #include file="../js/user.inc" --> 
<!-- #include file="../js/course.inc" --> 

<%  if ((Request.QueryString("eliminar").Count > 0) && (Request.QueryString.Item("eliminar") != ""))
    {
      var  filePath = Application("dataPath");   
      var  oConn    = Server.CreateObject("ADODB.Connection");
      var  oComm    = Server.CreateObject("ADODB.Command");
      oConn.Errors.Clear();
      oConn.Open( filePath );
      oComm.ActiveConnection = oConn;
      oComm.CommandText = "delete from Anotaciones where ((alumn = " + Session("userID") + ") and (Material = "+ Request.QueryString.Item("material") + "))";
      oComm.Execute();
    }
    else if ((Request.QueryString("material").Count > 0) && (Request.QueryString.Item("material") != ""))
    {
      var  filePath = Application("dataPath");   
      var  oConn    = Server.CreateObject("ADODB.Connection");
      var  oComm    = Server.CreateObject("ADODB.Command");
      var  oRec     = Server.CreateObject("ADODB.Recordset");
      oConn.Errors.Clear();
      oConn.Open( filePath );
      oRec.Open("select * from Anotaciones where ((alumn = " + Session("userID") + ") and (Material = "+ Request.QueryString.Item("material") + "))",oConn,3,3);
      if (oRec.EOF) {oRec.AddNew;}
      oRec.Fields.Item("alumn").Value = Session("userID");
      oRec.Fields.Item("Material").Value = Request.QueryString.Item("material");
      oRec.Fields.Item("Texto").Value = Request.Form.Item("texto");
      oRec.Update();
      oRec.Close();
    }
    
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       
      var lesson = Request.QueryString.Item("lesson") + "";

      if (lesson == "-1") {
        lesson = Session("lastLesson");
        lessonurl = Session("lastLessonUrl");
        //Response.Write(lesson);
      }
      
      var lessonurl = Request.QueryString.Item("lessonurl") + "";
      
      var updateState = Request.QueryString.Item("newState").Count;
      var newState = Request.QueryString.Item("newState") + "";
      Session("lastLesson") = lesson;
      Session("lastLessonUrl") = lessonurl;
      lessons = Session("lessons");
      fullCourse = Session("fullCourse");
      fullCourse.lastLessonID = lesson; 
      
      lessonindex = GetLessonIndex(lesson);
      
      lessonname = lessons.items.name[lessonindex];
      lessonurl = lessons.items.dir[lessonindex];
      actualstate = lessons.items.state[lessonindex];

//      Response.Write(lesson + "<BR>" + lessonindex + "<BR>" + actualstate + "<BR>" + newState + "<BR>");

      if (updateState == 0)
        {

          var dataPath = Application('dataPath');
          var oConn;
          var oRec;
          
           oConn = MakeConnection( oConn, dataPath );
          
          
           //Salvar la ultima leccion en la historia. 
   
           if ((lesson != "") && (Session("course") != "") && (Session("userID") != "")) {
                  
               Sql ='select * from Historia WHERE ([user] = ' + Session("userID") + ') and (course = ' +  Session("course")  + ')';
          
               oRec = Query( Sql, oRec, oConn  );		
               if (oRec.EOF) {
                 oRec.Close();  
                 Sql ='select * from Historia';
                 oRec = Query( Sql, oRec, oConn  );		
                 oRec.AddNew();
               }
                                
                 oRec.Fields.Item("User").value      = Session("userID");
                 oRec.Fields.Item("Lesson").value    = lesson;
                 oRec.Fields.Item("Course").value    = Session("course");
                 oRec.Fields.Item("LessonIndex").value    = lessonindex;
                 oRec.Update();      

               oRec.Close();  
           }    
          //*********************************//    
          
          // ********** Save to user history ********
          
           if (( lesson != "") && (Session("course") != "") && (Session("userID") != "")) 
           {
                  
               Sql ='SELECT * FROM HistoriaLecciones WHERE ([UserID] = ' + Session("userID") + ') AND (Lesson = ' +  lesson  + ')';
          
               oRec = Query( Sql, oRec, oConn  );	
               
               userId 		= Session("userID");	
               fullUserName = Session("fullName");              
               lessonId     = lesson;
               courseId     = Session("course");
               lessonName   = lessonname;
               cant         = 0;
               
               if ( oRec.EOF ) 
               {
                 oRec.Close();  
                 
                 Sql = "SELECT * FROM HistoriaLecciones";
                 oRec = Query( Sql, oRec, oConn  );		
                 oRec.AddNew();
               }
               else
               {
               	 cant = oRec.Fields.Item("Cant").value;
               }
                                
               oRec.Fields.Item("UserID").value  	   = userId;
               oRec.Fields.Item("Lesson").value  	   = lessonId;
               oRec.Fields.Item("Course").value  	   = courseId;
               oRec.Fields.Item("fullUserName").value  = fullUserName;
               oRec.Fields.Item("lessonName").value    = lessonName;
               oRec.Fields.Item("Cant").value    	   = cant + 1;

               oRec.Update();      
               oRec.Close();  
           }    
          // ******************************************
          
          Sql ='SELECT B.ID as MID, B.url, B.tama, B.title, B.description , A.Dir FROM Lecciones AS A, Ficheros AS B, Ficheros_de_Lecciones AS C, Cursos as D WHERE (A.ID = ' + lesson + ') and (A.id = C.lessonID) and (B.id = C.fileID) and (D.id = ' + Session("course") + ' ) ORDER BY B.ID';
          
          oRec = Query( Sql, oRec, oConn  );		
          
          
          cadena = "";
          
          url = oRec.Fields.Item("url") + "";

       }
       else
         {
              var  filePath = Application("dataPath");   
              var  oConn    = Server.CreateObject("ADODB.Connection");
              var  oComm    = Server.CreateObject("ADODB.Command");
              var  oRec     = Server.CreateObject("ADODB.Recordset");              
              oConn.Errors.Clear();
              oConn.Open( filePath );
              oRec.Open("select * from Lecciones where (ID=" + lesson + ")",oConn,3,3);
              oRec.Fields.Item("state").Value = newState;
              oRec.Update(); 
              Session("lessons").items.state[lessonindex] = newState + "";
              DestroyAdoObjects(oConn, oRec);
              Response.Redirect("docs.asp?uid=" + uid + "&lesson=" + lesson);
         }
%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/EvalResult.css" type="text/css">

<script type="text/javascript" language="javascript" src="../js/Windows.js"></script>

<script language=JavaScript>
    function openwin(uid){
      window.open("SubGrupo.asp?uid=" + uid,null,"height=200,width=340,toolbar=no,status=no,directories=no,menubar=no,scrollbars=yes,resizable=yes");
    }
</script>

</head>
<body bgcolor="#FFFFFF" text="#000000" style="overflow-y:scroll">


<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar" width=100%>
        <tr> 
          <td  width=10% ><img src="../images/<%=Session("skin")%>/LessonImg.gif" width="80" height="54"></td>
          <td  width=70% nowrap align=left class="HeaderTable">Aula Virtual</td>
<%
          var flag = ((Session("userID") == Session("courseOwner")) || (Session("userID") == Session("cordinador")) || (Session("PermissionType") == ADMINISTRATOR) );
          if (flag == true) {
%>          
          <td  width=8% align=right valign=center><b>Estado:</b></td>
          <td  width=12% align=right valign=center>
            <select name="StateCB" size="1" onchange="location='docs.asp?uid=<%=uid%>&lesson=<%=lesson%>&newState=' + this.options[this.selectedIndex].value">
<%
              if (actualstate == "1")
%>
                <OPTION VALUE="1" SELECTED>En curso
<%            
              else
%>              
                <OPTION VALUE="1">En curso
<%
              if (actualstate == "4")
%>
                <OPTION VALUE="4" SELECTED>No visible
<%
              else
%>
                <OPTION VALUE="4">No visible
<%
              if (actualstate == "5")
%>              
                <OPTION VALUE="5" SELECTED>Deshabilitado
<%
              else
%>
                <OPTION VALUE="5">Deshabilitado
            </select>
          </td>
<%
          }
%>          
        </tr>
      </table>
    </td>
  </tr>
</table>  
<br>  
      <table border="0" cellspacing="1" cellpadding="1" width=95% align=center class="prologueTable">
        <tr> 
          <td width=30% class="ToolBar" colspan=2 align=center><b>Datos Generales</b>
          </td>
        </tr>

        <tr> 
          <td width=30% class="MessageTR1" align=right><b>Modalidad académica:</b>
          </td>
          <td width=70% class="MessageTR1" align=left> <%=Session("moduloName")%>
          </td>
        </tr>

        <tr> 
          <td width=30% class="MessageTR" align=right><b>Módulo:</b>
          </td>
          <td width=70% class="MessageTR" align=left> <%=Session("courseName")%>
          </td>
        </tr>
        <tr> 
          <td width=30% class="MessageTR1" align=right><b>Lección:</b>
          </td>
          <td width=70% class="MessageTR1" align=left><%=lessonname%>
          </td>
        </tr>
      </table>
<% 
    if (Session("sgName") != "") {
%>        
            <br>
            <table border="0" cellspacing="1" cellpadding="1" width=95%  align=center class="prologueTable">
              <tr> 
               <td class="ToolBar" width=100% colspan=4  align=center><b>SubGrupo Tutoreado</b></td>
              </tr>

              <tr> 
               <td class="ToolBar" width=25%  align=center><b>SubGrupo</b></td>
               <td class="ToolBar" width=75% colspan=3  align=center><b>Tutor</b></td>
              </tr>
              <tr> 
               <td class="MessageTR1" width=25%  align=center><a href="javascript:openwin(<%=uid%>)" > <%=Session("sgName")%> </a></td>
               <td class="MessageTR1" width=25%   align=center><a href="javascript:abreVentana('',350,280,'userDetails.asp?uid=<%=uid%>&id=<%=Session("tutID")%>', 'no', 'no')"><%=Session("sgTutor")%></a></td>
               <td class="MessageTR1" width=25%   align=center><%=Session("sgTutorName")%></td>
               <td class="MessageTR1" width=25%   align=center><a href="mailto:<%=Session("sgEmail")%>"><%=Session("sgEmail")%></a></td>
              </tr>
            </table>             

<%
    }
%>        
<br>
<table class="prologueTable" border="0" cellspacing="0" cellpadding="0" width=95%  align=center >
  <tr> 
    <td> 
<%
 cadena = '';
HTML = "";

if ((updateState == 0) && (oRec.recordCount >= 0))
 {
 
%>
      <table width="100%" cellpadding="1" cellspacing="1" border="0">
        <tr> 
          <td  nowrap width="100%" colspan=5 class="ToolBar" align="center"><b>Materiales de la lección</b></td>
        </tr>

        <tr width="100%"> 
          <td width="3%" class="ToolBar"></td>
<%
    if (flag == true) {
%>          
          <td width="3%" class="ToolBar"></td>
<%}%>          
          <td width="30%" class="ToolBar" align="center"><b>Archivo</b></td>
          <td width="17%" class="ToolBar" align="center"><b>Tama&ntilde;o</b></td>
          <td width="47%" class="ToolBar" align="center"><b>Descripci&oacute;n</b></td>
        </tr>
 
<%
  trclass = 'MessageTR';
  while (oRec.EOF == false)
    {
       lessonurl = "/" + oRec.Fields.Item("dir") + "";
       url = oRec.Fields.Item("url") + "";
       title = oRec.Fields.Item("title") + "";
       if ((title == "") || (title + "" == "null")) title = url;
       desc = oRec.Fields.Item("description") + "";
       if (desc == "null") desc = "";
       size = oRec.Fields.Item("tama") + "";
/*     }
     else { //Esto es para compatibilizar con las lecciones que no tengan directorio...
       tmp = oRec.Fields.Item("url") + "";

       p1 = tmp.indexOf("@#");
       title = tmp.substring(0,p1);
       url = title;
       
       p2 = tmp.indexOf("@#", p1 + 2);
       size = tmp.substring(p1 + 2,p2);
       
       desc = "";
     }      
    
*/     
     cadena =  "../courses/course" + Session('course') +  lessonurl + '/' + url;     
     if (trclass == 'MessageTR') trclass = 'MessageTR1'; else trclass = 'MessageTR';
     ident = oRec.Fields.Item("MID");
%>     


        <tr width="100%" class="<%=trclass %>" valign="middle"> 
          <td class="ToolBar"  width="3%">
<%  
      if (Session("userid") != GUEST_USER) {
%>          
            <a href="Anotaciones.asp?uid=<%=uid%>&material=<%=ident%>&lesson=<%=lesson%>">
              <img title="Anotaciones" src="../images/gold/Anotaciones.gif"  width="20" height="20" border="0" >
            </a>   
<%
      }
%>            
          </td>
<%
    if (flag == true) {
%>          
          <td class="ToolBar"  width="3%">
            <a href="UpLoadMateriales.asp?uid=<%=uid%>&material=<%=ident%>&url=<%=lessonurl%>&lesson=<%= lesson %>">
              <img title="Administrar" src="../images/Admin.gif"  width="20" height="20" border="0" >
            </a>   
          </td>
<%}%>          
          <td align=center width="30%" align="left"><a target=_blank href="<%=cadena %>"><%=title %></a></td>
          <td width="17%" align="right"><%=size %>kb</td>
          <td width="45%" align="justify"><%=desc %></td>
        </tr>


<%     
     oRec.Move(1);


    }
%>    
       </table>
    </td>
  </tr>
</table>
<%  
 }  
 else
   {
%>
     <td>
       Esta lección no tiene materiales asociados
     </td>
  </tr>
</table>   
 
<%   
   }   
  if (updateState == 0)
  DestroyAdoObjects( oConn, oRec );        
%>
<br>
<% //Esto lo usa las busquedas para el patron por defecto... no quitar
   var pattern ="";
   var tipo=1;
%>
<!-- #include file="../js/avsearch.inc" -->

</body>
</html>
<%
 }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>
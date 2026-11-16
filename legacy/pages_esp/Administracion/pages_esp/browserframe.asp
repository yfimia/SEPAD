<%@ Language=JavaScript%>
<%
 Response.Expires = -1;
%>
<!-- #include file='../js/adolibrary.inc' -->
<!-- #include file='../js/user.inc' --> 

<% 

    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       
    //se reciben los parametros courseID,courseName, lastlesson, lastlessonName
 if (Request.QueryString.Count != 1)   { 
   //Error en el pase de parametros.
      Response.Redirect("errorpage.asp?tipo=Error&short=" + BAD_PARAMS_SHORT  + "&desc=" + BAD_PARAMS_TEXT);
 }    
    
//    Response.Write(Session("course") + "  " +    Session("courseName") + " " +    Session("lastLesson") + " " +  Session("lastLessonUrl"));
/*   if (Request.QueryString.Item("init") + "" == "1")  {
     deleteHistory(Session("userID"), Session("course"));
   } 
*/  
   function setLastLesson(lastLesson)
    {
     Session("lastLesson") = lastLesson;
     //Response.Write("alert('" + lastLessonUrl + "');");
    }

%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title><%=Request.QueryString.Item("courseName")%></title>
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">

</head>

<script language="vbscript" RUNAT="Server">
 function getAtualTime
  getAtualTime = Now()
 End function 

</script>
<% 
    // Variables Globales (campos de la base de datos )para las  leciones... 
    var lessonId = new Array();
    var lessonDir = new Array();
    var lessonName = new Array();
    var lessonNivel = new Array();
    var lessonPrimary = new Array();
    var lessonState = new Array();

     
     for (i = 0;i <= Session("lessonCant");i++)
        {
          lessonId[i]      = Session("lessons").items.id[i];
          lessonName[i]    = Session("lessons").items.name[i];
          lessonDir[i]    = Session("lessons").items.dir[i];
          lessonNivel[i]   = Session("lessons").items.nivel[i];
          lessonPrimary[i] = Session("lessons").items.primary[i];
          lessonState[i] = Session("lessons").items.state[i];
          
        }

     
      // para el Cliente....
       var codigo = "";      
       for (i = 0;i <= Session("lessonCant");i++)
        {
          codigo = codigo +
           "cid["   + i + "] = '" + lessonId[i]   + "'; " + 
           "cdir[" + i + "] = '" +   lessonDir[i]  + "'; "  +
           "cstate[" + i + "] = '" +   lessonState[i]  + "'; "  +
           "cname[" + i + "] = '" + lessonName[i] + "'; ";
        }
    
    
    // Consulta para el Grafo...

    var oConn;
    var oRec;
 
    var filePath = Application('dataPath');
    
    oConn = MakeConnection( oConn, filePath );       
  
   
    Sql ="SELECT A.Id, B.Lesson FROM Lecciones A, Proxima_Leccion B,Relaciones_del_Curso C WHERE (A.Course =" + Session("course") + ") AND (A.Id = C.Lesson) AND (C.Linked = B.Id) AND (A.state <> " + MOD_ACA_NOVISIBLE + ")";

    oRec = Query( Sql, oRec, oConn  );
     
    

    //Objetos...        
    function sitems(id,lesson)
     { 
       this.id = id;
       this.lesson = lesson;
       
     }
    
    function srecord(id,lesson)
     { 
      this.items = new sitems(id,lesson);
     } 
    
    // Varibles globales para el grafo... 
    var idgf = new Array();
    var lessongf = new Array();

       /* Inicializacion de los campos */  
  
      // para el Servidor....
    var i = 0;   
    while (oRec.EOF == false)
     { 
       
      idgf[i]      = oRec.Fields.Item("Id") + "";
      lessongf[i]  = oRec.Fields.Item("Lesson") + "";
      
      oRec.Move(1);
      i += 1;
     }
     Session("grafoCant") = i - 1;  
     Session("grafo") = new srecord(idgf,lessongf); 
    
   // para el Cliente....
       var codigogf = "";      
       for (i = 0;i <= Session("grafoCant");i++)
        {
          codigogf = codigogf +
           "gid["   + i + "] = '" + idgf[i]   + "'; " + 
           "glesson[" + i + "] = '" + lessongf[i] + "'; "; 
           
        }
          
  // Cierre de la coneccion...    
 DestroyAdoObjects( oConn, oRec );    
    
    
   
%>

<script LANGUAGE="javascript">
<!--
 var lastLesson = <%= "'" + Session("lastLesson") + "'"%>;
 
 
  function getFileBack()
   { 
    pos = -1;
    var lcant = <%= Session("lessonCant")%>;
    for (i=0;i<=lcant;i++)
     {
      if (lessons.items.id[i] == lastLesson)
       {
        if (i > 0)
        {
         lastLesson = lessons.items.id[i - 1];
         return i - 1;
         }
        else
         return -1;
       
       }
     }
    return -1;
   } 


 
    
  function getFileNext()
   {
    var lcant = <%= Session("lessonCant")%>;
    for (i=0;i<=lcant;i++)
     {
      //alert(lastLesson + "  " + lessons.items.id[i]) ;
      if (lessons.items.id[i] == lastLesson)
       { 
        if (i < lcant)
        {
         lastLesson = lessons.items.id[i + 1];
         return i + 1;
        }
        else
         return -1;
       }
     }
     return -1;
   }   
   
  



  //Objetos Records del cliente para Leccionnes...  
  function citems(id,name, dir, state)
     { 
       this.id = id;
       this.dir = dir;
       this.name = name;
       this.state = state;

     }
    
  function crecord(id,name, dir, state)
    { 
     this.items = new citems(id,name, dir, state);
    }
    
  //Objetos Records del cliente para Grafo...  
  function gitems(id,lesson)
     { 
       this.id = id;
       this.lesson = lesson;
     }
    
  function grecord(id,lesson)
    { 
     this.items = new gitems(id,lesson);
    }
  
  
  /* Incializacion de los campos en el cliente Leccionnes */     
   var 
    cid = new Array();
    cname = new Array();
    cdir = new Array();
    cstate = new Array();
    
    <%= codigo%>    //Inicializacion de los arreglos desde el servidor... 
    var lessons = new crecord(cid,cname, cdir, cstate);
  
  /* Incializacion de los campos en el cliente Grafo */       
    gid = new Array();
    glesson = new Array();
    
    <%= codigogf%>    //Inicializacion de los arreglos desde el servidor... 
    var grafo = new grecord(gid,glesson);
    
    
    
 
  // Devuelve la posicion en lessons dado el id(pid)...
  function getPos(pid)
   {
    var i;
    for (i = 0;i <= <%=Session("lessonCant")%>;i++)
     {
      if (lessons.items.id[i] == pid)
       {
        return i;
       }
     }
   }
   
 
 function getId(pname)
  {
   var i;
   for (i = 0;i <= <%=Session("lessonCant")%>;i++)
    {
     if (lessons.items.name[i] == pname)
      {
       return i;
      }
    }
  }   
//-->
</script>





<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--




//back 
function previousLesson() 
 {
    var ls = getFileBack();
    if (ls == -1) alert("No existe una lección anterior porque esta es la primera...")
    else if ((lessons.items.state[ls] != <%=MOD_ACA_INCOURSE%>)) previousLesson(); 
    else 
      { 
        parent.parent.frames("leftFrame").execScript("LocateChild(" +  ls  + "  )");                 

        parent.frames("mainFrame1").location.href = "docs.asp?uid=<%=uid%>&lesson=" + lessons.items.id[ls] + "&lessonurl=" + lessons.items.dir[ls] + "&lessonindex=" + ls + "&lessonname=" + lessons.items.name[ls];
        fullCombo(); 
      }  
 }

//next
function nextLesson() 
 { 
    var ls = getFileNext();
    //alert(ls);
    if (ls == -1) alert("No existe una próxima lección porque esta es la última...");
    else if ((lessons.items.state[ls] != <%=MOD_ACA_INCOURSE%>)) nextLesson();      
    else
      {
        parent.parent.frames("leftFrame").execScript("LocateChild(" +  ls  + "  )");   
        parent.frames("mainFrame1").location.href = "docs.asp?uid=<%=uid%>&lesson=" + lessons.items.id[ls] + "&lessonurl=" + lessons.items.dir[ls] + "&lessonindex=" + ls + "&lessonname=" + lessons.items.name[ls];
        fullCombo();
      } 
 }
 
//volver a lista de cursos
function courseList() 
 {
   parent.parent.frames("mainFrame").location.href = "Course.asp?uid=<%=uid%>&ID=<%=Session("course")%>&Name=<%=Session("courseName")%>";
   parent.parent.frames("leftFrame").location.href = "selection.asp?uid=<%=uid%>&flag=1";
 }

//evaluarse
function evaluate() 
 { 
//   if (confirm('¿Está seguro que desea evaluarse? Si deja la evaluación en blanco se le otorgara 2')==true)
    { 
      parent.frames("mainFrame1").location.href = "evaluaciones.asp?uid=<%=uid%>&lastLesson=" + lastLesson;
    }
 }

 
//-->
</script>


<body bgcolor="#FFFFFF" text="#000000" class="ComboBody">
<table border="0" cellspacing="0" cellpadding="0" width="100%">
  <tr> 
    <td align=right> 
      <table border="0"  cellspacing="0" cellpadding="2" width="171">
        <tr> 
          <td ><a  href="javascript:courseList()"  ><img title="Generalidades del módulo" src="../images/<%=Session("skin")%>/Generalidades.gif"  width="20" height="20" border="0" ></a></td>
          <td class="Separator">|</td>
          <td ><a id="back" href="docs.asp?uid=<%=uid%>&lesson=-1" target="mainFrame1"  ><img title="Aula Virtual" src="../images/<%=Session("skin")%>/avirtual.gif"  width="20" height="20" border="0" ></a></td>
          <td class="Separator">|</td>
          <td ><a id="back" href="javascript:previousLesson()"   ><img title="Lección&nbsp;Anterior" src="../images/<%=Session("skin")%>/PrevLesson.gif"  width="20" height="20" border="0" ></a></td>
          <td class="Separator">|</td>
          <td  ><a href="javascript:nextLesson()"    ><img title="Lección&nbsp;Siguiente" src="../images/<%=Session("skin")%>/NextLesson.gif"  width="20" height="20" border="0" ></a></td>
          <td class="Separator">|</td>
          <td ><a  href="javascript:evaluate()"  ><img title="Autoevaluación" src="../images/<%=Session("skin")%>/Evaluate.gif"  width="20" height="20" border="0" ></a></td>
          <td class="Separator">|</td>
          <td ><a  href="UpLoadFile.asp?uid=<%=uid%>"  target="mainFrame1"><img title="Enviar&nbsp;trabajos" src="../images/<%=Session("skin")%>/UploadDocument.gif"  width="20" height="20" border="0" ></a></td>
          <td class="Separator">|</td>
          
          <td ><a  href="ListadoAnotaciones.asp?uid=<%=uid%>"  target="mainFrame1"><img title="Anotaciones del M&oacute;dulo" src="../images/<%=Session("skin")%>/Anotaciones.gif"  width="20" height="20" border="0" ></a></td>
          <td class="Separator">|</td>
          
          <td ><a  href="uestadisticas.asp?uid=<%=uid%>" target="mainFrame1"  ><img title="Calificaciones" src="../images/<%=Session("skin")%>/Calificaciones.gif"  width="20" height="20" border="0" ></a></td>
          <td class="Separator">|</td>
          <td ><a  href="CalifOfic.asp?uid=<%=uid%>" target="mainFrame1"><img title="Calificaciones Oficiales" src="../images/<%=Session("skin")%>/CalifOfic.gif"  width="20" height="20" border="0" ></a></td>
          <td class="Separator">|</td>
          <td ><a  href="newChatModuleRoomList.asp?uid=<%=uid%>&course=<%=Session("course")%>&courseName=<%=Session("courseName")%>"  target="mainFrame1"><img title="Chat" src="../images/<%=Session("skin")%>/chat1.gif"  width="20" height="20" border="0" ></a></td>
          <td class="Separator">|</td>
          <td ><a  href="ShowForumTopics.asp?uid=<%=uid%>&course=<%=Session("course")%>&pn=0&courseName=<%=Session("courseName")%>"  target="mainFrame1"><img title="Foros de discusión" src="../images/<%=Session("skin")%>/foros.gif"  width="20" height="20" border="0" ></a></td>
          <td class="Separator">|</td>
          <td ><a  href="Anuncios.asp?uid=<%=uid%>&idmodulo=<%=Session("modulo")%>"  target="mainFrame1"><img title="Anuncios" src="../images/<%=Session("skin")%>/Anuncios.gif"  width="20" height="20" border="0" ></a></td>
          <td >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
          
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="2" class="PublishedCoursesTD">
  <tr> 
    <td width="14%" align="center"><B class="AsociativedThemes">Temas&nbsp;Relacionados:</b></td>
    <td width="78%"> 
      <select id="select1" name="select1" onChange="select1_onchange()" class="TemasAsociados">
      </select>
    </td>
  </tr>
</table>

<script LANGUAGE="javascript">
<!--
 
 // Inicializacion del comboBox.....

 //Devuelve el nobre dado el id (pid)...
 function getName(pid)
  {
   for (j = 0;j <= <%=Session("lessonCant")%>;j++)
    {
     if (lessons.items.id[j] == pid)    
      {
       return lessons.items.name[j]; 
      }
    }
   
  }

// Llena el combo con los temas asociados...  
function fullCombo()
 {
  var oOption;
  var cad;
  var sl;

   sl = document.all.select1; 
   
   while (sl.options.length > 0)
    {
     sl.options.remove(0);
    }
   oOption = document.createElement("OPTION");
   cad = getName(lastLesson);
   
   select1.title = cad;
   oOption.text =  cad;
   oOption.value = "0";
   sl.add(oOption);
   sl.options(0).selected = true;
   
   for (i = 0;i <= <%=Session("grafoCant")%>;i++)
    {
     if (grafo.items.lesson[i] == lastLesson)
      {
       oOption = document.createElement("OPTION");
       cad = getName(grafo.items.id[i]);
       if (cad != "")
        {
         oOption.text =  cad;
         oOption.value = i + "";
         sl.add(oOption);
        } 
     }
    else
     {
      if (grafo.items.id[i] == lastLesson)
       {
        oOption = document.createElement("OPTION");
        cad = getName(grafo.items.lesson[i]);
        if (cad != "")
         {
          oOption.text =  cad;
          oOption.value = i + "1";
          sl.add(oOption);
         } 
       }
    } 

   }
 }
 parent.parent.frames("leftFrame").location.href = "tmptreeview.asp?uid=<%=uid%>";
 if (lastLesson != -2) fullCombo(); 
//-->
</script>

<script LANGUAGE="javascript">
 <!--

//Eventos de los servicios del browser..

 
//Seleccionar leccion asociada
function select1_onchange() 
 {
   var sl = document.all.tags("SELECT").item(0);
   var p = select1.selectedIndex;
   var po = getId(sl.options(p).text);   
   if ((po != -1) && (lessons.items.state[po] == <%=MOD_ACA_INCOURSE%>)) {
     parent.parent.frames("leftFrame").execScript("LocateChild(" +  po  + "  )");   
     parent.frames("mainFrame1").location.href = "docs.asp?uid=<%=uid%>&lesson=" + lessons.items.id[po] + "&lessonurl=" + lessons.items.dir[po] + "&lessonindex=" + po + "&lessonname=" + lessons.items.name[po];
     lastLesson = lessons.items.id[po];
     lastLessonUrl = lessons.items.dir[po];
     lastLessonName = lessons.items.name[po];
     lastLessonIndex = po;
  
     fullCombo(); 
   }  
   else alert("Esta lección no está disponible en estos momentos porque se encuentra deshabilitada.");
 }

 //-->
 </script>

</body>
</html>
<%
 }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>

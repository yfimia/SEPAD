<%@ Language=JavaScript %>
<%
   Response.Expires = -1;
%>
<!-- #include file='../js/adolibrary.inc' -->
<!-- #include file="../js/user.inc" -->

<%   
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       
%>

<HTML>
<head>
<title>Listado de módulos</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">

<LINK href="../css/<%=Session("skin")%>/Main1.css" rel=stylesheet type=text/css>
</head>
<body bgcolor="#FFFFFF" text="#000000" class="TreeViewBody" style="overflow-y:scroll">

<% var pattern ="";
   var tipo=1;
%>
<!-- #include file="../js/search.inc" -->

<%

     

  var modules = new Array();

function Module(ID, Name, Description, State) 
     {this.ID = ID;
      this.Name = Name;
      this.Description = Description;
      this.State = State;
      this.Courses = new Array();
     }  

function load_Module_Courses() {

      var filePath = Application('dataPath');
      var oConn;
      var oRec;
 
      oConn = MakeConnection( oConn, filePath );
      Sql = "SELECT Modulo_Curso.Curso, Modulo_Curso.Modulo " + 
            "FROM Modulo INNER JOIN (Cursos INNER JOIN Modulo_Curso ON Cursos.ID = Modulo_Curso.Curso) ON Modulo.ID = Modulo_Curso.Modulo " + 
            "ORDER BY Modulo.Name, Cursos.Name";
      oRec = Query( Sql, oRec, oConn  );		

      var i = 0; 
      while (!oRec.EOF)
        {
         module           = oRec.Fields.Item("course").Value;
         course			  = oRec.Fields.Item("module").Value;
         
         if ((i != 0) || (module !=  lastmodule)) { 
           i++;
           lastmodule = module;
         }  
         
         modules[ i ].courses[ Modules[ i ].courses.length ] = getCourseIndexByID(course);
        } 
      DestroyAdoObjects( oConn, oRec );
  
}

function  load_modules() {
      var count = 0; 
  
      var filePath = Application('dataPath');
      var oConn;
      var oRec;
 
      oConn = MakeConnection( oConn, filePath );
      Sql = "select * FROM Modulo";

      oRec = Query( Sql, oRec, oConn  );		


      while (!oRec.EOF)
        {
         ID               = oRec.Fields.Item("ID").Value;
         Name		  = oRec.Fields.Item("Name").Value;
         Description      = oRec.Fields.Item("Description").Value;
         State            = oRec.Fields.Item("State").Value;
         modules[count]   = new Module(ID, Name, Description, State);
         count++;
        } 
      DestroyAdoObjects( oConn, oRec );

      load_Module_Courses();  
}         


function download_Modules() {
      
%>
<table border="0" celsspacing="0" cellspadding="0" width="100%" class="ToolBar">
  <tr>
    <td>
      <table border="0" celsspacing="0" cellspadding="0">
        <tr>
          <td><a class="ToolLink" href="javascript:Do_Expand_All()">&nbsp;Expandir&nbsp;</a></td>
          <td>|</td>
          <td><a class="ToolLink" href="javascript:Do_Collapse_All()">&nbsp;Contraer&nbsp;</a></td>
        </tr> 
      </table>
    </td>
  </tr>
</table>
<table border="0" width="100%" cellpadding="3" cellspacing="0" class="ToolBar">
    <tr> 
      <td><b>Módulos:</b></td>
	</tr>
</table>

<%
 
      for( var i = 0; i < modules.length; i++ ) {
%>

<ul>
             
      <li class="clsHasKidsCollapsed"> 
          <span>
             <IMG SRC="../images/<%=Session("skin")%>/TreeViewImg.gif"   class="IconClass" />  
             <%= modules[i].Name %>
          </span>
          <ul>
<%
        for (j = 0; j < modules[i].courses.length; j++ ) { 
          index = modules[i].courses[j]; 
%>             
             <li class="clsDocument" > 
               <a   href="Course.asp?uid=<%=uid%>&course=<%=courses[index].ID%>&courseName=<%=courses[index].Name%>&lastlesson=<%=lessons[index].ID%>1&lastlessonUrl=<%=lessons[index].dir%>&lastlessonindex=<%=lessons[index].index%>&lastlessonname=<%=lessons[index].name%>&Atmost=<%=courses[index].Atmost%>&prologue=<%='../courses/course' + courses[index].ID + '/' + courses[index].Prologue%>&owner=<%=courses[index].Owner%>&ctdExAct=<%=courses[index].ctdExAct%>&ctdExRet=<%=courses[index].ctdExRet%>&MaxFile=<%=courses[index].MaxFile%>&MaxMem=<%=courses[index].MaxMem%>">
                <IMG SRC="../images/<%=Session("skin")%>/TreeViewImg.gif"   class="IconClass" />  <%=courses[index].Name%> 
               </a>
             </li>             
<%
        }      
%>
        </ul>     
    </li>       
  
</ul>

<%
      }        
}


  
init_Modules();
%>

</BODY>

</HTML>
<%
 }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>

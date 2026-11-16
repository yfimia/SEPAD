<%@ Language=javascript %>
<%
 Response.Expires = -1;
%>

<!-- #include file="../js/Adolibrary.inc" -->
<!-- #include file="../js/user.inc" --> 
<!-- #include file="../js/course.inc" --> 

<%
function view()
{
  var IDLesson = new Array();
  var IDMaterial = new Array();
  var LessonID = new Array();
  
  var  filePath = Application("dataPath");   
  var  oConn    = Server.CreateObject("ADODB.Connection");
  var  oComm    = Server.CreateObject("ADODB.Command");
  var  oRec     = Server.CreateObject("ADODB.Recordset");
  var  oRec2     = Server.CreateObject("ADODB.Recordset");
  oConn.Open( filePath );
  oRec.Open("select ID from Lecciones where Course=" + Session("course"),oConn,3,3);
  i = -1;
  while (!oRec.EOF)
  {
    i = i + 1;
    IDLesson[i] = oRec.Fields("ID").Value;
    oRec.Move(1);
  }
  oRec.Close();
  k = -1;
  for (j=0; j<=i; j++)
  {
    oRec.Open("select * from Ficheros_de_lecciones where (lessonID=" + IDLesson[j] + ")",oConn,3,3);
    while (!oRec.EOF)
    {
      k = k + 1;
      IDMaterial[k] = oRec.Fields("fileID").Value;
      LessonID[k] = IDLesson[j];
      oRec.Move(1);
    }
    oRec.Close();
  }

%>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td class="HeaderTable" align=center>Anotaciones&nbsp;(<%=Session("courseName")%>)</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
  
<table border="0" cellspacing="1" cellpadding="1" width="100%">
  <tr width="100%">  
    <td Class="ToolBar" align="center" width="50%"><b>Material</b></td>
    <td Class="ToolBar" align="center" width="50%"><b>Lecci&oacute;n</b></td>
  </tr>
<%
  h = 0;
  if (IDMaterial.length > 0)
  {
    for (i=0; i<=k; i++)
    {
      oRec.Open("select ID from Anotaciones where (Material=" + IDMaterial[i] + ")");
      if (!oRec.EOF)
      {
        oRec.Close();
        oRec.Open("select title from Ficheros where (ID =" + IDMaterial[i] + ")", oConn,3,3);
        oRec2.Open("select Name from Lecciones where(ID =" + LessonID[i] + ")", oConn,3,3);
        if (h % 2 == 0)
        {
          h = h + 1;
%>
    <tr width="100%">      
      <td Class="MessageTR" width="50%" align="center">
        <a href="Anotaciones.asp?uid=<%=Request.QueryString.Item("uid")%>&material=<%=IDMaterial[i]%>&lesson=<%=LessonID[i]%>">
          <%=oRec.Fields("title").Value%>
        </a>
      </td>  
      <td Class="MessageTR" width="50%" align="center"><%=oRec2.Fields("Name").Value%></td>
    </tr>  
<%    
	    }
	    else
	    {
	      h = h + 1;
%>	  
	<tr width="100%">      
      <td Class="MessageTR1" width="50%" align="center">
        <a href="Anotaciones.asp?uid=<%=Request.QueryString.Item("uid")%>&material=<%=IDMaterial[i]%>&lesson=<%=LessonID[i]%>">
          <%=oRec.Fields("title").Value%>
        </a>
      </td>
      <td Class="MessageTR1" width="50%" align="center"><%=oRec2.Fields("Name").Value%></td>
    </tr>    
<%
  	    }
  	    oRec2.Close();
        oRec.Close();
      }
      else
      {
        oRec.Close();
      }
    }
    if (h == 0) 
    {
%>    
      <tr width="100%">
        <td colspan="2" width="100%" align="center">No hay anotaciones hechas en este M&oacute;dulo</td>
      </tr>
<%      
    }
  }
  else 
  {
%>
    <tr width="100%">
      <td colspan="2" width="100%" align="center">No hay anotaciones hechas en este M&oacute;dulo</td>
    </tr>  
<%
  }  
%>

</table>

<%  
} 
%>


<HTML>
<HEAD>

<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/EvalResult.css" type="text/css">
<script type="text/javascript" language="javascript" src="../js/Windows.js"></script>

</HEAD>


<BODY>
<%

if (Request.QueryString.Item("uid") == Session("uid"))
{
  view();
}
else
{
  Response.Redirect("errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);
}

%>
</BODY>
</HTML>

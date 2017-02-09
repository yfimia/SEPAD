<%@ Language=JavaScript %>

<!-- #include file="../js/adolibrary.inc" -->
<!-- #include file='../js/user.inc' -->

<%
Response.Expires = -1;
%>

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       



var RestrictedWords = new Array('A', 'ANTE', 'BAJO', 'CON', 'CONTRA', 'DE', 'DESDE', 'EN', 'ENTRE', 'HACIA', 'HASTA', 'PARA', 'POR', 'SEGUN', 'SIN', 'SOBRE', 'TRAZ', 'Y', 'E', 'NI', 'O', 'U', 'PERO', 'MAS', 'SINO');
var strings = new Array();
var urls = new Array();
var count = 0;
var flag = false;
pattern = Request.QueryString.Item("pattern") + "";

if (pattern.charAt(0) == '"')
    var string = pattern.substr(1, pattern.length - 2);
  else 
    var string = pattern;

function Matricula(CourseID)

  {
    result = ( IsValidInModule(Session("userID"), CourseID)  )

    return ((result));
  }

function Esta(word)

  {
    var i = 0;
    var stmp = word.toUpperCase();
    while ((i < RestrictedWords.length) && (RestrictedWords[i] != stmp))
      i++
    if (i == RestrictedWords.length)
        return (false)
      else
        return (true);
  }

function Extract(vName, string)

  {
    var result = "";
    var stmp = "";
    for (i = 0; i < string.length; i++)
      {
        if (string.charAt(i) != " ")
          stmp = stmp.concat(string.charAt(i));
          else
            {
              if (!(Esta(stmp)))
                {
                  if (result != "")
                      result = result.concat(" OR (" + vName + " LIKE '%" + stmp + "%')");
                    else
                      result = result.concat("(" + vName + " LIKE '%" + stmp + "%')");
                }
              stmp = "";
            }
      }
    if ((stmp != "") && (!(Esta(stmp))))
      {
        if (result != "")
            result = result.concat(" OR (" + vName + " LIKE '%" + stmp + "%')");
          else
            result = result.concat("(" + vName + " LIKE '%" + stmp + "%')");
      }
    return (result);
  }

function Find(CourseID, CourseName)

  {  
    flag = Matricula(CourseID);
    //flag = true;
    var dataPath = Application('dataPath');
    var oConn;
    var oRec;
    var stmp;
    oConn = MakeConnection(oConn, dataPath);
    stmp = Extract("Name", string);
    if (string == pattern)
      {
        if (stmp != "")
            SQL = "SELECT Name FROM Cursos WHERE (ID=" + CourseID + ") AND (" + stmp + ")";
          else
            SQL = "";
      }
      else
        SQL = "SELECT Name FROM Cursos WHERE (ID=" + CourseID + ") AND (Name LIKE '%" + string + "%')";
    if (SQL != "")
      {
        //Response.Write(SQL + "<BR>");
        oRec = Query(SQL, oRec, oConn);
        while (oRec.EOF == false)
          { 
            strings[count] = oRec.Fields.Item("Name") + "";
            urls[count] = "";
            count++;
            oRec.Move(1);
          }
      }
    if (string == pattern)
      {
        if (stmp != "")
            SQL = "SELECT Name FROM Lecciones WHERE (Course=" + CourseID + ") AND (" + stmp + ")";
          else
            SQL = "";
      }
      else
        SQL = "SELECT Name FROM Lecciones WHERE (Course=" + CourseID + ") AND (Name LIKE '%" + string + "%')";
    if (SQL != "")
      {
       // Response.Write(SQL + "<BR>");
        oRec = Query(SQL, oRec, oConn);
        while (oRec.EOF == false)
          { 
            strings[count] = CourseName + " >> " + oRec.Fields.Item("Name");
            urls[count] = "";
            count++;
            oRec.Move(1);
          }
      }
    stmp = Extract("A.title", string);
    if (string == pattern)
      {
        if (stmp != "")
            SQL = "SELECT A.title, A.url,  C.dir, C.Name FROM Ficheros A, Ficheros_de_lecciones B, Lecciones C WHERE (C.Course=" + CourseID + ") AND (B.lessonID=C.ID) AND (A.Id=B.fileID) AND (" + stmp + ")";
          else
            SQL = "";
      }
      else
        SQL = "SELECT A.title, A.url, C.Name, C.dir FROM Ficheros A, Ficheros_de_lecciones B, Lecciones C WHERE (C.Course=" + CourseID + ") AND (B.lessonID=C.ID) AND (A.Id=B.fileID) AND (A.title LIKE '%" + string + "%')";
    if (SQL != "")
      {
        //Response.Write(SQL + "<BR>");
        oRec = Query(SQL, oRec, oConn);
        while (oRec.EOF == false)
          {
            strings[count] = CourseName + " >> " + oRec.Fields.Item("Name") + " >> " + oRec.Fields.Item("title");
            if (flag == false)
               urls[count] = "";
            else
               urls[count] = "../courses/course" + CourseID + "/" + oRec.Fields.Item("dir") + "/" + oRec.Fields.Item("url");
            count++;
            oRec.Move(1);
          }
      }
    stmp = Extract("A.description", string);
    if (string == pattern)
      {
        if (stmp != "")
            SQL = "SELECT A.title, A.url, C.dir,  C.Name FROM Ficheros A, Ficheros_de_lecciones B, Lecciones C WHERE (C.Course=" + CourseID + ") AND (B.lessonID=C.ID) AND (A.Id=B.fileID) AND (" + stmp + ")";
          else
            SQL = "";
      }
      else
        SQL = "SELECT A.title, A.url, C.Name, C.dir FROM Ficheros A, Ficheros_de_lecciones B, Lecciones C WHERE (C.Course=" + CourseID + ") AND (B.lessonID=C.ID) AND (A.Id=B.fileID) AND (A.description LIKE '%" + string + "%')";
    if (SQL != "")
      {
       // Response.Write(SQL + "<BR>");
        oRec = Query(SQL, oRec, oConn);
        while (oRec.EOF == false)
          { 
            strings[count] = CourseName + " >> " + oRec.Fields.Item("Name") + " >> " + oRec.Fields.Item("title");
            if (flag == false)
              urls[count] = "";
            else
              urls[count] = "../courses/course" + CourseID + "/" + oRec.Fields.Item("dir") + "/" + oRec.Fields.Item("url");
            count++;
            oRec.Move(1);
          }
      }
  }

function Search()

  {
    if ((Request.QueryString.Item("courseID").count > 0) && (Request.QueryString.Item("courseID") + "" != "-1"))
      {
        //Find(Request.QueryString.Item("courseID"));
          var dataPath1 = Application('dataPath');
          var oConn1;
          var oRec1;
          oConn1 = MakeConnection(oConn1, dataPath1);
          SQL = "SELECT ID, Name FROM Cursos where ID = " + Request.QueryString.Item("courseID");
          oRec1 = Query(SQL, oRec1, oConn1);
          while (oRec1.EOF == false)
            {  
              Find(oRec1.Fields.Item("ID"), oRec1.Fields.Item("Name") + "");
              oRec1.Move(1);
            }
          DestroyAdoObjects(oConn1, oRec1);
        
      }
      else
       if ((Request.QueryString.Item("courseID").count == 0) || (Request.QueryString.Item("courseID") + "" == "-1"))       
        {

          var dataPath1 = Application('dataPath');
          var oConn1;
          var oRec1;
          oConn1 = MakeConnection(oConn1, dataPath1);
          SQL = "SELECT ID, Name FROM Cursos";
          //Response.Write(SQL + "<BR><BR><BR><BR>");
          oRec1 = Query(SQL, oRec1, oConn1);
          while (oRec1.EOF == false)
            {  
              Find(oRec1.Fields.Item("ID"), oRec1.Fields.Item("Name") + "");
              oRec1.Move(1);
              //Response.Write("<BR><BR><BR>");
            }
          DestroyAdoObjects(oConn1, oRec1);
        }
  }


function Print()

  {

    var clase = "MessageTR1";
   	last = -1;          

  
    for (i = 0; i < count; i++)
      {
      	last = i;          
      
        if (clase == "MessageTR1") clase = "MessageTR"; else clase = "MessageTR1";
        if (urls[i] != "") {
%>            
          <tr><td align="left" width="90%" colspan=2 class="<%=clase%>"><br><a target=_blank href="<%=urls[i]%>"><%=strings[i]%></a><br><br></td></tr>
<%              
        }
        else {
%>
          <tr><td align="left" width="90%" colspan=2 class="<%=clase%>"><br><%=strings[i]%><br><br></td></tr>

<%      }  
      }
      if (last < 0) 
        {
%>        
          <tr><td align="center" width="90%" colspan=2 class="MessageTR"><br>No hubo resultados.<br><br></td></tr>        
<%          
        }
   }
   
if ((Request.QueryString.Item("pattern").count > 0))   
  Search();

%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<LINK href="../css/<%=Session("skin")%>/MenuCSS.css" rel=stylesheet type=text/css>
<LINK href="../css/<%=Session("skin")%>/Main1.css" rel=stylesheet type=text/css>
<link rel="stylesheet" href="../css/<%=Session("skin")%>/EvalResult.css" type="text/css">
<title>Búsquedas</title>
</HEAD>
<body bgcolor="#ffffff" text="#000000">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td><img src="../images/<%=Session("skin")%>/LessonImg.gif" width="80" height="54"></td>
          <td class="HeaderTable">Búsquedas</td>
        </tr>
      </table>
    </td>
  </tr>
</table>  
<%
if ((Request.QueryString.Item("pattern").count > 0))   
 {
%>
<br>
<b><p>Nota: Los resultados de las búsquedas tienen la siguiente estructura: <br>
Nombre del módulo >> Nombre de la lección >> Título del material<br>
Si se encuentra un material y está matriculado en el módulo al cual pertenece entonces se le proveerá un enlace al mismo</p>
</b>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar" width="80%" align="center">Búsqueda por <b><%=pattern%></b></td>
    <td class="ToolBar" align="left" width="20%" >Resultado <b><%=count%></b></td>
   </tr>   
   <tr> 
       <td colspan=2 width=90%> 
         <table width="100%" border="0" cellspacing="1" cellpadding="0">
           <%=Print()%>
         </table>
       </td>
   </tr>
</table>
<%
}
else
 {pattern = "";}
 
%>

<center>
<% var tipo=0%>
<!-- #include file="../js/avsearch.inc" -->  
</center>
</body>
</html>


</HTML>
<%
      
  }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>


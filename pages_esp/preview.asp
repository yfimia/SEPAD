<%@ Language=JavaScript %>

<%
 Response.Expires = -1;
%>
<!-- #include file="../js/Adolibrary.inc" -->
<!-- #include file="../js/user.inc" -->

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

      var id = Request.QueryString.Item("id") + "";       
      
%>


<HTML>
<HEAD>
<link REL="stylesheet" TYPE="text/css" HREF="../css/<%=Session("skin")%>/Main.css" />
<link rel="stylesheet" href="../css/<%=Session("skin")%>/EvalResult.css" type="text/css">

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script language=jscript src="../js/previewejercicios.js"></script>
<script type="text/javascript" language="javascript" src="../js/Windows.js"></script>
<script type="text/javascript">
<!--

function getFullPath(path) {
  return '../Courses/course<%=Session("course")%>/' + path;

}




function putImg(path) {

  var im = new Image;
  im.src = '../Courses/course<%=Session("course")%>/' + path; 
  //alert(getFullPath(path));
  document.write('<img src=' + getFullPath(path) + ' />');  

}

function putVideo(path) {
  //var im = new Image;
  //im.dynsrc = '../Courses/course<%=Session("course")%>/' + path; 
  //document.write('<br><center><img loop=infinity dynsrc="' + getFullPath(path) + '" alt="' + getFullPath(path) + '" /></center>');  
  document.write('<br><object id="NSPlay" classid="CLSID:22D6F312-B0F6-11D0-94AB-0080C74C7E95">' + 
                 '<param name="ControlType" value="0">' + 
  		 '<param name="AutoRewind" value="1">' + 
                 '<param name="AutoSize" value="1">' + 
                 '<param name="AutoStart" value="0">' + 
                 '<param name="BaseURL" value="">' + 
                 '<param name="Filename"  value= "' + getFullPath(path) + '" >' + 
                 '<param name="ShowControls" value="1">' + 
                 '<param name="ShowAudioControls" value="1">' + 
                 '<param name="ShowPositionControls" value="1">' + 
                 '<param name="ShowTracker" value="1">' + 
                 '<param name="ShowGotoBar" value="0">' + 			 		
                 '<param name="InvokeURLs" value="1">' + 
                 '<param name="Volume" value="10">' + 
                 '<param name="AllowChangeDisplaySize" value="0">' + 
                 '</object>');
  
}

function putSound(path) {
  //var im = new Image;
  //im.dynsrc = '../Courses/course<%=Session("course")%>/' + path; 
  //document.write('<br><center><img dynsrc="' + getFullPath(path) + '"  /></center>');  
  document.write('<br><object id="NSPlay" classid="CLSID:22D6F312-B0F6-11D0-94AB-0080C74C7E95">' + 
                 '<param name="ControlType" value="0">' + 
  		 '<param name="AutoRewind" value="1">' + 
                 '<param name="AutoSize" value="1">' + 
                 '<param name="AutoStart" value="0">' + 
                 '<param name="BaseURL" value="">' + 
                 '<param name="Filename"  value= "' + getFullPath(path) + '" >' + 
                 '<param name="ShowControls" value="1">' + 
                 '<param name="ShowAudioControls" value="1">' + 
                 '<param name="ShowPositionControls" value="1">' + 
                 '<param name="ShowTracker" value="1">' + 
                 '<param name="ShowGotoBar" value="0">' + 			 		
                 '<param name="InvokeURLs" value="1">' + 
                 '<param name="Volume" value="10">' + 
                 '<param name="AllowChangeDisplaySize" value="0">' + 
                 '</object>');
  
}

function putLink(path) {
  document.write(' <a target=_blank href="' + getFullPath(path) + '"  >Descargar</a>');  
}


function goFile(path) {
  abreVentana('', 600, 400, '../Courses/course<%=Session("course")%>/' + path, 'yes');
}
//-->
</script>

</HEAD>
<body bgcolor="#FFFFFF" text="#000000" style="overflow-y:scroll">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td><img src="../images/<%=Session("skin")%>/EvalResult.gif" width="80" height="54"></td>
          <td class="HeaderTable">Respuesta</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="10" cellpadding="10">
  <tr> 
    <td>
<%

//Creacion de las preguntas

 
function setAsk()
  { 
      var  filePath = Application("dataPath");   
      var  oConn    = Server.CreateObject("ADODB.Connection");
      var  oRec     = Server.CreateObject("ADODB.Recordset");
      oConn.Open(filePath);
      oRec.Open("select ejercicios.*, ficheros.url, evaluaciones.response  from evaluaciones, ejercicios, ficheros where (evaluaciones.exercise = ejercicios.id) and(ficheros.id = ejercicios.[file]) and (evaluaciones.id = '" + id + "')",oConn,3,3);    
      if (!oRec.EOF){
          resp = oRec.Fields.Item("response").Value
          if ((resp == null)) resp = "";
          Response.Write('<SCRIPT LANGUAGE=JAVASCRIPT >response="' + resp + '"; </Script>');
          Response.Write('<SCRIPT LANGUAGE=JAVASCRIPT SRC="../Courses/Course' + Session("course") + "/" +  oRec.Fields.Item("url").Value + '"></Script>');
      }
      oRec.Close();
      oConn.Close();
  }

setAsk();

%>

<center><input type="button" onclick="history.back(-1);" value="Regresar"></center>

    </td>
  </tr>
</table>

</BODY>
</HTML>
<%
 }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>

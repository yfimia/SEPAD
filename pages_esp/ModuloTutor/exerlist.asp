<%@ Language=JavaScript %>


<%
 Response.Expires = -1;
%>
<!-- #include file='../../js/user.inc' -->
<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

 if ((Session("isTutor") > -1) || (Session("PermissionType") == ADMINISTRATOR) || (Session("userID") == Session("tutcursocordinador"))  || (Session("userID") == Request.QueryString.Item("courseOwner")))
    {
    first = "-1";
    where = "";
    if (Request.QueryString.Item("first").Count > 0) {
      first = Request.QueryString.Item("first") + "";
      
      if (first != "-1") {
        where = " and (Date < " + Application("dchar") +  first + Application("dchar") + ")";
      } 
    } 
%>


<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/Main.css" type="text/css">

<title>Calificar ejercicios pendientes</title>
</HEAD>
<body bgcolor="#FFFFFF" text="#000000">
<form id="borrar" name="borrar" action="DeleteCalif.asp?uid=<%=uid%>" method="post">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar"  align="center"><b>Ejercicios pendientes</b></td>
  </tr>
<%
    
  var oConn,oRec,filePath,HTML,fecha="";
  
  filePath = Application("dataPath"); 
  oConn    = Server.CreateObject("ADODB.Connection");
  oRec     = Server.CreateObject("ADODB.Recordset");
  
  if (Session("isTutor") > -1) {
  
    SQL      = "SELECT  Top " + SHOW_CANT + " Lecciones.Name as LName, Usuarios.ID as UID, Usuarios.FullName, Evaluaciones_Pendientes.Date, Evaluaciones_Pendientes.ID " + 
               "FROM (Lecciones INNER JOIN (Usuarios INNER JOIN Evaluaciones_Pendientes ON Usuarios.ID = Evaluaciones_Pendientes.[User]) ON Lecciones.ID = Evaluaciones_Pendientes.Lesson) INNER JOIN UsuariosSubGrupo ON Usuarios.ID = UsuariosSubGrupo.Usuario " + 
               "WHERE (((UsuariosSubGrupo.Subgrupo)=" +  Session("isTutor") +  ")) " +  where + " Order By Date Desc";
  }
  else {                
    SQL 	   = "select  Top " + SHOW_CANT + " D.Name as LName, C.ID as UID, C.FullName,A.Date,A.ID, A.Response from Evaluaciones_Pendientes A,Usuarios C, Lecciones D where (A.Lesson = D.ID) and (A.Course = " + Session("tutcurso") + ") and (A.[User] = C.ID) " +  where + " Order By Date Desc";    
  }
      
 // Response.Write(SQL);
  Code     = "var files     = new Array();";
  oRec.CursorLocation = 3;
  oRec.CursorType     = 3;

%>
  <tr> 
    <td> 
      <table width="100%" border="0" cellspacing="1" cellpadding="0">
        <tr> 
          <td width="1%" class="ToolBar" align="center"></td>
          <td width="9%" class="ToolBar" align="center"></td>		  
          <td width="30%" class="ToolBar" align="center"><b>Realizado por</b></td>
          <td width="30%" class="ToolBar" align="center"><b>Lección</b></td>
          <td width="30%" class="ToolBar" align="center"><b>Fecha</b></td>
        </tr>
        <tr> 
<%    
      var clase = "MessageTR1";
             
  i    = 0;
  fecha    = new Date();

  oConn.Open( filePath );
         
    oRec.Open(SQL,oConn,3,3);
   	last = "-1";          
    while (!oRec.EOF)
      {
        	last = oRec.Fields.Item("Date").value;          
       // fecha.setTime( Date.parse( oRec.Fields.Item("Date").value + "") );
        //date  = fecha.toLocaleString();

        if (clase == "MessageTR1") clase = "MessageTR"; else clase = "MessageTR1";

   
     
%>            
          <td align="center" width="1%" class="<%=clase%>"><input type="checkbox" id="CheckBox<%=i%>" name="CheckBox<%=i%>"><input type="hidden" id="Hidden<%=i%>" name="Hidden<%=i%>" value="<%=oRec.Fields.Item("ID").value%>"></td>
		  <td align="center" width="9%" class="<%=clase%>"><a href="exerview.asp?uid=<%=uid%>&courseOwner=<%=Request.QueryString.Item("courseOwner")%>&exer=<%=oRec.Fields.Item("ID").value%>&username=<%=oRec.Fields.Item("FullName").value%>">Calificar</a></td>
          <td align="center" width="30%" class="<%=clase%>"><a href="userDetails.asp?uid=<%=uid%>&courseOwner=<%=Request.QueryString.Item("courseOwner")%>&id=<%=oRec.Fields.Item("UID").value%>"><%=oRec.Fields.Item("FullName").value%></a></td>
          <td align="center" width="30%" class="<%=clase%>"><%=oRec.Fields.Item("LName").value%></td>
          <td nowrap align="center" width="30%" class="<%=clase%>"><%=last%></td>
              
<%              
     
       oRec.MoveNext();       
       i++;
%>
        </tr>
<%
          
      }  
      if (i > 0) 
        {
          oRec.MoveFirst();
//          Response.Write('<tr><td><INPUT id=Buton name=Buton type=Submit value=Borrar></td><td><INPUT id=Buton name=Buton type=Submit value=Regresar onclick="history.back(-1)"></td></tr>');
        }
      else
        {
%>        
          <tr><td align="center" width="100%" colspan="5" class="MessageTR">No hay ejercicios pendientes para este curso.</td></tr>        
<%          
        }  
     
%>
      </table>
    </td>
  </tr>
  <tr>
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td><input type="hidden" id="counter" name="counter" value="<%=i%>"><a href="javascript:borrar.submit()" class="ToolLink">&nbsp;Eliminar&nbsp;</a></td>
		  <td>|</td>
		  <td><a href="exerlist.asp?uid=<%=uid%>&first=<%=last%>&courseOwner=<%=Request.QueryString.Item("courseOwner")%>" class="ToolLink">&nbsp;Próximos&nbsp;<%=SHOW_CANT%>&nbsp;</a></td>
        </tr>
      </table>
    </td>
  </tr>
  
<%
    }
  else  
    {
     Response.Redirect("../errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);     
    }
%>
  
</table>

</form>      

</body>
</html>


<script language=jscript>
  var ant=0;activeNotes = new Array();
      
  function Onclick( Id )
    {     
     before.bgColor        = "silver";            
     files[Id].bgColor     = "#ffd700";              
     before	               = files[Id];           
     window.parent.frames("view").execScript("stateDisabled()","jscript");
     window.parent.frames("view").execScript("activate(" + Id + ")","jscript");
     window.parent.frames("view").execScript("stateEnabled()","jscript");
     ant                 = Id;
    }    
  
    
  function nextNote( Id )
    {
     result = -1;     
     for (i=Id + 1;i < activeNotes.length;i++)
      { if (activeNotes[i]) {result = i;i=activeNotes.length;}}
     
     if (result == -1)
      {
       for (i=0;i < Id;i++)
       if (activeNotes[i]) {result = i;i = Id;}    
      }     
     if (result == -1)  
      {window.parent.frames("view").execScript("stateDisabled()","jscript");}
     else 
      {Onclick( result )}; 
    }  
  
   function deleteNote( Id )
    {     
     activeNotes[Id]                = false;          
     files[Id].style.display        = "none";          
     nextNote( Id );     
    }
    
</script>

<%
  }
 else
   Response.Redirect("../errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
  //Response.Write(Request.QueryString.Item("uid"));
%>
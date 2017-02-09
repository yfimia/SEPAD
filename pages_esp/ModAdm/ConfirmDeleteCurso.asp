d:<%@ Language=JScript %>
<%
  Response.Expires = -1;
  Response.Buffer = true;  
%>

<!-- #include file='../../js/user.inc' -->
<!-- #include file='inc/ConfirmDeleteCurso.inc' -->

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       
  
 Session('flag') = 0;
  if ((Session("PermissionType") == ADMINISTRATOR) || (Session("userID") == Session("admcordinador")))
    {
      var SQL = Application("SQL");	 
      var  filePath = Application("dataPath");   
      var  oRec  = Server.CreateObject("ADODB.Recordset");
      var  oConn = Server.CreateObject("ADODB.Connection"); 
      var  oComm    = Server.CreateObject("ADODB.Command");
      oRec.CursorLocation = 3;
      oRec.CursorType     = 3;
      oConn.Open(filePath);
      oComm.ActiveConnection = oConn;
      
      var files = new Array();
      var i, k;



     var CID = Request.QueryString("course") + "";
     files[1] = Request.QueryString("prologue") + "";
     if ((files[1] == "undefined") || (files[1] == "null") || (files[1] == "")) k = 0; else k = 1;
     
        k = 1;  // da la cantidad de ficheros asociados 
        oComm.CommandText = "Select B.fileID,A.ID from Lecciones As A,Ficheros_de_lecciones As B Where (A.Course = " + CID + ") and (B.LessonID = A.ID)";
        oRec = oComm.Execute();
        while (oRec.EOF == false) 
          {
            k = k + 1;
            files[k] = oRec.Fields.Item("fileID").Value; 
            oRec.Move (1);
          }
        oRec.Close(); 
        //cojo los ficheros de ejercicios de las lecciones
        oComm.CommandText = "SELECT Ejercicios.[File] FROM Lecciones INNER JOIN Ejercicios ON Lecciones.ID = Ejercicios.Lesson WHERE (((Lecciones.Course)=" + CID +  "))";
        oRec = oComm.Execute();
        while (oRec.EOF == false)
          {
            k = k + 1;
            files[k] = oRec.Fields.Item("file").Value; 
            oRec.move(1);
          }  
        oRec.Close(); 
     if (SQL == '1')      
       oConn.BeginTrans();      
    
     try {                            
       //borro el curso
       oComm.CommandText = "delete  from Cursos where (ID = " + CID + ")";
       oComm.Execute();
   
       //Borro los ficheros de la base de datos
       for (i=1;i<=k;i++)
         {
           if (files[i] != "") {
            oComm.CommandText = "delete  from Ficheros where (ID = " + files[i] + ")";
            oComm.Execute();
           }
           
           //Response.Write("<font color=black>" + files[i] + "</font>&nbsp;<font color=red>Borrado</font><br>");
         }  

       
       //borro el directorio fisico del curso
       var killer = new ActiveXObject("Scripting.FileSystemObject");
       //Response.Write(killer.BuildPath("",Server.MapPath("../courses/" + "course" + CID)));
       if (killer.FolderExists(killer.BuildPath("",Server.MapPath("../../courses/" + "course" + CID))) == true)
         {
           killer.DeleteFolder(killer.BuildPath("",Server.MapPath("../../courses/" + "course" + CID)),true);
         }
       killer = null;  

       if (SQL == '1') oConn.CommitTrans();
       oConn.Close();   
      }
      catch(e) {	
      	if (SQL == '1') 
      	  oConn.RollbackTrans();
      	oConn.Close();   
      	Response.Redirect("../errorpage.asp?tipo=Error&short=Error eliminando módulo&desc=" + e.description);	
      }

      Response.Redirect("verEliminarCurso.asp?uid=" + uid);     
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</head>
<script language="jscript">
  function gogogo()
    {
     // ret.src = "../images/<%=Session("skin")%>/RegresarDn.gif";
     window.parent.selectionCourse.location = "selection.asp";
     // ret.src = "../images/<%=Session("skin")%>/RegresarUp.gif";
     window.location="CourseManager.asp";
     // window.navigate("p0.htm");
    }
</script>
<body>

<p>
<%
            
      Response.Write ('<script languaje=javascript>function pclick(){gogogo();}</script>');      
      Response.Write ("<center><input type=Button Value=" + Regresar + "onclick='return pclick()' id=Button1 name=Button1></center>" );
    }
  else
    {  
     Response.Redirect("../errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);
    }
 %>
</p>
 </body>
</html>
<%
  }
 else
   Response.Redirect("../errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
  //Response.Write(Request.QueryString.Item("uid"));
%>
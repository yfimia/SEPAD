<%@ Language=jScript %>
<%  
 Response.Expires = -1;
%>
<!-- #include file='../js/user.inc' -->
<SCRIPT language=vbscript RUNAT=Server>
 function getAtualTime
  getAtualTime = Now()
  
 End function 


</SCRIPT>

<!-- #include file='../js/adolibrary.inc' -->
<!-- #include file="../js/library.inc" -->

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

%>

<%
function UsedMem(course) {

    var dataPath = Application('dataPath');
    
    var oConn;
    var oRec;  
    oConn = MakeConnection( oConn, dataPath );
    Sql ="SELECT Top 1 Sum([Seminarios].[fileSize]) AS Total" +
         " FROM Seminarios " +
         " WHERE (((Seminarios.courseID)= " + course +  "))";

    oRec = Query( Sql, oRec, oConn  );		
    
    result = oRec.Fields.Item("Total").value;
 
    DestroyAdoObjects( oConn, oRec );
    return result;
}    
%>

<%
   var mySmartUpload;

   var mySmartUpload = Server.CreateObject("aspSmartUpload.SmartUpload");

   mySmartUpload.DeniedFilesList = "asp,php,php3,pl,dll";

   fullCourse = Session("FullCourse");

   maxFile = parseInt(fullCourse.MaxFile, 10); //kb
   maxMem = parseInt(fullCourse.MaxMem, 10);  //kb
   
   //kb           //kb         //kb  
   availMem = maxMem - UsedMem(Session("Course"));
   if (availMem >  maxFile) //Si la mem disponible es mayos que el tam permitido para un fichero, o sea si cabe otro fichero
     mySmartUpload.MaxFileSize = maxFile * 1024
   else if (availMem > 0)  mySmartUpload.MaxFileSize = availMem * 1024
        else Response.Redirect("errorpage.asp?tipo=Error&short=" + DIRECTORY_FULL_SHORT  + "&desc=" + DIRECTORY_FULL_TEXT);
   
    try {
      mySmartUpload.Upload();
    }
    catch(e) {		
      code = e.number & 0xFFFF;
      switch (code) {
        case 1010 : Response.Redirect("errorpage.asp?tipo=Error&short=" + INVALID_FILE_EXTENTION_SHORT  + "&desc=" + INVALID_FILE_EXTENTION_TEXT);
        break;
        case 1015 : Response.Redirect("errorpage.asp?tipo=Error&short=" + INVALID_FILE_EXTENTION_SHORT  + "&desc=" + INVALID_FILE_EXTENTION_TEXT);
        break;
        case 1105 : Response.Redirect("errorpage.asp?tipo=Error&short=" + INVALID_FILE_SIZE_SHORT  + "&desc=" + INVALID_FILE_SIZE_TEXT);
        break;
        case 1040 : Response.Redirect("errorpage.asp?tipo=Error&short=" + INVALID_FILE_SHORT  + "&desc=" + INVALID_FILE_TEXT);
        break;
      }
    }	
    
   fileSize = Math.round( (mySmartUpload.Files.Item("FileName").Size / 1024) * 1000) / 1000; 

   if (availMem - fileSize < maxFile)
     SendMsg( 'SEPAD', Session("UserId"), 'Directorio lleno', 'El directorio para recibir los trabajos de sus alumnos está a punto de llenarse. Tiene ' + (availMem - fileSize)  + ' de ' + maxFile + ' posibles.');
    
   if (mySmartUpload.Files.Item("FileName").Size > 0) { 

//   else error

    var dataPath = Application('dataPath');
    
    var oConn;
    var oRec;  
    oConn = MakeConnection( oConn, dataPath );
    Sql ="select * from Seminarios";

    oRec = Query( Sql, oRec, oConn  );		
    if (oRec.EOF) { id = 0 }
    else  {
      oRec.MoveLast;
      id = oRec.Fields.Item("id").value + 1; 
    }  
    
    oRec.AddNew();
    
    oRec.Fields.Item("fileTitle").value  = mySmartUpload.Files.Item("FileName").FileName;
    
    mySmartUpload.Files.Item("FileName").SaveAs("../Courses/Course" + Session("course") + "/Seminarios/" +  id + '_' + mySmartUpload.Files.Item("FileName").FileName);
    


    oRec.Fields.Item("fileSize").value  = fileSize;
    oRec.Fields.Item("fileDescription").value  = mySmartUpload.Form.Item('FileDescription').Values;
    oRec.Fields.Item("fileURL").value  = id + '_' + mySmartUpload.Files.Item("FileName").FileName;
    oRec.Fields.Item("uploadDate").value      = getAtualTime();
    oRec.Fields.Item("userID").value  = Session("userID");
    oRec.Fields.Item("courseID").value  = Session("course");          
    oRec.Update();      
 
    DestroyAdoObjects( oConn, oRec );

    Response.Redirect('UpLoadFile.asp?uid=' + uid);        
  }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + INVALID_FILE_SHORT  + "&desc=" + INVALID_FILE_TEXT);

%>

<%
  }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>

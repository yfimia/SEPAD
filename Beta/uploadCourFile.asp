<%@ Language=JavaScript %>

<%
  Response.Expires = 0;
%>
<!-- #include file='../js/user.inc' -->
<%
  var unPacker = Server.CreateObject("SEPADRW.Utils");


//  var oXmlDoc = Server.CreateObject("MICROSOFT.XMLDOM");
//  oXmlDoc.async = false; 
//  oXmlDoc.load(Request);

  var SQL;
  var dataPath;
  var PerType = -1;
  var  oRec  = Server.CreateObject("ADODB.Recordset");
  var  oConn = Server.CreateObject("ADODB.Connection"); 


function GetDSN() {
    place = "1";
/*    var sXml = "../Xml/config.xml";
    var XmlDoc = Server.CreateObject("MICROSOFT.XMLDOM");
    XmlDoc.async = false; 
    XmlDoc.load(Server.MapPath(sXml));


    //DBConn
    SQL  = XmlDoc.documentElement.selectSingleNode("/SysConfig/DBConnection/SQL").text;
    dataPath = XmlDoc.documentElement.selectSingleNode("/SysConfig/DBConnection/DSN").text;
    XmlDoc = null;
  */
  
    SQL  = Application("SQL");
    dataPath = Application("datapath");
    
}

function CheckUser(user, password, courseID) {
   place = "2";
   var res = -1;

   oRec.Open("select * from Usuarios where ([name] = '" + user + "')",oConn,3,3);
   if (!(oRec.EOF)) {
     if (oRec.Fields.Item("Password").Value + "" == password) {
       res = oRec.Fields.Item("ID").Value;
       PerType = oRec.Fields.Item("PermissionType").Value;
     }  
   }   
   oRec.Close();
   
   if (res != -1) {
     
     oRec.Open("select Cursos.ID FROM Cursos INNER JOIN Modulo ON Modulo.ID = Cursos.modulo WHERE (Cursos.ID = " + courseID  + ") and ((owner = " + res + ") or (Modulo.cordinador = " + res + "))",oConn,3,3);
     if ((oRec.EOF)) {
         res = -1;
     }   
     oRec.Close();
    
   }
   return res;
}


function IsUserInGroup(userid, groupname) {
    var Sql = " SELECT Grupos_de_Usuarios.ID " + 
              " FROM Grupos INNER JOIN Grupos_de_Usuarios ON Grupos.ID = Grupos_de_Usuarios.[Group] " +
              " WHERE (((Grupos_de_Usuarios.[User])=" + userid + ") AND ((Grupos.Name)='" + groupname + "'))";

    oRec.Open(Sql,oConn,3,3);
    var res =  !(oRec.EOF); 
//    Response.Write(Sql);
    oRec.Close();
    return(res);
}


 var error = false;
 xml = '<?xml version="1.0" encoding="iso-8859-1"?><FileResult>';
 
 try {
   GetDSN();
   oConn.Open( dataPath );
 }  
 catch(e) {		
     error = true;
     xml = xml + '<ErrorNo>-1</ErrorNo>';
     xml = xml + '<uid>' + Session("uid") + '</uid>';
     xml = xml + '<Error>' + e.description +  '</Error>';
 }
 	


   var mySmartUpload = Server.CreateObject("aspSmartUpload.SmartUpload");

   mySmartUpload.DeniedFilesList = "asp,php,php3,pl,dll";

   //maxFile = parseInt(fullCourse.MaxFile, 10); //kb
   //maxMem = parseInt(fullCourse.MaxMem, 10);  //kb
   //kb           //kb         //kb  
   //availMem = maxMem - UsedMem(Session("Course"));
   //if (availMem >  maxFile) //Si la mem disponible es mayos que el tam permitido para un fichero, o sea si cabe otro fichero
   //  mySmartUpload.MaxFileSize = maxFile * 1024
   //else if (availMem > 0)  mySmartUpload.MaxFileSize = availMem * 1024
   //     else Response.Redirect("errorpage.asp?tipo=Error&short=" + DIRECTORY_FULL_SHORT  + "&desc=" + DIRECTORY_FULL_TEXT);

    try {
      mySmartUpload.Upload();
      
    }
    catch(e) {		
      code = e.number & 0xFFFF;
      switch (code) {
        case 1010 :  xml = xml + '<ErrorNo>-1</ErrorNo>';
          xml = xml + '<uid>' + Session("uid") + '</uid>';
          xml = xml + '<Error>' + INVALID_FILE_EXTENTION_TEXT +  '</Error>';

        //Response.Redirect("errorpage.asp?tipo=Error&short=" + INVALID_FILE_EXTENTION_SHORT  + "&desc=" + INVALID_FILE_EXTENTION_TEXT);
        break;
        case 1015 : xml = xml + '<ErrorNo>-1</ErrorNo>';
          xml = xml + '<uid>' + Session("uid") + '</uid>';
          xml = xml + '<Error>' + INVALID_FILE_EXTENTION_TEXT +  '</Error>';
          //Response.Redirect("errorpage.asp?tipo=Error&short=" + INVALID_FILE_EXTENTION_SHORT  + "&desc=" + INVALID_FILE_EXTENTION_TEXT);
        break;
        case 1105 : xml = xml + '<ErrorNo>-1</ErrorNo>';
          xml = xml + '<uid>' + Session("uid") + '</uid>';
          xml = xml + '<Error>' + INVALID_FILE_SIZE_TEXT +  '</Error>';
          //Response.Redirect("errorpage.asp?tipo=Error&short=" + INVALID_FILE_SIZE_SHORT  + "&desc=" + INVALID_FILE_SIZE_TEXT);
        break;
        case 1040 : 
          xml = xml + '<ErrorNo>-1</ErrorNo>';
          xml = xml + '<uid>' + Session("uid") + '</uid>';
          xml = xml + '<Error>' + INVALID_FILE_TEXT +  '</Error>';
          //Response.Redirect("errorpage.asp?tipo=Error&short=" + INVALID_FILE_SHORT  + "&desc=" + INVALID_FILE_TEXT);
        break;
      }
    }	

   courseId = mySmartUpload.Form.Item('CourseID').Values;



  try {	 
     var courseName = mySmartUpload.Form.Item('CourseName').Values; 
     var owner = CheckUser(mySmartUpload.Form.Item('User').Values, mySmartUpload.Form.Item('Password').Values, courseId); 	
     
     if ( (owner != -1) || (PerType == ADMINISTRATOR) ) {
     
       if (mySmartUpload.Files.Item("FileName").Size > 0) { 
          //xml = xml + '<uid1>' + Server.MapPath('../Courses') + '\\' + mySmartUpload.Form.Item('Dir').Values + mySmartUpload.Files.Item("FileName").FileName + '</uid1>';

          unPacker.CreateDirs(mySmartUpload.Form.Item('Dir').Values, Server.MapPath('../Courses') + '\\');
          mySmartUpload.Files.Item("FileName").SaveAs(Server.MapPath('../Courses') + '\\' + mySmartUpload.Form.Item('Dir').Values + mySmartUpload.Files.Item("FileName").FileName);
          xml = xml + '<ErrorNo>1</ErrorNo>';
          xml = xml + '<uid>' + Session("uid") + '</uid>';
          
       }
       else {
          xml = xml + '<ErrorNo>-1</ErrorNo>';
          xml = xml + '<uid>' + Session("uid") + '</uid>';
          xml = xml + '<Error>' + INVALID_FILE_TEXT +  '</Error>';
       }	
     
//       unPacker.UnZip(Server.MapPath('../Courses') + '\\' +  oXmlDoc.selectSingleNode('File/Dir').text, oXmlDoc.selectSingleNode('File/Name').text, oXmlDoc.selectSingleNode('File/Data').text);
     }
     else  {
       xml = xml + '<ErrorNo>-2</ErrorNo>';
       xml = xml + '<uid>' + Session("uid") + '</uid>';
       xml = xml + '<Error>Error de auntentificación. El usuario especificado no tiene suficientes permisos para subir archivos al módulo. Contacte con el coordinador de la módalidad académica.</Error>';
     } 
  
   }   
   catch(e) {		
     xml = xml + '<ErrorNo>-1</ErrorNo>';
     xml = xml + '<uid>' + Session("uid") + '</uid>';
     xml = xml + '<Error>' + e.description +  '</Error>';
   }

  xml = xml + '</FileResult>';
  unPacker = null; 

 Response.ContentType = "text/Xml" ;
 Response.Charset = "iso-8859-1";
 Response.Write(xml);
%>





<%@ Language=JavaScript %>

<%
  Response.Expires = 0;
%>
<!-- #include file='../js/user.inc' -->
<%

  SELECTION  = 1;
  JOIN       = 2;		
  COMPLETE   = 3;
  SUPERVISED = 6;
  
  var place = "-1";
  
  var SQL;
  var dataPath;
  var ExersId = new Array();
  var ExerLessonId = new Array();
  var ExerIndex = new Array();
  var exerCant = 0;	
  var xml = '';
  var PerType = -1;
  
  
  try {
     var unPacker = Server.CreateObject("SEPADRW.Utils");
  }   
  catch(e) {		
     xml = '<Public>';
     xml = xml + '<CourseId>-3</CourseId>';
     xml = xml + '<Error>El objeto SEPADRW no está registrado.</Error>';
     xml = xml + '</Public>';
     
  }

  var  oRec  = Server.CreateObject("ADODB.Recordset");
  var  oConn = Server.CreateObject("ADODB.Connection"); 
  var  oComm    = Server.CreateObject("ADODB.Command");

  var oXmlDoc = Server.CreateObject("MICROSOFT.XMLDOM");
  oXmlDoc.async = false; 
  
  
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
   var  oRec  = Server.CreateObject("ADODB.Recordset");

   place = "2";
   var res = -1;

   oRec.Open("select * from Usuarios where (name = '" + user + "')",oConn,3,3);
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


function CreateFile(title, desc, url, size) {
   var  oRec  = Server.CreateObject("ADODB.Recordset");
//Response.Write(title + " ** " +  desc + " ** " + url + " ** " + size + "<br>");
   place = place  + title ;
   var res = -1;
   	
   oRec.Open("SELECT * From Ficheros",oConn,3,3);
   oRec.AddNew();
   oRec.Fields.Item("url").Value = url;
   oRec.Fields.Item("tama").Value = size;
   oRec.Fields.Item("title").Value = title;
   oRec.Fields.Item("description").Value = desc;
   oRec.Update();      
   oRec.MoveLast();
   res = oRec.Fields.Item("ID").Value;

   oRec.Close();
//Response.Write("ok<br>");
   
   return res;	
}

function ModifyFile(fileID, title, desc, url, size) {
   var  oRec1  = Server.CreateObject("ADODB.Recordset");
   //Response.Write(fileID + "**" + title + " ** " +  desc + " ** " + url + " ** " + size + "<br>");
   place = "3.1";
   var res = -1;
   	
   oRec1.Open("SELECT * From Ficheros Where ID =" + fileID,oConn,3,3);
   if (!oRec1.EOF) {
     //oRec1.MoveLast();	
     oRec1.Fields.Item("url").Value = url;
     oRec1.Fields.Item("tama").Value = size;
     oRec1.Fields.Item("title").Value = title;
     oRec1.Fields.Item("description").Value = desc;
     oRec1.Update();      
   }  
   
   oRec1.Close();
   oRec1 = null;
   
//Response.Write("ok<br>");
   
   return res;	
}

function AddFileToLesson(fileId, lessonId) {
   place = "4 " + fileId + " , " + lessonId;
   var res = -1;
   	
   oRec.Open("SELECT * From Ficheros_de_Lecciones",oConn,3,3);
   
   oRec.AddNew();
   oRec.Fields.Item("fileID").Value = fileId;
   oRec.Fields.Item("lessonID").Value = lessonId;
   res = oRec.Fields.Item("ID").Value;
   oRec.Update();      
   //oRec.MoveLast();

   oRec.Close();
   return res;	
}

 
function PublicExers(exerNode, courseID, lessonId, exerIndex) {
  place = "5";
  var fileId;
  //Response.Write("ejer<br>");


  var exerType = parseInt(exerNode.selectSingleNode('Tipo').text, 10);	
  cant = 0;
  cad1 = '';  
  switch (exerType) {
    case  SELECTION: 
        var optionList = exerNode.selectNodes('Elementos/Opcion');	
        cant = optionList.length;
        cad1 = '';
        cad2 = '';
        for(var i = 0; i < optionList.length; i++) {
            if (optionList[i].selectSingleNode('Respuesta').text == '1' ) temp = 'T'; else temp = 'F';
            cad1 = cad1 + temp;
            if (i != optionList.length - 1) cad1 = cad1 + ';';
            
            cad2 = cad2 + optionList[i].selectSingleNode('Puntos').text;
            if (i != optionList.length - 1) cad2 = cad2 + ';';
        }
        cad1 = cad1 + '&' + cad2;
    break; 	  

    case  JOIN: 
        var optionList = exerNode.selectNodes('ColumnaB/Fila/Opcion');	
        cant = optionList.length;
        cad1 = '';
        cad2 = '';
        for(var i = 0; i < optionList.length; i++) {
            cad1 = cad1 + optionList[i].selectSingleNode('Respuesta').text;
            if (i != optionList.length - 1) cad1 = cad1 + ',';
            
            cad2 = cad2 + optionList[i].selectSingleNode('Puntos').text;
            if (i != optionList.length - 1) cad2 = cad2 + ',';
        }
        cad1 = cad1 + '&' + cad2;      
    break; 	  

    case  COMPLETE: 
        var compOptionList = exerNode.selectNodes('Elementos/Completar_Opcion');	
        cant = compOptionList.length;        
        var optionList;   
        cad1 = '';
        cad2 = '';
        for(var i = 0; i < compOptionList.length; i++) {
          optionList = compOptionList[i].selectNodes('Opcion');	
          for(var j = 0; j < optionList.length; j++) {
             cad1 = cad1 + optionList[j].selectSingleNode('Respuesta').text;
            if (j != optionList.length - 1) cad1 = cad1 + ',';
            
            cad2 = cad2 + optionList[j].selectSingleNode('Puntos').text;
            if (j != optionList.length - 1) cad2 = cad2 + ',';
          }
          cad1 = cad1 + ';';  
          cad2 = cad2 + ';';  
          
        }
        cad1 = cad1 + '&' + cad2;            
    break; 	  

    case  SUPERVISED: 
        cad1 = ' ';  
        cant = 1;
    break; 	  
  }
  //Response.Write("<br>" + lessonId + " *** " + fileId + " *** " + exerType +  " *** " + cad1  );

  var fileId;

  oRec.Open("SELECT * From Ejercicios WHERE ([key] = '" + exerNode.selectSingleNode('Llave').text + "') and (Lesson = " + lessonId + ") ",oConn,3,3);  
  if (oRec.EOF) { //No existe,... crear uno nuevo                                                             
     fileId = CreateFile('', '', 'e' + exerIndex + 'l' + lessonId + 'c' + courseID + '.js', 0);	 

     oRec.AddNew();
     oRec.Fields.Item("key").Value = exerNode.selectSingleNode('Llave').text;
     
     oRec.Fields.Item("lesson").Value = lessonId;
     oRec.Fields.Item("file").Value = fileId;
     oRec.Fields.Item("type").Value = exerType;
     oRec.Fields.Item("cant").Value = cant;
     oRec.Fields.Item("response").Value = cad1;
     oRec.Update();      
     oRec.MoveLast();
     res = oRec.Fields.Item("ID").Value;
     oRec.Close();
  } 
  else {//Ya existe,... modificarlo
     ModifyFile(oRec.Fields.Item("file").Value, '', '', 'e' + exerIndex + 'l' + lessonId + 'c' + courseID + '.js', 0);	 
       	
     oRec.Fields.Item("lesson").Value = lessonId;
     //oRec.Fields.Item("file").Value = fileId;
     fileId = oRec.Fields.Item("file").Value;
     oRec.Fields.Item("type").Value = exerType;
     oRec.Fields.Item("cant").Value = cant;
     oRec.Fields.Item("response").Value = cad1;
     res = oRec.Fields.Item("ID").Value;
     oRec.Update();      
     oRec.Close();
     
  }   
//  Response.Write("OKejer<br>");
  
  return res;	
}

function MaterialExistsInLesson(surl, stitle, lesson) {
   var  oRec  = Server.CreateObject("ADODB.Recordset");
   var result = -1;
   var sSQL = "SELECT Ficheros.Id AS FID FROM Ficheros INNER JOIN Ficheros_de_lecciones ON Ficheros.Id = Ficheros_de_lecciones.fileID " +
             "WHERE (Ficheros_de_lecciones.lessonID = " + lesson + ") AND " + 
             "(Ficheros.url = '" + surl + "') AND (Ficheros.title = '" + stitle + "')";
   place = "6.1.1 " +  sSQL;             
   oRec.Open(sSQL, oConn,3,3);
   place = "6.1.2 " +  sSQL;
   if (!oRec.EOF) { 
     result = oRec.Fields.Item("FID").Value;
   }
   oRec.Close();
   oRec = null;
   return result;
}

function PublicLessonExers(lessonNode, courseID, order) {
   place = "6";
   var lessonId = -1;

   oRec.Open("SELECT * From Lecciones where (Name = '" + lessonNode.selectSingleNode('Nombre').text + "') and (course = " + courseID + ")",oConn,3,3);
   if (oRec.EOF) { //Es una nueva...
   
   
     oRec.AddNew();
     oRec.Fields.Item("name").Value = lessonNode.selectSingleNode('Nombre').text;
     oRec.Fields.Item("dir").Value = lessonNode.selectSingleNode('Directorio').text;
     oRec.Fields.Item("course").Value = courseID;
     oRec.Fields.Item("primary").Value = (lessonNode.selectSingleNode('Primaria').text == '1') ? true : false;
     oRec.Fields.Item("nivel").Value = lessonNode.selectSingleNode('Nivel').text;
     oRec.Fields.Item("order").Value = order;
     oRec.Fields.Item("state").Value = 1;
     
     oRec.Update();      
     oRec.MoveLast();
     lessonId = oRec.Fields.Item("ID").Value;
   }
   else { //ya existe
     oRec.Fields.Item("name").Value = lessonNode.selectSingleNode('Nombre').text;
     oRec.Fields.Item("dir").Value = lessonNode.selectSingleNode('Directorio').text;
     oRec.Fields.Item("course").Value = courseID;
     oRec.Fields.Item("primary").Value = (lessonNode.selectSingleNode('Primaria').text == '1') ? true : false;
     oRec.Fields.Item("nivel").Value = lessonNode.selectSingleNode('Nivel').text;
     oRec.Fields.Item("order").Value = order;
     
     oRec.Update();      
     oRec.MoveLast();
     lessonId = oRec.Fields.Item("ID").Value;
   }  
   oRec.Close();

   place = "6.1" + lessonId; 
   var cond = "and (NOT (Ficheros.ID  IN (";
   var matList = lessonNode.selectNodes('Materiales/Material');
   var fileId = -1;
   for(i = 0; i < matList.length; i++) {
     place = "6.3";
     fileId = MaterialExistsInLesson(matList[i].selectSingleNode('URL').text, matList[i].selectSingleNode('Titulo').text, lessonId);
     if (fileId == -1) {
       fileId = CreateFile( matList[i].selectSingleNode('Titulo').text,  matList[i].selectSingleNode('Descripcion').text,  matList[i].selectSingleNode('URL').text,  matList[i].selectSingleNode('SIZE').text);	 
       AddFileToLesson(fileId, lessonId);
     }  
     else  {
       ModifyFile(fileId, matList[i].selectSingleNode('Titulo').text,  matList[i].selectSingleNode('Descripcion').text,  matList[i].selectSingleNode('URL').text,  matList[i].selectSingleNode('SIZE').text);
     }  
      

     if (i < matList.length - 1)
         cond = cond + fileId + ",";
     else  
       cond = cond + fileId;
   }
   cond = cond + ")))";
   if (matList.length == 0) cond = "";
   
   place = "6.4" + "DELETE From Ficheros where (ID IN (SELECT [fileID] From Ficheros_de_Lecciones where lessonID = " + lessonId + ")) " + cond;
   oComm.CommandText = "DELETE From Ficheros where (ID IN (SELECT [fileID] From Ficheros_de_Lecciones where lessonID = " + lessonId + ")) " + cond;
   oComm.Execute();
   

   var cond_Evaluaciones = "(NOT (Evaluaciones.Exercise  IN (" ;
   var cond_Ejercicios_Pendientes = "(NOT (Evaluaciones_Pendientes.Exercise  IN (";
   var cond_Ejercicios = "(NOT (Ejercicios.ID  IN (";
   var cond = '';
   var exerList = lessonNode.selectNodes('Ejercicios/Ejercicio');
   for(var i = 0; i < exerList.length; i++) {
   
   
      ExersId[exerCant] = PublicExers(exerList[i], courseID, lessonId, i);	   	
     
     
      if (i < exerList.length - 1)
         cond = cond + ExersId[exerCant] + ",";
       else  
         cond = cond + ExersId[exerCant];
      ExerLessonId[exerCant] = lessonId;
      ExerIndex[exerCant] = i;
      exerCant++;	
   }
   if (exerList.length == 0) {
      cond_Evaluaciones = '';
      cond_Ejercicios_Pendientes = '';
      cond_Ejercicios = '';
   }   
   else {
      cond_Evaluaciones = cond_Evaluaciones + cond + "))) " + " and ";
      cond_Ejercicios_Pendientes = cond_Ejercicios_Pendientes + cond + "))) " + " and ";
      cond_Ejercicios = cond_Ejercicios + cond + "))) " + " and ";
   }

   place = "5.1" + " Select [File] as FID FROM Ejercicios WHERE " + cond_Ejercicios + "  (Lesson = " + lessonId + ")";
   var cond1 = "";
   oRec.Open("Select [File] as FID FROM Ejercicios WHERE " + cond_Ejercicios + "  (Lesson = " + lessonId + ")",oConn,3,3);
   if (!oRec.EOF) { 
      cond1 = cond1 + oRec.Fields.Item("FID").Value;
      oRec.Move(1);
   }   
   while (!oRec.EOF) { 
     cond1 = cond1 + "," + oRec.Fields.Item("FID").Value ;
     oRec.Move(1);
   } 
   oRec.Close();
   place = "5.2";

   oComm.CommandText = "DELETE FROM Evaluaciones WHERE " + cond_Evaluaciones + " (Lesson = " + lessonId + ")";
   oComm.Execute();

   place = "5.3" + "DELETE FROM Evaluaciones_Pendientes WHERE " + cond_Ejercicios_Pendientes + " (Lesson = " + lessonId + ")";
   oComm.CommandText = "DELETE FROM Evaluaciones_Pendientes WHERE " + cond_Ejercicios_Pendientes + " (Lesson = " + lessonId + ")";
   oComm.Execute();

   place = "5.4";
   oComm.CommandText = "DELETE FROM Ejercicios WHERE " + cond_Ejercicios + " (Lesson = " + lessonId + ")";
   oComm.Execute();

   
   if (cond1 != "") {
     place = "5.5" + "DELETE FROM Ficheros WHERE (Ficheros.ID IN (" + cond1 + ") )";
     oComm.CommandText = "DELETE FROM Ficheros WHERE (Ficheros.ID IN (" + cond1 + ") )";
     oComm.Execute();  
   }  
   

  //Response.Write("lessOK<br>");
   
   return lessonId;	
}

function GetLessonId(lessonName, courseID) {
   place = "7";
   var res = -1;
   	
   oRec.Open("select id from Lecciones where ((name = '" +  lessonName  +  "') and (course = " + courseID + "))",oConn,3,3);
   if (!(oRec.EOF)) res = oRec.Fields.Item("ID").Value;

   oRec.Close();
   return res;	
}

function PublicLessonRels(lessonNode, courseID, lessonId) {
   place = "8";
   oComm.CommandText = "DELETE From Proxima_Leccion WHERE ID IN (SELECT [linked] From Relaciones_del_Curso WHERE Relaciones_del_Curso.Lesson=" + lessonId + ")";
   oComm.Execute();

   oComm.CommandText = "DELETE FROM Relaciones_del_Curso WHERE Relaciones_del_Curso.Lesson=" + lessonId;
   oComm.Execute();
   
   
   var relLessonId;
   var tmpLessonId;
   var relList = lessonNode.selectNodes('Relaciones/Leccion_Relacionada');
   for(var i = 0; i < relList.length; i++) {
     tmpLessonId = GetLessonId(relList[i].text, courseID); 	
     //Response.Write(tmpLessonId + " -- " + relList[i].text + "<br>");
     oRec.Open("SELECT * From Proxima_Leccion",oConn,3,3);
     oRec.AddNew();
     oRec.Fields.Item("lesson").Value = tmpLessonId;
     oRec.Update();      
     oRec.MoveLast();
     relLessonId = oRec.Fields.Item("ID").Value;
     oRec.Close();
     
     oRec.Open("SELECT * From Relaciones_del_Curso",oConn,3,3);
     oRec.AddNew();
     oRec.Fields.Item("lesson").Value = lessonId;
     oRec.Fields.Item("linked").Value = relLessonId;
     oRec.Update();      
     oRec.Close();
   }
}

function ExistCourse(courseName) {
   place = "9";
   var res = -1;	  
   oRec.Open("select * from Cursos where (name = '" + courseName + "')" ,oConn,3,3);   	
   if (!(oRec.EOF)) res = oRec.Fields.Item("ID").Value;
   oRec.Close();
   return res;
}

var x;

function PublicCourse() {
   place = "10";
   var courseId;
   var courseName = oXmlDoc.selectSingleNode('Curso/Nombre').text;	   
   try {
     courseId = oXmlDoc.selectSingleNode('Curso/Codigo').text + "";	   
   }  
   catch(e) {		
     return "-4";
   }	

      try {	 
        
        oRec.Open("select * from Cursos where ID =" + courseId,oConn,3,3);   	
        if (oRec.EOF) {
          //No existe un curso con ese codigo
          courseId = "-2";
        }
        else {

          var owner = CheckUser(oXmlDoc.selectSingleNode('Curso/User').text, oXmlDoc.selectSingleNode('Curso/Password').text, courseId); 	
          if ( (owner != -1) || (PerType == ADMINISTRATOR) ) {
        
           var prolId = CreateFile('', '', oXmlDoc.selectSingleNode('Curso/Prologo').text, 0);
           
            oRec.Fields.Item("Name").Value = courseName;
            var oldprol = oRec.Fields.Item("Prologue").Value + "";
            oRec.Fields.Item("Prologue").Value = prolId;
            oRec.Update();      
            oRec.MoveLast();
            oRec.Close();


            place = "11";
            try {
              oComm.CommandText = "DELETE From Ficheros where ID = " + oldprol;
              oComm.Execute();
            }    
            catch(e) {		
            }	
            

            place = "12";
            var lessonIds = new Array();
            var lessonList = oXmlDoc.selectNodes('Curso/Lecciones/Leccion');
            
            var cond = "(NOT (ID IN (";
            
            for(var i = 0; i < lessonList.length; i++) {
              lessonIds[i] = PublicLessonExers(lessonList[i], courseId, i);	 

              
              if (i < lessonList.length - 1)
                cond = cond + lessonIds[i] + ",";
              else  
                cond = cond + lessonIds[i];
            } 
            place = "12.1";
            cond = cond + "))) and ";
            if (lessonList.length == 0) cond = "";
            x = oComm.CommandText;
            oComm.CommandText = "DELETE FROM Lecciones WHERE " + cond + "(course = " + courseId + ")";
            oComm.Execute();

            place = "13";
            
            for(var i = 0; i < lessonList.length; i++) {
              PublicLessonRels(lessonList[i], courseId, lessonIds[i]);	 
              
              
            }
          }  
          else        
            return "-1";
            
        }  
        return courseId;	
      } 
      catch(e) {		
        if (SQL == '1') oConn.RollbackTrans();
      }	
 
 }




 var cid = "-3";
 
 xml = '<Public>';
  try {	 
 
   oXmlDoc.load(Request);
   //oXmlDoc.save('c:\\temp\\x.xml');
   //oXmlDoc.load(Server.MapPath("../a.xml"));
 

   GetDSN();
   oConn.Open( dataPath );
   oComm.ActiveConnection = oConn;
  
   if (SQL == '1')
     oConn.BeginTrans();

   cid = PublicCourse();	  
   if ((cid != "-1") && (cid != "-2") && (cid != "-4"))
     unPacker.CreateDirs('course' + cid + '\\Seminarios\\', Server.MapPath('../Courses') + '\\');
   unPacker = null;


   if (SQL == '1') oConn.CommitTrans();
   oConn.Close();

     
      
   //Response.Write(Server.MapPath('../Courses'));
   xml = xml + '<CourseId>' +  cid + '</CourseId>';
      
   if (cid != "-4")  {
      xml = xml + '<Exercises>';
    
      for(i = 0; i < exerCant; i++) {
        xml = xml + '<Exercise>';
        xml = xml + '  <ExerId>';
        xml = xml +      ExersId[i];
        xml = xml + '  </ExerId>';
        xml = xml + '  <LessonId>';
        xml = xml +      ExerLessonId[i];
        xml = xml + '  </LessonId>';
        xml = xml + '  <ExerIndex>';
        xml = xml +      ExerIndex[i];
        xml = xml + '  </ExerIndex>';
        xml = xml + '</Exercise>';
      }
      xml = xml + '</Exercises>';
   } 
   else {
     xml = xml + '<ErrorNo>-4</ErrorNo>';
     xml = xml + '<Error>Error, código de módulo no válido o Herramienta de Publicación no compatible con la versión actual del servidor. Para publicar en este servidor se requiere SEPADHP1.5</Error>';
   }  
 
}   
catch(e) {		
   xml = xml + '<CourseId>-3</CourseId>';
   xml = xml + '<Error>Lugar:' + place + ' Descripción:' + e.description + '</Error>';
}



 xml = xml +  '</Public>';
 

 Response.ContentType = "text/Xml" ;
 Response.Charset = "iso-8859-1";
 //Response.Write(place);
 Response.Write(xml);
%>





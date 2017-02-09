<%
   var lessons = new Array();       
   var courses = new Array();                                              


   
   function GetLessonIndex(lessonID) {
     var filePath = Application('dataPath');
     var oConn;
     var oRec;
     var result = -1;
      
     oConn = MakeConnection( oConn, filePath );
     Sql = "SELECT [order] Top 1 FROM Lecciones WHERE (ID " + lessonID + ") order by [order] Asc"; 

                                                                              
     oRec = Query( Sql, oRec, oConn  );		                          
     if (!oRec.EOF)
       {
        result = oRec.Fields.Item("orderID").Value;
       } 
      DestroyAdoObjects( oConn, oRec );
      return result;	       	  
   }

   function Course(ID,Name,owner, state)
     {this.ID = ID;this.state = state;this.Name = Name;this.owner = owner;}
     
/*
      Session("tutcurso") = Request.QueryString.Item("course") + "";       
      Session("tutcursoName") = Request.QueryString.Item("courseName") + "";       
      Session("tutcursoOwner") = Request.QueryString.Item("courseOwner") + "";       
      Session("tutcursomodulo")        =      Session("modulo");    
      Session("tutcursomoduloName")    =      Session("moduloName");
      Session("tutcursostate")         =      Session("state");     
      Session("tutcursocordinador")    =      Session("cordinador");
      Session("tutcursogrupo")         =      Session("grupo");     
      Session("tutcursoclaustro")      =      Session("claustro");  
*/
   function AdmCourse(ID,Name,owner, state, Modulo, ModuloName, ModuloState, Cordinador, grupo, claustro, MaxFile, MaxMem)
     {this.ID = ID;this.state = state;this.Name = Name;this.owner = owner;
      this.ID = ID;this.Modulo = Modulo;this.ModuloName = ModuloName;this.ModuloState = ModuloState; 
      this.Cordinador = Cordinador;this.grupo = grupo; this.claustro = claustro; this.MaxFile = MaxFile; this.MaxMem = MaxMem;
     }

   function GetAdmCourse( course ) {
     var filePath = Application('dataPath');
     var oConn;
     var oRec;
     var result;
      
     oConn = MakeConnection( oConn, filePath );
     Sql = "SELECT Cursos.Name as CName, Cursos.maxFile, Cursos.maxMem, Cursos.state as CState, Cursos.owner, Cursos.ID as CID, Modulo.cordinador, Modulo.state as MState, Modulo.grupo, Modulo.claustro, Modulo.ID as MID, Modulo.Name  as MName "  + 
           "FROM Cursos, Modulo " + 
           "WHERE (Cursos.modulo = Modulo.ID) and (Cursos.ID=" + course + ")";

//Response.Redirect("../errorpage.asp?tipo=Error&short=" + Sql  + "&desc=" + SESSION_TIMEOUT_TEXT);
                                                                              
     oRec = Query( Sql, oRec, oConn  );		                          
     if (!oRec.EOF)
       {
        var ID               = oRec.Fields.Item("CID").Value;
        var Name		= oRec.Fields.Item("CName").Value;
        var state            = oRec.Fields.Item("CState").Value;
        var owner            = oRec.Fields.Item("owner").Value;
        var Modulo         = GetPrologue(oRec.Fields.Item("MID").Value);
        var ModuloName               = oRec.Fields.Item("MName").Value;
        var MaxFile            = oRec.Fields.Item("MaxFile").Value;
        var MaxMem            = oRec.Fields.Item("MaxMem").Value;
        var ModuloState            = oRec.Fields.Item("MState").Value;
        var Cordinador            = oRec.Fields.Item("cordinador").Value;
        var grupo            = oRec.Fields.Item("grupo").Value;
        var claustro         = oRec.Fields.Item("claustro").Value;
   
        result =  new AdmCourse(ID,Name,owner, state, Modulo, ModuloName, ModuloState, Cordinador, grupo, claustro, MaxFile, MaxMem);
       } 
      DestroyAdoObjects( oConn, oRec );
      return result;	       	  
   }



   function AdmModulo(ID,Name, state, Cordinador, grupo, claustro)
     {this.ID = ID;this.state = state;this.Name = Name;
      this.Cordinador = Cordinador;this.grupo = grupo; this.claustro = claustro; 
     }

   function GetAdmModulo( modulo ) {
     var filePath = Application('dataPath');
     var oConn;
     var oRec;
     var result;
      
     oConn = MakeConnection( oConn, filePath );
     Sql = "SELECT Modulo.cordinador, Modulo.state as MState, Modulo.grupo, Modulo.claustro, Modulo.ID as MID, Modulo.Name  as MName "  + 
           "FROM Modulo " + 
           "WHERE  (Modulo.ID=" + modulo + ")";


                                                                              
     oRec = Query( Sql, oRec, oConn  );		                          
     if (!oRec.EOF)
       {
        var ID               = oRec.Fields.Item("MID").Value;
        var Name		= oRec.Fields.Item("MName").Value;
        var state            = oRec.Fields.Item("MState").Value;
        var Cordinador            = oRec.Fields.Item("cordinador").Value;
        var grupo            = oRec.Fields.Item("grupo").Value;
        var claustro         = oRec.Fields.Item("claustro").Value;
        result =  new AdmModulo(ID,Name, state, Cordinador, grupo, claustro);
        //Response.Redirect("../errorpage.asp?tipo=Error&short=" + result.ID + " " + result.Name + " " +  result.state + " " +  result.Cordinador + " " +  result.grupo + " " +  result.claustro  + "&desc=" + SESSION_TIMEOUT_TEXT);   

       } 
      DestroyAdoObjects( oConn, oRec );
      return result;	       	  
   }



   function FullCourse(ID,Name,owner, state, prologue, Atmost, ctdExAct, ctdExRet, MaxFile, MaxMem, lastLessonID, firstLessonID)
     {
      this.ID = ID;this.state = state;this.Name = Name;this.owner = owner;
      this.Atmost = Atmost; this.ctdExAct = ctdExAct; this.ctdExRet = ctdExRet; this.MaxFile = MaxFile; this.MaxMem = MaxMem;
      this.prologue = prologue; this.lastLessonID = lastLessonID; this.firstLessonID = firstLessonID; 
     }

 function GetFileUrl(fileID) {
     var filePath = Application('dataPath');
     var oConn;
     var oRec;
     var result = "";
      
     oConn = MakeConnection( oConn, filePath );
     Sql = "SELECT  Top 1  url FROM Ficheros  WHERE (id = " + fileID + ")"; 
                                                                              
     oRec = Query( Sql, oRec, oConn  );		                          
     if (!oRec.EOF)
       {
        result = oRec.Fields.Item("url").Value;
       } 
      DestroyAdoObjects( oConn, oRec );
      return result;	       	  
   }
   
   function GetFirstLessonID(course) {
     var filePath = Application('dataPath');
     var oConn;
     var oRec;
     var result = -2;
      
     oConn = MakeConnection( oConn, filePath );
     Sql = "SELECT  Top 1  ID FROM Lecciones WHERE (Course = " + course + ") and (state = " + MOD_ACA_INCOURSE + ") order by [order] Asc"; 

                                                                              
     oRec = Query( Sql, oRec, oConn  );		                          
     if (!oRec.EOF)
       {
        result = oRec.Fields.Item("ID").Value;
       } 
      DestroyAdoObjects( oConn, oRec );
      return result;	       	  
   }


   function GetLastLessonID(course) {
     var filePath = Application('dataPath');
     var oConn;
     var oRec;
     var result = -1;
      
     oConn = MakeConnection( oConn, filePath );
     Sql = "SELECT Top 1 Historia.Lesson "  +
           "  FROM Historia, Lecciones " + 
           "  WHERE ((Lecciones.ID = Historia.Lesson) and (state = " + MOD_ACA_INCOURSE + ") and (Historia.[user] = " + Session("userID") + ") and (Historia.Course = " + course + ")) " 
           "  ORDER BY  Historia.ID Desc";

                                                                              
     oRec = Query( Sql, oRec, oConn  );		                          
     if (!oRec.EOF)
       {
        result = oRec.Fields.Item("Lesson").Value;
       } 
      DestroyAdoObjects( oConn, oRec );
      return result;	       	  
   }

   function GetPrologue(fileID) {
     var filePath = Application('dataPath');
     var oConn;
     var oRec;
     var result = "";
      
     oConn = MakeConnection( oConn, filePath );
     Sql = "SELECT url FROM Ficheros WHERE (ID = " + fileID + ")" ;

                                                                              
     oRec = Query( Sql, oRec, oConn  );		                          
     if (!oRec.EOF)
       {
        result = oRec.Fields.Item("url").Value;
       } 
      DestroyAdoObjects( oConn, oRec );
      return result;	       	  
   }

   function GetFullCourse( course ) {
     var filePath = Application('dataPath');
     var oConn;
     var oRec;
     var result;
      
     oConn = MakeConnection( oConn, filePath );
     Sql = "SELECT Cursos.* "  + 
           "FROM Cursos " + 
           "WHERE (Cursos.ID=" + course + ")";

                                                                              
     oRec = Query( Sql, oRec, oConn  );		                          
     if (!oRec.EOF)
       {
        ID               = oRec.Fields.Item("ID").Value;
        Name		= oRec.Fields.Item("Name").Value;
        state            = oRec.Fields.Item("State").Value;
        owner            = oRec.Fields.Item("owner").Value;
        prologue         = GetPrologue(oRec.Fields.Item("prologue").Value);
        Atmost               = oRec.Fields.Item("Atmost").Value;
        ctdExAct		= oRec.Fields.Item("ctdExAct").Value;
        ctdExRet            = oRec.Fields.Item("ctdExRet").Value;
        MaxFile            = oRec.Fields.Item("MaxFile").Value;
        MaxMem            = oRec.Fields.Item("MaxMem").Value;
        firstLessonID     = GetFirstLessonID(course);
        lastLessonID     = GetLastLessonID(course);
   
        result =  new FullCourse(ID,Name,owner, state, prologue, Atmost, ctdExAct, ctdExRet, MaxFile, MaxMem, lastLessonID, firstLessonID);
       } 
      DestroyAdoObjects( oConn, oRec );
      return result;	       	  
   }
   
   function Lesson(ID, dir, index, name) 
     {this.ID = ID;
      this.dir = dir;
      this.index = index;
      this.name = name;
     }   
            
 
   function GetModuloCourses( modulo ) {

    var filePath = Application('dataPath');
    var oConn;
    var oRec;
    var result = new Array();
     
    oConn = MakeConnection( oConn, filePath );
//    Sql = "select * FROM Cursos Where (modulo = " + modulo + ") Order By ModOrder, ID asc";
    Sql = "select * FROM Cursos Where (modulo = " + modulo + ") Order By ID asc";
  
    oRec = Query( Sql, oRec, oConn  );		
    count = 0;
    while (!oRec.EOF)
      {
       ID               = oRec.Fields.Item("ID").Value;
       Name		= oRec.Fields.Item("Name").Value;
       state            = oRec.Fields.Item("State").Value;
       owner            = oRec.Fields.Item("owner").Value;
  
       result[count] =  new Course(ID,Name,owner, state);
       count++
       oRec.Move(1);  
       
      } 
     DestroyAdoObjects( oConn, oRec );
     return result;	
   }
   	

  function lesson_order( course ) 
    {
     var k;
     for (k=0;k < order.length;k++)
      {if (order[k] == course) {return k; break;}
      }
    }               
    
   
    
   function insCtocourse(user, course)
    {
    
      var filePath = Application('dataPath');
      var oConn;
      var oRec;
 
    oConn = MakeConnection( oConn, filePath );
      Sql = "select * from Conexiones_a_Cursos";
      oRec = Query( Sql, oRec, oConn  );

       //				Response.Write( user + " - - " + course + " - - " + getAtualTime());

      oRec.AddNew();
      oRec.Fields.Item("User").value      = user;
      oRec.Fields.Item("Course").value    = course;
      oRec.Fields.Item("Date").value    = getAtualTime();
      oRec.Update();
      DestroyAdoObjects( oConn, oRec );
//      Response.Write(course);
    }
  
  function deleteHistory(user, course) {
      var filePath = Application('dataPath');
      var oConn;
      var oRec;
 
      oConn = MakeConnection( oConn, filePath );
      Sql = "select * from Historia where ([User] =" + user + ") and (Course=" + course + ")";
      oRec = Query( Sql, oRec, oConn  );		    
      if (!oRec.EOF) oRec.Delete();
      
      DestroyAdoObjects( oConn, oRec );
   }    
   
  function GetLessonIndex(lessonID) {
    var i;  	  
    var l = Session("lessons");
    for(i = 0; i <= Session("lessonCant"); i++)  {
      //Response.Write(l.items.id[i]  + "==" +  lessonID + "<br>");  
      if (l.items.id[i] == lessonID) return i;  	
    } 	
    return -1;
  }
%>
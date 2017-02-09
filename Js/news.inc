<%
 
   //Notifica si hay ejercicios por revizar...
    function hasMail()
     {
       var filePath = Application('dataPath');
       var oConn;
       var oRec;
       var result = "";
 
       oConn = MakeConnection( oConn, filePath );
       Sql = "select top 1 * from Mensajeria where ([To] ="+ Session("userID") +") and (readed =" + Application("dfalse") + ")";
       oRec = Query( Sql, oRec, oConn  );		
 
       if (!oRec.EOF) result = "Tiene correos en su buzón.";           
       
       DestroyAdoObjects( oConn, oRec );
     
      return result;
   }     

   //Notifica si hay ejercicios por revizar...
    function hasTask()
     {
       var filePath = Application('dataPath');
       var oConn;
       var oRec;
       var result = "";
 
       oConn = MakeConnection( oConn, filePath );
       Sql = "select top 1 * from Evaluaciones_Pendientes A,Cursos B where (B.Owner ="+ Session("userID") +") and (A.Course = B.ID)";
       oRec = Query( Sql, oRec, oConn  );		
 
       if (!oRec.EOF) result = "Tienes ejercicios pendientes por revizar en alguno de tus módulos.";           
       
       DestroyAdoObjects( oConn, oRec );
     
      return result;
   }     
    
    
    //Calacula mejor usuario en cuanto a la docencia...
    function getBestUser()
     {
       var filePath = Application('dataPath');
       var oConn;
       var oRec;
       var result = "";
        
 
       oConn = MakeConnection( oConn, filePath );
       
       var best = 0;
       var bestUser;
       var minimo = 20; //minima cantidad de evaluaciones para poder ser el mejor usuario...
       
       Sql = "SELECT Top 1 B.Name FROM Evaluaciones A,Usuarios B WHERE (A.[User] = B.ID)  Group by B.Name Having (Count(*) >= " + minimo + ")  Order by AVG(A.Puntuation) DESC";
       oRec = Query( Sql, oRec, oConn  );		
           
       if (oRec.EOF  == true)
           {result = "Embúllense y comprueben sus conocimientos,los reto";}
          else  
           {result =  "El usuario más destacado en la docencia : " + oRec.Fields.Item("Name").Value;}

       DestroyAdoObjects( oConn, oRec );
           
       return  result;
     }
   
 
 //Devuelve el curso mas visitado... 
 function getBestCourse()
  {
       var filePath = Application('dataPath');
       var oConn;
       var oRec;
       var result = "";
        
 
       oConn = MakeConnection( oConn, filePath );
       
       
       Sql = "SELECT Top 1 A.Name,COUNT(*) as Cantidad FROM Cursos A, Conexiones_a_Cursos B WHERE (A.ID = B.Course) GROUP BY A.Name Order by COUNT(*) DESC";
       oRec = Query( Sql, oRec, oConn  );		
   
       if (!oRec.EOF)
         result = "El curso mas visitado es  " + oRec.Fields.Item("Name").Value; 
       
       DestroyAdoObjects( oConn, oRec );
           
       return  result;       
  }

 //Crea un objeto arreglo...
 function array(arr)
  {
   this.item = arr;
  }
 
 //Devuelve el numero de visitas...
 function getCantVisistText()
  {
  
   return "SEPAD ha sido visitado " + Application("usersCount") + " veces";
  } 
 
 //Crea un arreglo de noticias...
 function addNews()
  {

   var newsCant = 10; // Maxima cantidad de noticias...

   var arrNews = new Array();
   getCantVisist();  
   arrNews[0] = getCantVisistText();
   arrNews[1] = hasMail();
   if (Session("permissionType") == PUBLICATOR)
    {arrNews[2] = hasTask();}  else arrNews[2] = "";
 //  arrNews[3] = getBestCourse(); 
 //  arrNews[4] = getBestUser();
 //  arrNews[5] = "Visite la sección de estadisticas y descubrirá datos increibles acerca de usted y SEPAD";
  

       var filePath = Application('dataPath');
       var oConn;
       var oRec;
       var result = "";
        
 
       oConn = MakeConnection( oConn, filePath );
       
       var best = 0;
       var bestUser;
       var minimo = 20; //minima cantidad de evaluaciones para poder ser el mejor usuario...
       
       Sql = "select top "+ newsCant +" Titulo from Noticias order by Fecha_Publicada DESC"; 
       oRec = Query( Sql, oRec, oConn  );		

   
   
       while (oRec.EOF == false)
        {    
          arrNews[arrNews.length] = oRec.Fields.Item("Titulo") + "";
          oRec.Move(1);
        }
   DestroyAdoObjects( oConn, oRec );
   
   //Session("news") = new array(arrNews);
    
   var cad = "";
   for (i = 0; i < arrNews.length; i++)
    {
     cad = cad + arrNews[i] + "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" ;
    }
   Session("fullNews") = cad;  
   
 }

%>
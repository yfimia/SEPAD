 //Obtiene la cantidad de visitas al sistema
 function getCantVisist()
  {
    var filePath = Application('dataPath');
    var oConn;
    var oRec;
 
    oConn = MakeConnection( oConn, filePath );
    Sql = "select SUM(Logins) as numvisit from Usuarios"; 

    oRec = Query( Sql, oRec, oConn  );		
    
    
    Application.Lock;
     Application("usersCount") = oRec.Fields.Item(0).value;
    Application.UnLock;  

    Session("userVisit") = oRec.Fields.Item(0).value;

    DestroyAdoObjects( oConn, oRec );
   
   return Application("usersCount");
  }


//Chequea si no estan tratanso de ejecutar un script...
  function notIsScript(cad)

      {  //alert(cad);
       if (cad != '')
         {
          if ((cad.indexOf('>') != -1) || (cad.indexOf('&#62;') != -1) || (cad.indexOf('&gt;') != -1))
             {
              return (false); 
             }
          else
           return true;  
         } 
       else
         return true; 

      }

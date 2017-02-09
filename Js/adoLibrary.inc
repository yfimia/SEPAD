<%

                     
//Inincializa el Recordset para hacer una consulta...    
function  InitRecordset( oRec )
	{
	    	oRec.CursorLocation = 3;
        	oRec.CursorType     = 3;
	}    

//Establece la coneccion con el dsn especificado...    
function MakeConnection( oConn, filePath )
	{ 
	  //filePath="DSN=sepad1.1a";
	  var oConn = Server.CreateObject("ADODB.Connection");
      try {
 	  oConn.Open (filePath);       	 	
      }
        catch(e) {		
          code = e.number & 0xFFFF;
          switch (code) {
//            case 1010 : Response.Redirect("errorpage.asp?tipo=Error&short=" + INVALID_FILE_EXTENTION_SHORT  + "&desc=" + INVALID_FILE_EXTENTION_TEXT);
//            break;
              default :  Response.Redirect("errorpage.asp?tipo=Error&short=" + code + "&desc=" + e.description);
          }
      }	

          return oConn;
	}	

function DestroyAdoObjects( oConn, oRec )
	{
	 	oRec.Close(); 	
	 	oConn.Close();
	 	
	}		

//Devuelve true si esta vacia el recorset corespondiente a la consulta Sql...
function empty(Sql) {
    var oConn;
    var oRec;
    var result;
 
    oConn = MakeConnection( oConn, Application("filePath"));
 

    oRec = Query( Sql, oRec, oConn  );		
    result = oRec.EOF;

    DestroyAdoObjects( oConn, oRec );
    return result;
 }



//Devuelve un Recordset con la consulta especificada en Sql y con la BD en filePath...
function Query( Sql, oRec, oConn  )			
	{
         	
	  oRec  = Server.CreateObject("ADODB.Recordset");
      InitRecordset( oRec );
//      try {
 	    oRec.Open( Sql,oConn,3,3);		
  /*    }
      catch(e) {		
        code = e.number & 0xFFFF;
        switch (code) {
//            case 1010 : Response.Redirect("errorpage.asp?tipo=Error&short=" + INVALID_FILE_EXTENTION_SHORT  + "&desc=" + INVALID_FILE_EXTENTION_TEXT);
//            break;
              default :  Response.Redirect("errorpage.asp?tipo=Error&short=" + code + "&desc=" + e.description);
        }
      }*/
          	
            	
   	return oRec;
  }		
    
function ComandQuery( Sql, oRec, oConn  )			
	{
         	
 	  oRec  = Server.CreateObject("ADODB.Recordset");
      InitRecordset( oRec );
      try {
       	oRec.Execute( Sql,oConn,3,3);		
      }
      catch(e) {		
        code = e.number & 0xFFFF;
        switch (code) {
//            case 1010 : Response.Redirect("errorpage.asp?tipo=Error&short=" + INVALID_FILE_EXTENTION_SHORT  + "&desc=" + INVALID_FILE_EXTENTION_TEXT);
//            break;
              default :  Response.Redirect("errorpage.asp?tipo=Error&short=" + code + "&desc=" + e.description);
        }
      }


      return oRec;
  }		
    

%>  
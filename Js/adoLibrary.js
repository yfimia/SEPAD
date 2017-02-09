

                     
//Inincializa el Recordset para hacer una consulta...    
function  InitRecordset( oRec )
	{
	    	oRec.CursorLocation = 3;
        	oRec.CursorType     = 3;
	}    

//Establece la coneccion con el dsn especificado...    
function MakeConnection( oConn, filePath )
	{ 
	  Response.Write(filePath);
	  var oConn = Server.CreateObject("ADODB.Connection");
	  oConn.Open (filePath);       	 	
          return oConn;
	}	

function DestroyAdoObjects( oConn, oRec )
	{
	 	oRec.Close(); 	
	 	oConn.Close();
	 	
	}		

//Devuelve un Recordset con la consulta especificada en Sql y con la BD en filePath...
function Query( Sql, oRec, oConn  )			
	{
         	
    		oRec  = Server.CreateObject("ADODB.Recordset");
          	InitRecordset( oRec );
              	oRec.Open( Sql,oConn,3,3);		
          	return oRec;
	}		
    
function ComandQuery( Sql, oRec, oConn  )			
	{
         	
    		oRec  = Server.CreateObject("ADODB.Recordset");
          	InitRecordset( oRec );
              	oRec.Execute( Sql,oConn,3,3);		
          	return oRec;
	}		
    


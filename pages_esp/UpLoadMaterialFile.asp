<%@ Language=jScript %>
<%  
 Response.Expires = -1;
%>
<!-- #include file='../js/user.inc' -->
<!-- #include file='../js/adolibrary.inc' -->
<!-- #include file="../js/library.inc" -->

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

%>


<%
   var mySmartUpload, maxFile;
   
   var mySmartUpload = Server.CreateObject("aspSmartUpload.SmartUpload");
   
   mySmartUpload.DeniedFilesList = "asp,php,php3,pl,dll";
   maxFile 			 = 10240 * 1024;     
   mySmartUpload.MaxFileSize 	 = maxFile;
     
      
    
   if (
       (Request.QueryString.Item("mid").Count > 0) &&
       (Request.QueryString.Item("url").Count > 0) &&
       (Request.QueryString.Item("lesson").Count > 0)       
       )
   { 

    	var filePath = Application('dataPath');    
    	var oConn    = Server.CreateObject("ADODB.Connection");
    	var oRec     = Server.CreateObject("ADODB.Recordset");  
    	var sql, description, title, mid, lname;

	mid          = Request.QueryString.Item("mid") + "";    	    	
	url          = Request.QueryString.Item("url") + "";    	    	
	lesson       = Request.QueryString.Item("lesson") + "";    	    	
    	sql	     = 'SELECT * from Ficheros WHERE (Id = ' + mid + ')';                	
    	
	try 
   	{
   		mySmartUpload.Upload();
   	}
   	catch(e) 
   	{		
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

    	description  = mySmartUpload.Form.Item('FileDescription').Values;
    	title        = mySmartUpload.Form.Item('FileTitle').Values; 
    	
	oConn.Open( filePath );	   
    	oRec.Open( sql, oConn, 3, 3);
    	    	
    	if (! oRec.EOF ) 
    	{ 
  	   
  	   if (mySmartUpload.Files.Item("FileName").Size > 0)
  	   {
  	
	    	oRec.Fields.Item("url").value          = mySmartUpload.Files.Item("FileName").FileName;
  	    	oRec.Fields.Item("tama").value         = Math.round( (mySmartUpload.Files.Item("FileName").Size / 1024) * 1000) / 1000;   	      	    	

  	    	mySmartUpload.Files.Item("FileName").SaveAs("../Courses/Course" + Session("course") + url + "/" + mySmartUpload.Files.Item("FileName").FileName); 	    	 	    
  	    }

  	    oRec.Fields.Item("description").value  = description;
  	    oRec.Fields.Item("title").value  	   = title;
      	    
       	    oRec.Update();   		
    	}
    
    
     
    	DestroyAdoObjects( oConn, oRec );
    	Response.Redirect('docs.asp?uid=' + uid + '&lesson=' + lesson);        
  }
%>

<%
  }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>


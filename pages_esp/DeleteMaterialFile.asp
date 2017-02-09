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
	course	     = 	Session("course");
    	sql	     = 'SELECT * from Ficheros WHERE (Id = ' + mid + ')';

	oConn.Open( filePath );	   
    	oRec.Open( sql, oConn, 3, 3);
    	file = oRec.Fields.Item("url").value;
    	    	
    	if (! oRec.EOF ) { 
    	
    		oRec.Close();
	    	
	    	sql	     = 'DELETE from Ficheros WHERE (Id = ' + mid + ')';
	
	    	oRec.Open( sql, oConn, 3, 3);
	
	    	fso 	     = Server.CreateObject("Scripting.FileSystemObject");
    		fso.DeleteFile(Server.MapPath("../courses/course" + course + url + "/" + file), true);
	    	//Response.Write(Server.MapPath("../courses/course" + course + url + "/" + file));
    	
    		Response.Redirect('docs.asp?uid=' + uid + '&lesson=' + lesson);        

    	}
	oConn.Close();
  }
%>

<%
  }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>


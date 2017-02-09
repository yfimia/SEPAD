<%@ Language=JavaScript%>
<%
  Response.Expires = -1;
%>
<!-- #include file="../js/adoLibrary.inc" -->
<!-- #include file='../js/user.inc' -->

<%  

  var filePath = Application('dataPath');
  
  var oConn;
  var  oComm    = Server.CreateObject("ADODB.Command");

%>

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       


      oConn = MakeConnection( oConn, filePath );
      oComm.ActiveConnection = oConn;
      
      oComm.CommandText = "DELETE FROM Anuncios WHERE (ID = " + Request.QueryString.Item("IDAnuncio") + ")";
      oComm.Execute();
      oConn.Close
           

       var idmod = Request.querystring("idmodulo");
       Response.Redirect("Anuncios.asp?uid=" + uid +  "&idmodulo=" + idmod);
       
  }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
  //Response.Write(Request.QueryString.Item("uid"));
%>

<%@LANGUAGE="JAVASCRIPT"%>
<%
  Response.Expires = -1;
%>
<!-- #include file='../../js/adoLibrary.inc' -->
<!-- #include file='../../js/user.inc' -->

<%
if (Request.QueryString.Item("uid") + "" == Session("uid"))
{    
  var uid = Request.QueryString.Item("uid") + ""; 
  if ((Session("isTutor") > -1) || 
      (Session("PermissionType") == ADMINISTRATOR) || 
	  (Session("userID") == Session("tutcursocordinador"))  || 
	  (Session("userID") == Request.QueryString.Item("courseOwner")))
  {
    var dataPath = Application('dataPath');
    oConn = Server.CreateObject("ADODB.Connection");
	oConn.Open( dataPath );
	
	for (i=0; i<=Request.Form("counter"); i++)
	{
	  if (Request.Form("CheckBox" + i) == "on")
	  {
	    var Sql = "DELETE FROM Evaluaciones_Pendientes WHERE ID = " + Request.Form("Hidden" + i);
		oConn.Execute(Sql);
	  }
	}
	oConn.Close();
  }
  
  Response.Redirect("exerlist.asp?uid=" + uid);	  
}
	
%>
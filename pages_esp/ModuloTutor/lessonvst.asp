<%@ Language=JavaScript %>
<!-- #include file="../../js/Adolibrary.inc" -->
<!-- #include file="../../js/library.inc" -->
<!-- #include file='../../js/user.inc' -->
<!-- #include file='inc/lessonVst.inc' -->

<%
 Response.Expires = -1;
 
 if (Request.QueryString.Item("uid") + "" == Session("uid"))
 {    
      var uid      = Request.QueryString.Item("uid") + "";       
	  var courseID = Session("tutcurso") + "";

	  if ((Session("isTutor") > -1) || (Session("PermissionType") == ADMINISTRATOR) || (Session("userID") == Session("tutcursocordinador"))  || (Session("userID") == Request.QueryString.Item("courseOwner")))
   	  {     
      
     	var filter = -1; 
     
     	if ( Request.Form("Filter").Count > 0 )
     		filter = Request.Form("Filter") + "";
     	
     	var oConn;
     	var oRec;
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title><%= TXT_TITLE %></title>
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/EvalResult.css" type="text/css">

</head>
<body bgcolor="#FFFFFF" text="#000000" style="overflow-y:scroll">
  <table width="100%">
  	<tr>
  	  <form name="form1" action="lessonvst.asp?uid=<%=uid%>&courseOwner=<%=Request.QueryString.Item("courseOwner")%>"" method="post">	
  		<td align="right">		  
		  	<table>
		  		<tr>
		  			<td>
		  				Lección : 
		  			</td>
		  			<td>  			
	  					<select  class=SearchCombo name="Filter" id="Filter" onchange="fOnChange()"> 
		              		<option  value=-1>Todas</option>		              		
<%
	 	sql   = "SELECT * FROM Lecciones WHERE Course = " +  courseID;
	 	oConn = MakeConnection( oConn, Application("dataPath") );
	 	oRec  = Query( sql, oRec, oConn );
	 
	 	while ( oRec.EOF == false ) 
	 	{
                    text = oRec.Fields.Item('Name').value + "";
                  if (text.length > 30) {
                    text = text.substring(0, 30) + "...";
                  }
%>	
				          <option  value="<%= oRec.Fields.Item('Id').value %>"><%=text%></option>		              		
<%
			oRec.Move( 1 );
		}
	
		oRec.Close();
%>	              		
		            	</select>		            
		  			</td>
			  	</tr>
		  </table>
		</td>
	  </form>	 	
	</tr>
	<tr>	  
		<td>
 <%	
	  Sql = "SELECT fullUserName AS Usuario, lessonName AS Lección, Cant AS Visitas FROM HistoriaLecciones " + 
	        "WHERE Course = " + courseID;
	        
	  if ( -1 != filter )
	     Sql += " AND Lesson = " + filter;
	  	
	  putTable( Application("dataPath"), TXT_TITLE, "../../images/" + Session("skin") + "/ReportesIMG.gif", Sql, -1, TEXT2); 
 %>
 	   </td>
 	</tr>   	
</table>			
</body>

<script language="jscript">
	function fOnChange()
	{
		document.forms[ "form1" ].submit();
	}	
	
	document.all[ "Filter" ].value = <%= filter %>;	
</script>

</html>
<%
   }
  else
    Response.Redirect("../errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);     
 }
 else
   Response.Redirect("../errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>
	

































<%@ Language=JavaScript %>
<%
  Response.Expires = -1;
%>
<!-- #include file='../js/user.inc' -->
<!-- #include file='inc/userSearch.inc' -->


<%
  if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       


  if (Session("PermissionType") == ADMINISTRATOR)
    {
    first = "-1";
    where = "";
    if (Request.QueryString.Count >= 2) {
      first = Request.QueryString.Item("first") + "";
      
      if (first != "-1") {
        where = " and (Name > '" +  first + "')";
      } 
    } 
    
	var display		  = false;
	var uid;
	var oRec;	
	
	if (Request.QueryString.Item("uid") + "" == Session("uid"))
	{
	   uid = Request.QueryString.Item("uid") + "";       		
	}
	else
           Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
	
			
	if (
	    (Request.Form("pattern").Count > 0) && (Request.Form("option").Count > 0)	    
	   )
	{
	
		var pattern   = Request.Form("pattern") + "";
		var option    = parseInt( Request.Form("option") );
		var query     = "";
		var escape    = "\\";
		var filePath  = Application("dataPath");  		
			 display   = true;
			 
		
		switch ( option )
		{
			case 0 : {
							query = "select * from usuarios WHERE Name LIKE '%" + pattern +"%' ESCAPE '" + escape +"' ORDER BY Name, FullName, Email";
							break;
						}
			case 1 : {
							query = "select * from usuarios WHERE FullName LIKE '%" + pattern +"%' ESCAPE '" + escape +"' ORDER BY FullName, Name, Email";
							break;
						}
			case 2 : {
							query = "select * from usuarios WHERE Email LIKE '%" + pattern +"%' ESCAPE '" + escape +"'  ORDER BY Email, Name, FullName";
							break;
						}			
		}

	
		
		oRec   =  Server.CreateObject( "ADODB.RecordSet" );
		
		oRec.Open ( query, filePath );
		  		
	}
%>

<HTML>
	<HEAD>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">			
		<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main1.css" type="text/css">			
		<link rel="stylesheet" href="../css/<%=Session("skin")%>/EvalResult.css" type="text/css">
		<script src="../js/CheckBoxes.js" language="JavaScript"></script>
	
         	<title><%= TITLE1 %></title>
<style type="text/css">
<!--
.EditInput {
	width: 100%;
}
-->
</style>
	</HEAD>
	<body bgcolor="#ffffff" text="#000000">
		<table align="center" width="100%" border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td class="ToolBar">
					<table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
						<tr>
							<td><img src="../images/<%=Session("skin")%>/LessonImg.gif" width="80" height="54"></td>
							<td class="HeaderTable"><%=PAGE_HEADER_TEXT%></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
		<br>
		<form name="fsearch" id="fsearch" method=post action="userSearch.asp?uid=<%=uid%>">
			<table border="0" cellspacing="0" cellpadding="0" width="100%">
				<tr>
					<td nowrap="nowrap" width="1%" class="ToolBar"><%=SEARCH_LABEL_TEXT%><b>
					</td>
					<td width="98" class="ToolBar" align="left">
						<input id="pattern" name="pattern" class="EditInput">
					</td>
					<td class="ToolBar" width="1%">
						<input type="submit" value="<%=SEARCH_TEXT%>" ID="Submit1" NAME="Submit1">
					</td>
				</tr>
				<tr>
				  <td colspan="3">
				    <table border="0" cellspacing="0" cellpadding="0" style="border-width: 1px; border-style: solid; border-color: black" width="100%">
				      <tr>
					<td class="MessageTR1"><input type="radio" name="option" id="option" value="0" checked=1><%=SEARCH_ID_TEXT%></td>
					<td class="MessageTR1"><input type="radio" name="option" id="option" value="1"><%=FULL_NAME_TEXT%></td>
					<td class="MessageTR1"><input type="radio" name="option" id="option" value="2"><%=EMAIL_TEXT%></td>
				      </tr>		
				    </table>  
				  </td>  
				</tr>
			</table>
		</form>
	
	   <% if (display) { %>
<form name=Delnews action="ConfirmDeleteUser.asp?uid=<%=uid%>" method="post">
<table border="0" cellspacing="0" cellpadding="0" class="ToolBar" width="100%">
  <tr>
    <td>
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td nowrap><a href="AgregaUser.asp?uid=<%=uid%>" class="ToolLink">&nbsp;<%= TEXT1 %>&nbsp;</a></td>
          <td>|</td>
          <td nowrap><a href="javascript:Delnews.submit()" class="ToolLink">&nbsp;<%= TEXT2 %>&nbsp;</a></td>
          <td>|</td>
          <td><a href="#" class="ToolLink" onClick="CheckAll(GetObjectCollection( document, 'FORM'))">&nbsp;<%= TEXT3 %>&nbsp;</a></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar" align="center"><b><%= USER_LIST_TEXT %></b></td>
  </tr>
  <tr> 
    <td> 
      <table width="100%" border="0" cellspacing="1" cellpadding="0">
        <tr> 
          <td width="4%" class="ToolBar" align="center"></td>
          <td width="20%" class="ToolBar" align="center"><b><%= TEXT5 %></b></td>
          <td width="30%" class="ToolBar" align="center"><b><%= TEXT6 %></b></td>
          <td width="30%" class="ToolBar" align="center"><b><%= TEXT7 %></b></td>
          <td width="8%" class="ToolBar" align="center"><b><%= TEXT8 %></b></td>
          <td width="8%" class="ToolBar" align="center"><b><%= TEXT9 %></b></td>
        </tr>   
             		<% 
             		      var clase = "MessageTR1"; 
             		      var i = 0;
             		%>
        		<%
        		      while ( !oRec.EOF ) { 
        		          i = i + 1;
        		%>

        <tr> 
          <td width="4%" class="<%=clase%>">
<% 
     if(oRec.Fields.Item("ID").value != GUEST_USER && oRec.Fields.Item("ID").value != ADMIN_USER){ 
%>
	     <input type="checkbox" id=Check<%=i%> name=check<%=i%> >
             <input type="hidden" id=ID<%=i%> name="ID<%=i%>" value="<%=oRec.Fields.Item("ID").value%>">
             <input type="hidden" id=Name<%=i%> name="Name<%=i%>"  value="<%=oRec.Fields.Item("Name").value%>" >
<%
     }
%>             
          </td>
          <td width="20%" class="<%=clase%>"><a href="ModifyUser.asp?uid=<%=uid%>&user=<%=oRec.Fields.Item("ID").value%>"><%=oRec.Fields.Item("name").value%></a></td>
          <td width="30%" class="<%=clase%>"><%=oRec.Fields.Item("FullName").value%></td>
          <td width="30%" class="<%=clase%>"><a href="mailto:<%=oRec.Fields.Item("email").value%>" ><%=oRec.Fields.Item("email").value%></a></td>
          <td width="8%" class="<%=clase%>"><%=oRec.Fields.Item("fechaIngreso").value%></td>
          <td width="8%" class="<%=clase%>"><%=oRec.Fields.Item("lastLogin").value%></td>         
        </tr> 
        		<% 
        			
        			if (clase == "MessageTR1") clase = "MessageTR"; else clase = "MessageTR1";
        			
        			oRec.Move( 1 );
        			} 
        			
        			oRec.Close();
        		%>
        		
        		</table>
         </td>        		
         </tr>      
		</table>
<input type="hidden"  name="Count" value="<%=i%>" >
<input type="hidden"  name="First" value="-1" >
		
<table border="0" cellspacing="0" cellpadding="0" class="ToolBar" width="100%">
  <tr>
    <td>
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td nowrap><a href="AgregaUser.asp?uid=<%=uid%>" class="ToolLink">&nbsp;<%= TEXT1 %>&nbsp;</a></td>
          <td>|</td>
          <td nowrap><a href="javascript:Delnews.submit()" class="ToolLink">&nbsp;<%= TEXT2 %>&nbsp;</a></td>
          <td>|</td>
          <td><a href="#" class="ToolLink" onClick="CheckAll(GetObjectCollection( document, 'FORM'))">&nbsp;<%= TEXT3 %>&nbsp;</a></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
		</form>
<%	}	%>
<%
    }
  else  
    {
      Response.Redirect("errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);         
      }
%>

</BODY>
</HTML>
<%
  }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
  //Response.Write(Request.QueryString.Item("uid"));
%>
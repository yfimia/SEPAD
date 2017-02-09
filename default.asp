<%@ Language=JavaScript %>
<%
 Response.Expires = 0;





 /*
     Pagina para el chequeo de los datos del usuario
     
     Parametros:
       -username       : login a comprobar
       -userpassword   : password a comprobar 
 */


 //var pagesdir = "beta/";
 var pagesdir = "pages_esp/";

%>

<!-- #include file="js/Adolibrary.inc" -->
<!-- #include file="js/user.inc" -->
<!-- #include file="js/modulo.inc" -->
<!-- #include file="js/md5.inc" -->
<%


 var userID,userName="",userPassword="";
 
uid = Session("uid");
%>
 
 


<%
  
 if (Request.Form.Count >= 2) { 
       userName                = Request.Form('userName') + "";   
       userPassword            = Request.Form('userPassword') + "";
 }
 else {
   userName                = GUEST_LOGIN;		
   userPassword            = calcMD5(GUEST_PASSWORD);
       
 }      
//       Response.Write(userName +   userPassword);     
%>
   
<% 
    
     var access = checkuser(userName, userPassword);    
  //   Response.Write('    ' + access);     
     switch (access) {
%>

<%
		case UNDEFINED_USER: Response.Redirect(pagesdir + "errorpage.asp?tipo=Error&short=" + UNDEFINED_USER_SHORT  + "&desc=" + UNDEFINED_USER_TEXT);
        break;
        
		case INVALID_PASSWORD: Response.Redirect(pagesdir + "errorpage.asp?tipo=Error&short=" + INVALID_PASSWORD_SHORT  + "&desc=" + INVALID_PASSWORD_TEXT);
        break;
		
		case GUEST_USER: 
%>

		Guest

             
<%
        break;
		
		case VALID_USER: 
		//Es un usuario autentificado
		Session("ChatTimeToRefresh") = 10;
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<Title>Sistema de Enseñanza Personalizada A Distancia</Title>

<script language="jscript">
 function goGenMod(ID, Name, state, cordinador, grupo, claustro) {
   parent.frames("mainFrame").location = "modulo.asp?uid=<%=uid%>&ID=" + ID + "&Name=" + Name  + "&state=" + state + "&cordinador=" + cordinador + "&grupo=" + grupo + "&claustro=" + claustro; 	
 }
 
</script>

</HEAD>

<frameset rows="112,*" frameborder="NO" border="0" framespacing="0" cols="*"> 
  <frame name="topFrame" scrolling="NO" noresize src="<%=pagesdir%>choicecourse.asp?uid=<%=uid%>" >
  <frameset cols="187,*" frameborder="NO" border="0" framespacing="0" rows="*"> 
    <frame name="leftFrame" scrolling="YES" noresize src="<%=pagesdir%>selection.asp?uid=<%=uid%>&flag=0">
    <%
    /***********************************************************
    Si es una modalidad lo que se pasa por QueryString
    ************************************************************/
    if (!isNaN(Request.QueryString.Item("modalidad")))
    {
    	modalidad = Request.QueryString.Item("modalidad");
		  var filePath = Application('dataPath');
		  var oConn;
		  var oRec;
		
		  oConn = MakeConnection( oConn, filePath );
		  Sql = "select count(*) as Cantidad FROM modulo WHERE id = " + modalidad;
		
		  oRec = Query( Sql, oRec, oConn  );
		  if (oRec.Fields.Item("Cantidad").Value == 0) {
		  	var defaultpage = Application("HomePage");
		  }
			else
			{
		  	Sql = "select * FROM modulo WHERE id = " + modalidad;
		
		  	oRec = Query( Sql, oRec, oConn  );
				var defaultpage = "modulo.asp?uid=" + uid + "&ID=" + modalidad + "&Name=" + oRec.Fields.Item("Name").Value  + "&state=" + oRec.Fields.Item("state").Value + "&cordinador=" + oRec.Fields.Item("cordinador").Value + "&grupo=" + oRec.Fields.Item("grupo").Value + "&claustro=" + oRec.Fields.Item("claustro").Value;
			}
    %>
    <frame name="mainFrame" scrolling="YES" noresize src="<%=pagesdir + defaultpage%>" >
    <%
    }
  	else
    /***********************************************************
    Si es un curso lo que se pasa por QueryString
    ************************************************************/
    if (!isNaN(Request.QueryString.Item("curso"))) {
    	curso = Request.QueryString.Item("curso");
		  var filePath = Application('dataPath');
		  var oConn;
		  var oRec;
		
		  oConn = MakeConnection( oConn, filePath );
		  Sql = "select count(*) as Cantidad FROM cursos WHERE id = " + curso;
		
		  oRec = Query( Sql, oRec, oConn  );
		  if (oRec.Fields.Item("Cantidad").Value == 0) {
		  	var defaultpage = Application("HomePage");
		  }
			else
			{
		  	Sql = "select * FROM cursos WHERE id = " + curso;
		
		  	oRec = Query( Sql, oRec, oConn  );
				var defaultpage = "course.asp?uid=" + uid + "&ID=" + curso + "&Name=" + oRec.Fields.Item("Name").Value;
			}
    %>
    <frame name="mainFrame" scrolling="YES" noresize src="<%=pagesdir + defaultpage%>" >
    <%
    }
  	else
  	{
    %>
    <frame name="mainFrame" scrolling="YES" noresize src="<%=pagesdir + Application("HomePage")%>" >
    <%
    }
    %>
  </frameset>
</frameset>

</HTML>
         
<%
        break;

     }     

%>     




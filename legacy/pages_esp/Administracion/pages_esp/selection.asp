<%@ Language=JavaScript %>
<%
   Response.Expires = -1;
%>
<!-- #include file='../js/adolibrary.inc' -->
<!-- #include file='../js/course.inc' -->
<!-- #include file='../js/modulo.inc' -->
<!-- #include file="../js/user.inc" -->

<%   
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       
%>

<HTML>
<head>
<title>Listado de las modalidades académicas</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<LINK href="../css/<%=Session("skin")%>/Main1.css" rel=stylesheet type=text/css>
<link rel="stylesheet" href="../css/<%=Session("skin")%>/listspan.css" type="text/css">


<script language="jscript">
 function goGenMod(ID, Name, state, cordinador, grupo, claustro) {
   parent.frames("mainFrame").location = "modulo.asp?uid=<%=uid%>&ID=" + ID + "&Name=" + Name  + "&state=" + state + "&cordinador=" + cordinador + "&grupo=" + grupo + "&claustro=" + claustro; 	
 }
</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" class="TreeViewBody" style="overflow-y:scroll">

<% //Esto lo usa las busquedas para el patron por defecto... no quitar
   var pattern ="";
   var tipo=1;
%>
<!-- #include file="../js/search.inc" -->
<%
   download_Modules()

%>

</BODY>

</HTML>
<%
 }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>

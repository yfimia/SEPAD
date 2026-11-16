<%@ Language=JavaScript %>

<!-- #include file='../js/adolibrary.inc' -->
<!-- #include file='../js/user.inc' -->
<!-- #include file='../js/course.inc' -->
<!-- #include file='../js/change.inc' -->
<!-- #include file='inc/CalifOfic.inc' -->

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       
      var temp = GetUserData(Session("userID"));
      var alumnFullName = temp.fullName;
%>

<html>
<head>
<title><%=title%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/Reports.css" rel="stylesheet" type="text/css" />

<script src="../js/windows.js" language="JavaScript"></script>
</head>

<body bgcolor="#FFFFFF" text="#000000" style="overflow-y:scroll">
<center>
<H3><%=stmp1%></H3>
<H4><%=stmp7%><I><U><%=Session("fullName")%></U><I></H4>
<H4><%=stmp8%><I><U><%=Session("courseName")%></U><I></H4>
</center>
<br>
<br>
<br>
<%

          var dataPath = Application('dataPath');
          var oConn;
          var oRec;
          
          oConn = MakeConnection( oConn, dataPath );

          Sql = "SELECT Calif_Ofic.titulo, Calif_Ofic.calif, Usuarios.ID, Usuarios.fullName, Calif_Ofic.obs " + 
                "FROM Calif_Ofic, Usuarios " +
                "WHERE (Calif_Ofic.alumn = " + Session("userID") + ") and (Calif_Ofic.curso = " + Session("course") + ") and (Usuarios.ID = Calif_Ofic.prof)";
          
          oRec = Query( Sql, oRec, oConn  );		
%>        
  <center>
  
         <table width="90%" border="0" cellspacing="1" cellpadding="0">
           <tr> 
             <td colspan=4 width="100%" class="Main" align="center"></td>
           </tr>
         
           <tr> 
             <td width="25%%" class="Main" align="center"><b><%=stmp2%></b></td>
             <td width="5%" class="Main" align="center"><b><%=stmp3%></b></td>
             <td width="25%" class="Main" align="center"><b><%stmp4%></b></td>
             <td width="35%" class="Main" align="center"><b><%=stmp5%></b></td>
           </tr>

<%
          var i = 0;
          while (oRec.EOF == false)
            { 
              i = i + 1;
%>
          <tr> 
            <td width="25%" class="First">&nbsp;<%=oRec.Fields.Item("titulo").value%>&nbsp;</td>
            <td width="5%" class="First">&nbsp;<%=oRec.Fields.Item("calif").value%>&nbsp;</td>
            <td width="25%" class="First">&nbsp;<%=oRec.Fields.Item("fullName").value%>&nbsp;</td>
            <td width="35%" class="Last">&nbsp;<%=oRec.Fields.Item("obs").value%>&nbsp;</td>
          </tr> 
  
<%
           oRec.Move(1);
         }  
         if (i == 0)
           {
%>        
             <tr><td align="center" width="100%" colspan="4" class="Last"><%=msg1%></td></tr>        
<%
           }
%>        
         </table>
  </center>
</html>
<%
  DestroyAdoObjects( oConn, oRec );        
 }
 else
   {Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT)};
%>


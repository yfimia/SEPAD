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
%>

<html>
<head>
<title><%=title%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/EvalResult.css" type="text/css">

<script src="../js/windows.js" language="JavaScript"></script>
</head>

<body bgcolor="#FFFFFF" text="#000000" style="overflow-y:scroll">

<table width="100%" border="0" cellspacing="0" cellpadding="2">
  <tr> 
    <td width="99%" class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td><img src="../images/<%=Session("skin")%>/ReportesIMG.gif" width="80" height="54"></td>
          <td class="HeaderTable"><%=stmp1%></td>
        </tr>
      </table>
    </td>
    <td class="ToolBar" width="1%" valign="bottom">
      <a target="_blank" href="printableCalifOfic.asp?uid=<%=uid%>"><img height=20 width=20 src="../images/<%=Session("skin")%>/imprimir.gif" title="<%=stmp6%>" border=0></a>
    </td>
  </tr>
</table>
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
  
         <table width="90%" border="0" cellspacing="1" cellpadding="0" class="PrologueTable">
           <tr> 
             <td colspan=4 width="100%" class="ToolBar" align="center"></td>
           </tr>
         
           <tr> 
             <td width="25%%" class="ToolBar" align="center"><b><%=stmp2%></b></td>
             <td width="5%" class="ToolBar" align="center"><b><%=stmp3%></b></td>
             <td width="25%" class="ToolBar" align="center"><b><%stmp4%></b></td>
             <td width="35%" class="ToolBar" align="center"><b><%=stmp5%></b></td>
           </tr>

<%
          var clase = "MessageTR1";
          
          var i = 0;
          while (oRec.EOF == false)
            { 
              i = i + 1;
              if (clase == "MessageTR1") clase = "MessageTR"; else clase = "MessageTR1";
%>
          <tr> 
            <td width="25%" class="<%=clase%>">&nbsp;<%=oRec.Fields.Item("titulo").value%>&nbsp;</td>
            <td width="5%" class="<%=clase%>">&nbsp;<%=oRec.Fields.Item("calif").value%>&nbsp;</td>
            <td width="25%" class="<%=clase%>">&nbsp;<a href="javascript:abreVentana('',350,280,'userDetails.asp?uid=<%=uid%>&id=<%=oRec.Fields.Item("ID").value%>', 'yes', 'yes')"><%=oRec.Fields.Item("fullName").value%></a>&nbsp;</td>
            <td width="35%" class="<%=clase%>">&nbsp;<%=oRec.Fields.Item("obs").value%>&nbsp;</td>
          </tr> 
  
<%
           oRec.Move(1);
         }  
         if (i == 0)
           {
%>        
             <tr><td align="center" width="100%" colspan="4" class="MessageTR"><%=msg1%></td></tr>        
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

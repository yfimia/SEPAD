<%@ Language=JScript %>
<%Response.Expires = -1;%>
<!-- #include file='../../js/adoLibrary.inc' -->
<!-- #include file='../../js/user.inc' -->
<!-- #include file='../../js/change.inc' -->
<!-- #include file='inc/verCalifOfic.inc' -->

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       
      if (Request.QueryString.Item("curso").count > 0)
        var curso = Request.QueryString.Item("curso") + "";
      if (Request.QueryString.Item("alumn").count > 0)
        var alumn = Request.QueryString.Item("alumn") + "";
      var temp = GetUserData(alumn);
      alumnFullName = temp.fullName;
%>

<%
 

 if ((Session("isTutor") > -1) || (Session("PermissionType") == ADMINISTRATOR) || (Session("userID") == Session("tutcursocordinador"))  || (Session("userID") == Request.QueryString.Item("courseOwner")))
    {
        
        
%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/EvalResult.css" type="text/css">
<script src="../../js/windows.js" language="JavaScript"></script>
<script src="../../js/CheckBoxes.js" language="JavaScript"></script>
</head>
<body bgcolor="#FFFFFF" text="#000000">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td class="HeaderTable"  align=center><b><%= TEXT1 %></b></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<br>

<center>
  <b><%= TEXT2 %>:   </b><i><%=Session("courseName")%></i><br>
  <b><%= TEXT3 %>:   </b><i><a href="javascript:abreVentana('',350,280,'userDetails.asp?uid=<%=uid%>&courseOwner=<%=Request.QueryString.Item("courseOwner")%>&id=<%=alumn%>', 'yes', 'yes')"><%=alumnFullName%></a></i>
</center>
<br>
    <form name=DelCalif action="ConfirmDeleteCalif.asp?uid=<%=uid%>&courseOwner=<%=Request.QueryString.Item("courseOwner")%>" method="post">
      <table width="100%" border="0" cellspacing="1" cellpadding="0">
        <tr> 
          <td width="1%" class="ToolBar" align="center">&nbsp; </td>
          <td width="25%" class="ToolBar" align="center"><b><%= TEXT4 %></b></td>
          <td width="5%" class="ToolBar" align="center"><b><%= TEXT5 %></td>
          <td width="34%" class="ToolBar" align="center"><b><%= TEXT6 %></b></td>
          <td width="25%" class="ToolBar" align="center"><b><%= TEXT7 %></b></td>
        </tr>


<%
          var dataPath = Application('dataPath');
          var oConn;
          var oRec;
          
          oConn = MakeConnection( oConn, dataPath );

          Sql = "SELECT * " + 
                "FROM Calif_Ofic " +
                "WHERE (curso = " + curso + ") and (alumn = " + alumn + ")";
         
          oRec = Query( Sql, oRec, oConn  );		

          var clase = "MessageTR1";
          var i = 0;
          while (oRec.EOF == false)
            { 
              i = i + 1;
              if (clase == "MessageTR1") clase = "MessageTR"; else clase = "MessageTR1";
%>
          <tr> 
 
            <td width="1%" class="<%=clase%>">
              <input type="checkbox" id=Check<%=i%> name=check<%=i%> > 
            </td>
            <td width="25%" class="<%=clase%>"><a href='addCalifOfic.asp?uid=<%=uid%>&courseOwner=<%=Request.QueryString.Item("courseOwner")%>&id=<%=oRec.Fields.Item("ID").value%>&alumn=<%=alumn%>&curso=<%=curso%>'><%=oRec.Fields.Item("titulo").value%></a></td>
            <td width="5%" class="<%=clase%>">&nbsp;<%=oRec.Fields.Item("calif").value%>&nbsp;</td>
            <td width="35%" class="<%=clase%>">&nbsp;<%=oRec.Fields.Item("obs").value%>&nbsp;</td>
            <td width="30%" class="<%=clase%>">&nbsp;<a href="javascript:abreVentana('',350,280,'userDetails.asp?uid=<%=uid%>&courseOwner=<%=Request.QueryString.Item("courseOwner")%>&id=<%=oRec.Fields.Item("prof").value%>', 'yes', 'yes')"><%=GetUserData(oRec.Fields.Item("prof").value).fullName%></a>&nbsp;</td>
          </tr> 

            
<%
           oRec.Move(1);
         }  
         if (i == 0)
           {
%>        
             <tr><td align="center" width="100%" colspan="5" class="MessageTR"><%= TEXT8 %>.</td></tr>        
              <tr>
                 <td class="ToolBar" colspan=5>
                   <input type="hidden" name="curso" value=<%=curso%>>
                   <input type="hidden" name="alumn" value=<%=alumn%>>
                   <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
                     <tr> 
                       <td nowrap><a class="ToolLink" href='addCalifOfic.asp?uid=<%=uid%>&courseOwner=<%=Request.QueryString.Item("courseOwner")%>&alumn=<%=alumn%>&curso=<%=curso%>'>&nbsp;<%= TEXT9 %>&nbsp;</a></td>
                       <td>|</td>
                     </tr>
                   </table>
                 </td>
               </tr>
<%
           }
           else
             {
%>        
              <tr>
                 <td class="ToolBar" colspan=5>
                   <input type="hidden" name="curso" value=<%=curso%>>
                   <input type="hidden" name="alumn" value=<%=alumn%>>
                   <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
                     <tr> 
                       <td nowrap><a class="ToolLink" href='addCalifOfic.asp?uid=<%=uid%>&courseOwner=<%=Request.QueryString.Item("courseOwner")%>&alumn=<%=alumn%>&curso=<%=curso%>'>&nbsp;<%= TEXT9 %>&nbsp;</a></td>
                       <td>|</td>
                       <td nowrap><a href="javascript:DelCalif.submit()" class="ToolLink">&nbsp;<%= TEXT10 %>&nbsp;</a></td>
                       <td>|</td>
                       <td><a href="#" class="ToolLink" onClick="CheckAll(GetObjectCollection( document, 'FORM'))">&nbsp;<%= TEXT11 %>&nbsp;</a></td>
                     </tr>
                   </table>
                 </td>
               </tr>
<%
              }
%>               
           </table>
         </form>

<%
    }    
  else
    {  
     Response.Redirect("../errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);                   
    }
%>

</BODY>
</HTML>
<%
  }
 else
   Response.Redirect("../errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
   
  //Response.Write(Request.QueryString.Item("uid"));
%>



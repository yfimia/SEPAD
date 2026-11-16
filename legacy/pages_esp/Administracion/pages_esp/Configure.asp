<%@ Language=JavaScript %>
<%
  Response.Expires = -1;
%>

<!-- #include file='../js/user.inc' -->
<!-- #include file='inc/configure.inc' -->

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       
  if (Session("PermissionType") == ADMINISTRATOR)
    {


%>

<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/EvalResult.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td class="HeaderTable"  align=center><b><%=stmp1%></b></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
    <FORM action="ConfirmConfigure.asp?uid=<%=uid%>" method="POST">
       <table width="100%" border="0" cellspacing="1" cellpadding="1" align="center">
        <TR>
          <TD colspan=2 class="ToolBar" align=center></b><%=stmp2%></b></TD>
        </TR>
        <TR>
          <TD width="50%" class="ToolBar" align=right><%=stmp3%></TD>
          <TD width="50%"  class="ToolBar" align=left><INPUT type=edit  size=30 name=DSN value="<%=Application("dataPath")%>">
           <br><%=stmp4%></TD>
        </TR>
        <TR>
          <TD align=right class="ToolBar" width="50%" ><%=stmp5%></TD>
          <TD align=left class="ToolBar" width="50%" >
            <SELECT  name=SQL>
         <%
            var selected1 = "";
            var selected2 = "selected";
            if (Application("SQL") == "1") {
              selected1 = "selected";
              selected2 = "";
            }
         %>
            <OPTION value=1  <%=selected1%>><%=stmp6%></OPTION>
            <OPTION value=0  <%=selected2%> ><%=stmp7%></OPTION>
          </SELECT></TD>
        </TR>
      </TABLE>
       <table border="0" cellspacing="1" cellpadding="1" align="center">
        <TR>
          <TD colspan=2 class="ToolBar" align=center></b><%=stmp8%></b></TD>
        </TR>

        <TR>
          <TD width="40%" class="ToolBar" align=right><%=stmp9%></TD>
          <TD width="50%"  class="ToolBar" align=left><INPUT type=edit  size=30 name=MailAddress value="<%=Application("MailAddress")%>"></TD>
        </TR>
        <TR>
          <TD align=right  width="40%" class="ToolBar"><%=stmp10%></TD>
          <TD align=left  width="50%" class="ToolBar"><INPUT type=edit  size=30 name=MailUser value="<%=Application("MailUser")%>"></TD>
        </TR>
        <TR>
          <TD align=right  width="40%" class="ToolBar"><%=stmp11%></TD>
          <TD align=left class="ToolBar"  width="50%" ><INPUT  size=30 type=password name=MailPassword value="<%=Application("MailPassword")%>"></TD>
        </TR>
        <TR>
          <TD align=right class="ToolBar" width="40%" ><%=stmp12%></TD>
          <TD align=left class="ToolBar" width="50%" ><INPUT type=edit size=30  name=MailServer value="<%=Application("MailServer")%>"></TD>
        </TR>
        <TR>
          <TD align=right class="ToolBar" width="40%" ><%=stmp13%></TD>
          <TD align=left class="ToolBar" width="50%" ><INPUT type=edit size=5  name=Port value="<%=Application("Port")%>"></TD>
        </TR>
        <TR>
          <TD align=right class="ToolBar" width="40%" ><%=stmp14%></TD>
          <TD align=left class="ToolBar" width="50%" ><SELECT  name=Auth>
         <%
            var selected1 = "";
            var selected2 = "selected";
            if (Application("Auth") == "1") {
              selected1 = "selected";
              selected2 = "";
            }
         %>
            <OPTION value=1  <%=selected1%>><%=stmp15%></OPTION>
            <OPTION value=0  <%=selected2%>><%=stmp16%></OPTION>
          </SELECT></TD>
        </TR>
        <TR>
          <TD align=right class="ToolBar" width="40%" ><%=stmp17%></TD>
          <TD align=left class="ToolBar" width="50%" ><INPUT type=edit size=70  name=Subject value="<%=Application("Subject")%>"></TD>
        </TR>
        <TR>
          <TD align=right class="ToolBar" width="40%" ><%=stmp18%></TD>
          <TD align=left class="ToolBar" width="50%" ><TEXTAREA name=Body rows=8 cols=70><%=Application("Body")%></TEXTAREA></TD>
        </TR>
        
      </TABLE>
       <table    border="0" cellspacing="1" cellpadding="1" align="center">
        <TR>
          <TD colspan=2 class="ToolBar" align=center></b><%=stmp19%></b></TD>
        </TR>

        <TR>
          <TD  width="40%" class="ToolBar" align=right><%=stmp20%></TD>
          <TD class="ToolBar"  width="50%" align=left><INPUT type=edit size=30  name=logoimg value="<%=Application("LogoImg")%>"></TD>
        </TR>
        <TR>
          <TD align=right class="ToolBar"  width="40%" ><%=stmp21%></TD>
          <TD align=left class="ToolBar" width="50%" ><INPUT type=edit size=30  name=homepage value="<%=Application("HomePage")%>"></TD>
        </TR>
        <TR>
          <TD align=right class="ToolBar"  width="40%" ><%=stmp22%></TD>
          <TD align=left class="ToolBar" width="50%" ><INPUT type=edit size=30  name=downloads value="<%=Application("Downloads")%>"></TD>
        </TR>
        
        <TR>
          <TD align=right class="ToolBar"  width="40%" ><%=stmp23%></TD>
          <TD align=left class="ToolBar" width="50%" ><INPUT type=edit size=30  name=listsize value=<%=Application("ListSize")%>></TD>
        </TR>
        <TR>
          <TD align=right class="ToolBar"  width="40%" ><%=stmp24%></TD>
          <TD align=left class="ToolBar"  width="50" ><INPUT type=edit size=30 name=newnews value=<%=Application("NewNews")%>></TD>
        </TR>
      </TABLE>
      
      <CENTER>
        <INPUT type=submit value="<%=stmp25%>">
      </CENTER>
    </FORM>
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

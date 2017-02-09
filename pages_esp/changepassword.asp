<%@ Language=JavaScript%>
<%
 Response.Expires = -1;



%>
<!-- #include file='../js/user.inc' -->
<!-- #include file='inc/changepassword.inc' -->
<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

%>

<SCRIPT LANGUAGE=javascript src="../js/md5.js"></script>
<SCRIPT src="../Js/user.js" language="JSCRIPT"></SCRIPT>

<html>
<head>
<title><%=title%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<script language="javascript">
<!--
function CloseWindow() {
	close();
}
//-->
</script>
</head>

<body bgcolor="#ffffff" text="#000000">



<script ID="clientEventHandlersJS" LANGUAGE="javascript">


function Ok_onclick()
 { 
  // alert(calcMD5(userOldPassword.value));
  if (frm.userConfirm.value == frm.userPassword.value)
    {
    
      frm.userOldPassword.value = calcMD5(frm.userOldPassword.value);
      frm.userPassword.value = calcMD5(frm.userPassword.value);
      frm.userConfirm.value = frm.userPassword.value;

      frm.submit();
    }
  else
    {
     // error el nuevo password y su confirmacion no coinciden....
     alert(DIFFERENT_PASSWORD_TEXT);     
    }    
 }

function Cancel_onclick()
 {
  frm.userConfirm.value     = "";
  frm.userPassword.value    = "";    
  frm.userOldPassword.value = "";    
  close();
 } 


</script>




<form name=frm id=frm action="checkpassword.asp?uid=<%=uid%>" method=post LANGUAGE=javascript >
<table border="0" cellspacing="0" cellpadding="5" bgcolor="#cccccc">
  <tr>
    <td>
      <table border="0" cellspacing="3" cellpadding="0">
        <tr> 
          <td align="left"><%=stmp1%></td>
          <td> 
            <input value="<%= Session('name')%>" id="userName" name="userName" readOnly >
          </td>
        </tr>
        <tr> 
          <td align="left"><%=stmp2%></td>
          <td> 
            <input type="password" id="userOldPassword" name="userOldPassword" maxLength=16 >
          </td>

        </tr>
        <tr> 
          <td align="left"><%=stmp3%></td>
          <td> 
            <input id="userPassword" name="userPassword" type="password" maxLength=16 >
          </td>

        </tr>
        <tr> 
          <td align="left"><%=stmp4%></td>
          <td> 
            <input id="userConfirm" name="userConfirm" type="password" maxLength=16 >
          </td>

        </tr>
        <tr> 
          <td colspan="2">
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr align="middle" valign="center"> 
                <td> 
                  <input type="button" name="Ok" id="Ok" value=<%=stmp5%> onclick="return Ok_onclick()">
                </td>
                <td> 
                  <input type="button" name="Cancel" id="Cancel" value=<%=stmp6%> onclick="return Cancel_onclick()">
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</form>

</body>

</html>

<%
  }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
  //Response.Write(Request.QueryString.Item("uid"));
%>

<%@ Language=JavaScript %>

<%
  Response.Expires = -1;
  
%>

<%=MATRICULAS_PAGE%>

<!-- #include file="../js/adolibrary.inc" -->
<!-- #include file="../js/library.inc" -->
<!-- #include file="../js/news.inc" -->
<!-- #include file="../js/user.inc" -->
<!-- #include file="../js/constants.inc" -->
<!-- #include file="inc/choicecourse.inc" -->
<SCRIPT language=vbscript RUNAT=Server>
 function getAtualTime
  getAtualTime = Now()
 End function 

</SCRIPT>


<% 
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       
      
    //Crea las noticias para pornerlas desde Session("fullnews") en el markee...
    addNews();
    
    
    var courName  = new Array();
    var courCount = new Array();
    
  
%>

<% 
  
  var fecha = new Date ();
  if (Request.QueryString.Count > 1)
    {
     Session("flag")          = "1";
     Session("lastLesson")    = Request.QueryString.Item("lastLesson") + "";
     Session("lastLessonUrl") = Request.QueryString.Item("lastLessonUrl") + "";
     lastlesson               = Session("lastLesson");
    } 
  else
    {
     Session("flag") = "0";
     lastlesson      = ""; 
    }   
 userVisit = Session("userVisit") + ""; 
%>



<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<script language=javascript src="../js/md5.js"></script>
<script type="text/javascript" language="javascript" src="../js/Windows.js"></script>

<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--
                     
var permission = <%= Session("PermissionType")%>

function home() { 
 parent.frames("leftFrame").location.href = "selection.asp?uid=<%=uid%>&flag=1";
 parent.frames("mainframe").location.href = "<%=Application("HomePage")%>";
}

function userPassword_onkeypress()
 {  
  if (event.keyCode == 13)
     { 
      aceptar_onclick(); 
     }
 }
 
function aceptar_onclick()
 {
  frm.userName.value = frm.userName.value.toLowerCase();          
  if (frm.userName.value == "") 
    {
      frm.userName.focus()
    }
  else
    { 
     frm.userPassword.value =  calcMD5(frm.userPassword.value);
     frm.submit(); 
     //parent.location.href = "../checkuser.asp?username=" + userName.value  + "&userpassword=" + calcMD5(userPassword.value + "");
    }                
 }


var lastlesson = "<%= lastlesson %>";
         
//-->
</script>


</head>
<body bgcolor="#ffffff" text="#000000" class="MainBody">
<form name=frm id=frm action="../default.asp" method=post LANGUAGE=javascript  target=_parent>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td align="left" width="40%"><img src="../Images/<%=Session("skin")%>/ChoiceCourse/<%=Application("LogoImg")%>"></td>
    <td align="right" valign="top" > 
      <table width="100%" border="0" cellspacing="0" cellpadding="5">
        <tr> 
          <td align="middle" valign="center" width="75%" class="LBTD"> 
            <table border="0" cellspacing="3" cellpadding="0" class="DateTable">
              <tr> 
                <td> 
                  <div align="left"><b><%=stmp1%></b><%=Session("fullName")%></div>
                </td>
              </tr>
<%  if ((Session("permissionType") == ADMINISTRATOR) || (Session("UserID") != GUEST_USER))
    {
%>
              <tr> 
                <td> 
                  <div align="left"><b><%=stmp2%></b><%= Session("textPermission")%></div>
                </td>
              </tr>
<%  }
%>
              
              <tr> 
                <td> 
                  <div align="left"><b><%=stmp3%></b><%=DAYS[fecha.getDay()] + ' ' + fecha.getDate() + ', de ' + MONTHS[ fecha.getMonth()] + ', ' + fecha.getYear()%></div>
                </td>
              </tr>
              <!--tr> 
                <td><b><%=stmp4%></b><%=userVisit%></td>
              </tr-->
            </table>
          </td>
          <td align="middle" valign="center" width="45%" class="LBTD"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0" class="LoginTable">
              <tr> 
                <td width="35%"><%=stmp5%></td>
                <td width="65%"> 
                  <input id="userName" maxlength="16" name="userName" size="16" onkeypress="return  userPassword_onkeypress()">
                </td>
              </tr>
              <tr> 
                <td width="35%"><%=stmp6%></td>
                <td width="65%"> 
                  <input type="password" name="userPassword" id=userPassword maxlength=16 onkeypress="return userPassword_onkeypress()"  size="16">
                </td>
              </tr>
              <tr align="middle"> 
                <td colspan="2"> 
                  <input type="button" onclick="aceptar_onclick()" name="Submit" value="<%=stmp7%>">
                </td>
              </tr>
              <tr align="middle"> 
                <td colspan="2"><A href='javascript:<%=MATRICULAS_PAGE%>'><%=stmp8%></a></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td align="left" colspan="3"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td class="MarqueeTD" width="99%"><marquee class="MarqueeStd"><%=Session("fullNews")%></marquee></td>
          <td class="MarqueeTD" width="1%">
	    <table width="100" height="14" border="0" cellspacing="0" cellpadding="0" >
	      <tr>
	        <td align="center"><a href="javascript:abreVentana('Downloads', 600, 400, '<%=Application("Downloads")%>', 'yes')"><img src="../images/<%=Session("skin")%>/Download.gif" width="100" height="14" border="0"></a></td>
	      </tr>
	    </table>
          </td>        
        </tr>
      </table>
    </td>
  </tr>
</table>  
<table border="0" cellspacing="0" cellpadding="0" width="100%" class="ServicesTable">
  <tr>
    <td>
      <table border="0" cellspacing="0" cellpadding="2" width="171">
        <tr> 
        
          <td><a href="javascript:home()" class="ToolLink" ><%=stmp9%></a></td>        
          <td class="Separator">|</td>                    
  <%
  if (Session("permissionType") == ADMINISTRATOR)
    {
  %>
          <td><A class=ToolLink href  ="javascript:abreVentana( <%=stmp10%>, 700, 400, 'Administracion/main.asp?uid=<%=uid%>')" ><%=stmp11%></a></td>
          <td class="Separator">|</td>          
  <%
    }
  if (Session("UserID") != GUEST_USER) 
    { //Response.Write(Session("UserID") + "!=" + GUEST_USER);
  %>
          <td><A class=ToolLink href="javascript:abreVentana(<%=stmp12%>, 500,400, 'DatosPersonales.asp?uid=<%=uid%>', 'yes', 'yes')" ><%=stmp13%></a></td>
          <td class="Separator">|</td>

          
          <td><A class=ToolLink href="javascript:abreVentana(<%=stmp14%>, 600, 400, 'inbox.asp?uid=<%=uid%>', 'yes')" ><%=stmp15%></a></td>
          <td class="Separator">|</td>

          <td><A class=ToolLink href="javascript:abreVentana(<%=stmp16%>, 600, 400, 'newChatRoomlist.asp?uid=<%=uid%>', 'no', 'no')" ><%=stmp17%></a></td>
          <td class="Separator">|</td>

          <!--td><A class=ToolLink href="javascript:abreVentana(<%=stmp18%> , 600, 400, 'Festadisticas.asp?uid=<%=uid%>', 'yes')"><%=stmp19%></a></td>
          <td class="Separator">|</td-->
 <%
    }
%>          
         
          <!--td><A class=ToolLink href="javascript:abreVentana( <%=stmp20%>, 600, 400, 'Sestadisticas.asp?uid=<%=uid%>', 'yes')" ><%=stmp21%></a></td>
          <td class="Separator">|</td-->
          
          <td><A class=ToolLink href="javascript:abreVentana(<%=stmp22%>, 600, 400, 'news.asp?uid=<%=uid%>', 'yes')" ><%=stmp23%></a></td>
          <%
            var filePath = Application('dataPath');
  
            var oConn;
            var oRec;
            oConn = MakeConnection( oConn, filePath );
  	        Sql = "SELECT Enlaces.Caption, Enlaces.url FROM Enlaces";

            oRec = Query( Sql, oRec, oConn  );		
            last = -1;
            
            while (!oRec.EOF) {
          %>
           <td class="Separator">|</td>
           <td><A class=ToolLink href="javascript:abreVentana('<%=oRec.Fields.Item("Caption")%>', 600, 400, '<%=oRec.Fields.Item("url")%>', 'yes', 'no')"><%=oRec.Fields.Item("Caption")%></A></td>
          <%
            oRec.move(1);
            }
          %> 
          <td class="Separator">|</td>
          <td><A class=ToolLink href="javascript:abreVentana(<%=stmp24%>, 600, 400, '../Help/Main.htm', 'yes')"><%=stmp25%></a></td>
          <td class="Separator">|</td>
          <td><A class=ToolLink href="javascript:abreVentana(<%=stmp26%>, 600, 400, 'Creditos.htm', 'yes', 'no')"><%=stmp27%></a></td>
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
%>


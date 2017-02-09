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
                  <div align="left"><b>Usuario: </b><%=Session("fullName")%></div>
                </td>
              </tr>
<%  if ((Session("permissionType") == ADMINISTRATOR) || (Session("UserID") != GUEST_USER))
    {
%>
              <tr> 
                <td> 
                  <div align="left"><b>Permisos: </b><%= Session("textPermission")%></div>
                </td>
              </tr>
<%  }
%>
              
              <tr> 
                <td> 
                  <div align="left"><b>Fecha: </b><%=DAYS[fecha.getDay()] + ' ' + fecha.getDate() + ', de ' + MONTHS[ fecha.getMonth()] + ', ' + fecha.getYear()%></div>
                </td>
              </tr>
              <!--tr> 
                <td><b>Visita No.: </b><%=userVisit%></td>
              </tr-->
            </table>
          </td>
          <td align="middle" valign="center" width="45%" class="LBTD"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0" class="LoginTable">
              <tr> 
                <td width="35%">Usuario:</td>
                <td width="65%"> 
                  <input id="userName" maxlength="16" name="userName" size="16" onkeypress="return  userPassword_onkeypress()">
                </td>
              </tr>
              <tr> 
                <td width="35%">Contraseña:</td>
                <td width="65%"> 
                  <input type="password" name="userPassword" id=userPassword maxlength=16 onkeypress="return userPassword_onkeypress()"  size="16">
                </td>
              </tr>
              <tr align="middle"> 
                <td colspan="2"> 
                  <input type="button" onclick="aceptar_onclick()" name="Submit" value="Aceptar">
                </td>
              </tr>
              <tr align="middle"> 
                <td colspan="2"><A href='javascript:<%=MATRICULAS_PAGE%>'>Registrarse en SEPAD</a></td>
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
        
          <td><a href="javascript:home()" class="ToolLink" >Inicio</a></td>        
          <td class="Separator">|</td>                    
  <%
  if (Session("permissionType") == ADMINISTRATOR)
    {
  %>
          <td><A class=ToolLink href  ="javascript:abreVentana( 'Adminsitracion', 700, 400, 'Administracion/main.asp?uid=<%=uid%>')" >Administración</a></td>
          <td class="Separator">|</td>          
  <%
    }
  if (Session("UserID") != GUEST_USER) 
    { //Response.Write(Session("UserID") + "!=" + GUEST_USER);
  %>
          <td><A class=ToolLink href="javascript:abreVentana('chpw', 500,400, 'DatosPersonales.asp?uid=<%=uid%>', 'yes', 'yes')" >Datos&nbsp;Personales</a></td>
          <td class="Separator">|</td>

          
          <td><A class=ToolLink href="javascript:abreVentana('Mensajeria', 600, 400, 'inbox.asp?uid=<%=uid%>', 'yes')" >Mensajería</a></td>
          <td class="Separator">|</td>

          <td><A class=ToolLink href="javascript:abreVentana('Chat', 600, 400, 'newChatRoomlist.asp?uid=<%=uid%>', 'no', 'no')" >Chat</a></td>
          <td class="Separator">|</td>

          <!--td><A class=ToolLink href="javascript:abreVentana( 'Estadisticas_del_Sistema', 600, 400, 'Festadisticas.asp?uid=<%=uid%>', 'yes')">Calificaciones</a></td>
          <td class="Separator">|</td-->
 <%
    }
%>          
         
          <!--td><A class=ToolLink href="javascript:abreVentana( 'Estadisticas_del_Sistema', 600, 400, 'Sestadisticas.asp?uid=<%=uid%>', 'yes')" >Datos&nbsp;del&nbsp;Sistema</a></td>
          <td class="Separator">|</td-->
          
          <td><A class=ToolLink href="javascript:abreVentana('Noticias', 600, 400, 'news.asp?uid=<%=uid%>', 'yes')" >Noticias</a></td>
          <td class="Separator">|</td>
          <td><A class=ToolLink href="javascript:abreVentana('Ayuda', 600, 400, '../Help/Main.htm', 'yes')">Ayuda</a></td>
          <td class="Separator">|</td>
          <td><A class=ToolLink href="javascript:abreVentana('Ayuda', 600, 400, 'Creditos.htm', 'yes', 'no')">Acerca&nbsp;de...</a></td>
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


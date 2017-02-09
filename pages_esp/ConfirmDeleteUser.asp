<%@ Language=JScript %>
<%
  Response.Expires = -1;
  Response.Buffer = true;  
%>

<!-- #include file='../js/user.inc' -->
<!-- #include file='../js/adolibrary.inc' -->
<!-- #include file='inc/confirmdeleteuser.inc' -->

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

%>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
</HEAD>


<script language="jscript">
  function gogogo()
    {
      ret.src = "../images/RegresarDn.gif";
      ret.src = "../images/RegresarUp.gif";
      window.navigate("p0.htm");
    }
</script>
<body>

<p>
<%
  if (Session("PermissionType") == ADMINISTRATOR)
    {
    
   var  oConn = Server.CreateObject("ADODB.Connection"); 
   var  oComm    = Server.CreateObject("ADODB.Command");
   var  filePath = Application("dataPath");   
   
   var Tot = parseInt(Request.Form("Count"));
   var first = parseInt(Request.Form("First"));
    

       oConn.Open(filePath);
       oComm.ActiveConnection = oConn; 

%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar" align="center"><h5><%=stmp1%></h5></td>
  </tr>
  <tr> 
    <td> 
      <table width="100%" border="0" cellspacing="1" cellpadding="0">
        <tr> 
          <td width="35%" class="ToolBar" align="center"><b><%=stmp2%></b></td>
          <td width="65%" class="ToolBar" align="center"><b><%=stmp3%></b></td>
        </tr>

<%   var clase = "MessageTR1";

     for (var i = 1; i <= Tot; i++) {
       if ((Request.Form("check" + i).Count > 0) && (Request.Form("ID" + i).Count > 0) && (Request.Form("ID" + i) != ADMIN_USER) && (Request.Form("ID" + i) != GUEST_USER) ) {

          if (clase == "MessageTR1") clase = "MessageTR"; else clase = "MessageTR1";
       	
%>
        <tr> 
<%
              if ( empty("SELECT ID FROM Modulo WHERE (cordinador = " + Request.Form("ID" + i) + ")") && empty("SELECT ID FROM Cursos WHERE (owner = " + Request.Form("ID" + i) + ")")
                    && empty("SELECT ID FROM Subgrupo WHERE (tutor = " + Request.Form("ID" + i) + ")")) {

                oComm.CommandText = "delete  from Usuarios where (ID = '" + Request.Form("ID" + i) + "')";
                oRec = oComm.Execute();
%>
                <td width="35%" class="<%=clase%>"><%=Request.Form("Name" + i)%></td>
                <td width="65%" class="<%=clase%>"><b><%=stmp4%></b></td>
                 
<%                
              }
              else {
%>
                <td width="35%" class="<%=clase%>"><%=Request.Form("Name" + i)%></td>
                <td width="65%" class="<%=clase%>"><b><%=stmp5%></b><%=stmp6%></td>
<%                 
              	
              }  
         
%>
        </tr> 

<%          
         
       }
     }

     
         oConn.Close();          
            
%>
</table>
<%      



        
      //Response.Redirect("Deleteuser.asp?uid=" + uid);      
      
      Response.Write ('<script languaje=javascript>function pclick(){window.location="Deleteuser.asp?uid=' + uid + '&first=' + first + '";}</script>');
      Response.Write ("<center><input type=Button Value=" + stmp7 + " onclick='return pclick()'></center>" );      ; 
    }
  else
    {  
       Response.Redirect("errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);         
    }

%>
</p>
    
</body>
</html>
<%
  }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
  //Response.Write(Request.QueryString.Item("uid"));
%>
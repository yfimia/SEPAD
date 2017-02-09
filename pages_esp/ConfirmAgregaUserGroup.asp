<%@ Language=JScript %>
<%
  Response.Expires = -1;
    Response.Buffer = true;
%>
<!-- #include file='../js/adolibrary.inc' -->
<!-- #include file='../js/user.inc' -->

<%

    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       
      var first = Request.QueryString.Item("first") + "";       
      var IDG = Request.Form.Item("GrupoADD");          
if (Session("PermissionType") == ADMINISTRATOR)
    {
      var  filePath = Application("dataPath");   
      var  oRec  = Server.CreateObject("ADODB.Recordset");
      var  oRec1  = Server.CreateObject("ADODB.Recordset");
      var  oConn = Server.CreateObject("ADODB.Connection"); 
      var  oComm    = Server.CreateObject("ADODB.Command");
      
      
      var Tot = parseInt(Request.Form("Count"));
      var first = parseInt(Request.Form("First"));
    
      oConn.Open(filePath);
      oComm.ActiveConnection = oConn; 

      
     for (var i = 1; i <= Tot; i++) {
       if ((Request.Form("check" + i).Count > 0) && (Request.Form("ID" + i).Count > 0) ) {
          AddUserToGroup(Request.Form("ID" + i), IDG);
       }
     }     
     
       Response.Redirect("AgregaUserGroup.asp?uid=" + uid + "&first=" + first + "&GrupoADD=" + IDG);      
       
       Response.Write ('<script languaje=javascript>function pclick(){window.location="AgregaUserGroup.asp";}</script>');      
       Response.Write ("<center><input type=Button Value=Regresar onclick='history.back(-1)' id=Button1 name=Button1></center>" );      ; 

%>

<html>
<head>
<link REL="stylesheet" TYPE="text/css" HREF="../css/SepadCss1.css" />
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</head>
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
    }
  else
    {  
     Response.Redirect("errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);                       
    }
%>
</body>
</html>
<%
  }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
  //Response.Write(Request.QueryString.Item("uid"));
%>

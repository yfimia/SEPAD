oo<%@ Language=JScript %>
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
      
if (Session("PermissionType") == ADMINISTRATOR)
    {
      var  filePath = Application("dataPath");   
      var  oRec  = Server.CreateObject("ADODB.Recordset");
      var  oRec1  = Server.CreateObject("ADODB.Recordset");
      var  oConn = Server.CreateObject("ADODB.Connection"); 
      var  oComm    = Server.CreateObject("ADODB.Command");
      var grupo = new Array();
      oRec  = Session("UserList");
      oConn = Session("Conection");
      oComm = Session("Command");
      var i = 0;
      var j = 0;
      var fname = new Array();
      var IDU = new Array();
      var nick = new Array();
      while (oRec.EOF == false)
        {
          i = i + 1;
          if (Request.Form("Check" + i) == "on") 
            {
              j = j + 1;
              fname[j] = oRec.Fields.Item("FullName").Value + "";
              IDU[j] = oRec.Fields.Item("ID").Value;
              nick[j] = oRec.Fields.Item("Name").Value;
            }
          oRec.Move(1);
        }
       oConn.Close();   
   
       if (j >= 1)
         {
/*           oConn.Open(filePath);
           oRec.Open("select * from Grupos where (ID = " + Request.Form("GrupoADD") + ")",oConn,3,3);
           Response.Write("select * from Grupos where (ID = '" + Request.Form("GrupoADD") + "')");
           var IDG = oRec.Fields.Item("ID").Value;
           var IDG =  Request.Form("GrupoADD")
           oConn.Close();
*/
           var IDG =  Request.Form("GrupoADD") + "";

           for (i=1; i<=j; i++) 
             {
               AddUserToGroup(IDU[i], IDG);
             }
           
         }  
         
       Session("UserList") = null;
       Session("Conection") = null;
       Session("Command") = null;
     
       Response.Redirect("AgregaUserGroup.asp?uid=" + uid + "&first=" + first + "&group=" + IDG);      
       
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

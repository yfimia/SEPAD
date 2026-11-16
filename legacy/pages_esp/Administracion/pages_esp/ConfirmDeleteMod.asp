<%@ Language=JScript %>
<%
  Response.Expires = -1;
  Response.Buffer = true;  
%>

<!-- #include file='../js/user.inc' -->

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

%>

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
      var  filePath = Application("dataPath");   
      var oConn      = Server.CreateObject ("ADODB.Connection");
      var SQL = "delete  from Modulo where ";

     oConn.Open(filePath);

     var Tot       = Request.Form("Count");
     var i;

      for (i=1;i <= Tot; i++) {                                        
        if ((Request.Form("check" + i).Count > 0) && (Request.Form("ID" + i).Count > 0))
          {
           SQL = SQL + "(ID=" + Request.Form("ID" + i) + ") OR ";
          }              
      }    

      SQL = SQL + "(1=0)";
      oConn.Execute (SQL);

//Response.Write(Tot + "  "  + SQL);

      oConn.Close();   
        
      Response.Redirect("DeleteMod.asp?uid=" + uid);      
      

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
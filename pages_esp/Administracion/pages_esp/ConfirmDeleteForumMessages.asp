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
      var  oRec  = Server.CreateObject("ADODB.Recordset");
      var  oConn = Server.CreateObject("ADODB.Connection"); 
      var  oComm    = Server.CreateObject("ADODB.Command");
      var grupo = new Array();
      oRec  = Session("UserList");
      oConn = Session("Conection");
      oComm = Session("Command");
      var i = 0;
      var j = 0;
      var fID = new Array();
      while (oRec.EOF == false)
        {
          i = i + 1;
          if (Request.Form("Check" + i) == "on") 
            {
              j = j + 1;
              fID[j] = oRec.Fields.Item("ID").Value + "";
            }
          oRec.Move(1);
        }
      oConn.Close();   
      if (j >= 1)
        {
          for (i=1; i<=j; i++) 
            {
              oConn.Open(filePath);
              oComm.ActiveConnection = oConn; 
              //Response.Write ("<font color=red>delete * from Grupos where (Name = '" + grupo[i] + "')</font>");         
              oComm.CommandText = "delete  from ForumMessages where (ID = " + fID[i] + ")";
              oRec = oComm.Execute();
              oConn.Close();          
            }
          Response.Write("<center><font color=red>Mensajes eliminados satisfactoriamente</font><br></center>");  
        }  
      Session("UserList") = null;
      Session("Conection") = null;
      Session("Command") = null;
        
      Response.Redirect("DeleteForumMessages.asp?uid=" + uid);      
      
      Response.Write ('<script languaje=javascript>function pclick(){window.location="DeleteForumMessages.asp";}</script>');      
      Response.Write ("<center><input type=Button Value=Regresar onclick='return pclick()'></center>" );      ; 
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
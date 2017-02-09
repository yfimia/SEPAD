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
      var IDMA = new Array();
      while (oRec.EOF == false)
        {
          i = i + 1;
          if (Request.Form("Check" + i) == "on") 
            {
              j = j + 1;
              IDMA[j] = oRec.Fields.Item("ID").Value;
            }
          oRec.Move(1);
        }
      oConn.Close();   
      
      
      if (j >= 1)
        {
          //borro las matriculas 
          for (i=1; i<=j; i++) 
            {
              oConn.Open(filePath);
              oComm.ActiveConnection = oConn; 
              oComm.CommandText = "delete  from matricula where (ID = " + IDMA[i] + ")";
              oRec = oComm.Execute();
              oComm.CommandText = "delete  from Mod_Matricula where (idmatricula = " + IDMA[i] + ")";
              oRec = oComm.Execute();
              
              oConn.Close();          
            }  
      
        }  
      Session("UserList") = null;
      Session("Conection") = null;
      Session("Command") = null;
      
      Response.Redirect("MatriculasManager.asp?uid=" + uid + "&id=" + Request.QueryString.Item("id"));      
      
      Response.Write ('<script languaje=javascript>function pclick(){window.location="deleteMatriculas.asp";}</script>');      
      Response.Write ("<center><input type=Button Value=Regresar onclick='return pclick()' id=Button1 name=Button1></center>" );      ;
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

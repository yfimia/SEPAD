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
      //window.parent.selectionCourse.navigate("selection.asp");
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
      var CID = new Array();
      var news = new Array();
      oRec  = Session("CoursesList");
      oConn = Session("Conection");
      oComm = Session("Command");
      var j = 0;
      var i = 0;
      while (oRec.EOF == false)
        {
          i = i + 1;
          if (Request.Form("Check" + i) == "on") 
            {
              j = j + 1;
              CID[j] = oRec.Fields.Item("ID").Value + "";
              news[j] = oRec.Fields.Item("Titulo").Value + "";    
            }
          oRec.Move(1);
        }
      oConn.Close();   
      //Response.Write("<font color=Black>" + j + "Elementos</font>");
      if (j >= 1)
        {
          for (i=1; i<=j; i++) 
            {
              //Response.Write("<font color=Black>" + i + "</font>");
              oConn.Open(filePath);
              oComm.ActiveConnection = oConn;
              oComm.CommandText = "delete  from Noticias where (ID = " + CID[i] + ")";
              oRec = oComm.Execute();
              oConn.Close();          
            }

        }  
      Session("CoursesList") = null;
      Session("Conection") = null;
      Session("Command") = null;
      
      Response.Redirect("newsManager.asp?uid=" + uid);      
      
      Response.Write ('<script languaje=javascript>function pclick(){window.location="newsManager.asp";}</script>');      
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

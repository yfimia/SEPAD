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
      var last = Request.QueryString.Item("first") + "";
      var group = Request.QueryString.Item("group") + "";
      var groupname = Request.QueryString.Item("groupname") + "";
      

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
      var UID = new Array();
      var GID = new Array();
      var NU = new Array();
      var NG = new Array();
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
              UID[j] = oRec.Fields.Item("UserID").Value + "";
              GID[j] = oRec.Fields.Item("GroupID").Value + "";    
              NU[j] = oRec.Fields.Item("Name").Value + "";
             // NG[j] = oRec.Fields.Item("Grupo").Value + "";
            }
          oRec.Move(1);
        }
      oConn.Close();   
      //Response.Write("<font color=Black>" + j + "Elementos</font>");
      if (j >= 1)
        {
          for (i=1; i<=j; i++) 
            {
              oConn.Open(filePath);
              oComm.ActiveConnection = oConn;
              oComm.CommandText = "delete  from Grupos_de_Usuarios where ([User] = " + UID[i] + ") and ([Group] = " + GID[i] + ")";
              oRec = oComm.Execute();
                            
              oConn.Close();          
            }

        }  
      Session("CoursesList") = null;
      Session("Conection") = null;
      Session("Command") = null;
      
      Response.Redirect("DeleteUserOneGroup.asp?uid=" + uid + "&group=" + group + "&groupname=" + groupname  + "&first=" + last);      
      
      Response.Write ('<script languaje=javascript>function pclick(){window.location="DeleteUserGroup.asp";}</script>');      
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

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


<%
  if ((Session("PermissionType") == ADMINISTRATOR) || (Session("PermissionType") == PUBLICATOR))
    {

  groupID	= Request.Form.Item('groupID') + '';
  //Response.Write(groupID);
%>
<%
      var  filePath = Application("dataPath");   
      var  oRec  = Server.CreateObject("ADODB.Recordset");
      var  oConn = Server.CreateObject("ADODB.Connection"); 
      var  oComm    = Server.CreateObject("ADODB.Command");

      oRec  = Session("UserList");
      oConn = Session("Conection");
      oComm = Session("Command");
      var i = 0;
      var j = 0;
      var IDTrab = new Array();
      var URLs = new Array();
      
      while (oRec.EOF == false)
        {
          i = i + 1;
          if (Request.Form("Check" + i) == "on") 
            { 
              j = j + 1;
              IDTrab[j] = oRec.Fields.Item("ID").Value;
              URLs[j] = oRec.Fields.Item("fileURL").Value;
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
              oComm.CommandText = "delete  from Seminarios where (courseID = " + Session("Course") + ") and (ID = " + IDTrab[i] + ")";
              oRec = oComm.Execute();
              oConn.Close();          

              //borro el fichero fisico del curso
              var killer = new ActiveXObject("Scripting.FileSystemObject");
              if (killer.FileExists(Server.MapPath("../courses/" + "course" + Session("Course") + "/Seminarios/" + URLs[i])) == true)
                {
                  killer.DeleteFile(Server.MapPath("../courses/" + "course" + Session("Course") + "/Seminarios/" + URLs[i]),true);
                }
              killer = null;  
              
            }
        } 

  Response.Redirect('AdminUpLoadedFiles.asp?uid=' + uid); 

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
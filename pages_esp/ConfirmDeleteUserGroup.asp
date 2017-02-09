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

if (Session("PermissionType") == ADMINISTRATOR)
    {
      var  filePath = Application("dataPath");   
      var  oConn = Server.CreateObject("ADODB.Connection"); 
      var  oComm    = Server.CreateObject("ADODB.Command");
      oConn.Open(filePath);
      oComm.ActiveConnection = oConn;


      var Tot = parseInt(Request.Form("Count"));
      var i = 0;

      for (var i = 1; i <= Tot; i++) {

        if ((Request.Form("check" + i).Count > 0) && (Request.Form("UserID" + i).Count > 0) && (Request.Form("UserID" + i).Count > 0) && (Request.Form("UserID" + i) != GUEST_USER) ) {            {

              oComm.CommandText = "delete  from Grupos_de_Usuarios where ([User] = " + Request.Form("UserID" + i) + ") and ([Group] = " + Request.Form("GroupID" + i) + ")";
              oComm.Execute();
            }
        }
      } 
              
      oConn.Close();          
      
      Response.Redirect("DeleteUserGroup.asp?uid=" + uid);      

    }
  else
    {  
      Response.Redirect("errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);                         
    }

  }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
  //Response.Write(Request.QueryString.Item("uid"));
%>

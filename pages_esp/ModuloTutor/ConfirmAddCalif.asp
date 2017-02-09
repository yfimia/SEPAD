<%@ Language=JScript %>
<%
  Response.Expires = -1;
  Response.Buffer = true;  
%>
<!-- #include file='../../js/adoLibrary.inc' -->
<!-- #include file='../../js/user.inc' -->

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       
      if (Request.QueryString.Item("curso").count > 0)
        var curso = Request.QueryString.Item("curso");
      if (Request.QueryString.Item("alumn").count > 0)
        var alumn = Request.QueryString.Item("alumn");
      if ((Session("isTutor") > -1) || (Session("PermissionType") == ADMINISTRATOR) || (Session("userID") == Session("tutcursocordinador"))  || (Session("userID") == Request.QueryString.Item("courseOwner")))
       {
         var ID = Request.Form("ID");
         var  filePath = Application("dataPath");   
         var  oConn    = Server.CreateObject("ADODB.Connection");
         var  oRec     = Server.CreateObject("ADODB.Recordset");
         oConn.Open(filePath);
         if (ID == -1)
         {
           oRec.Open("SELECT * From Calif_Ofic", oConn, 3, 3);
           oRec.AddNew();
           oRec.Fields.Item("curso").value = curso;
           oRec.Fields.Item("alumn").value = alumn;
           oRec.Fields.Item("prof").value = Session("userID");
         }
         else
         {
           oRec.Open("SELECT * From Calif_Ofic WHERE ID=" + ID, oConn, 3, 3);
         }
         var sendedtexto = Request.Form("COtitulo") + "";
         re = /</g;             
         sendedtexto = sendedtexto.replace(re,"&lt;");    
         re = />/g;             
         sendedtexto = sendedtexto.replace(re,"&gt;");    
         oRec.Fields.Item("titulo").value = sendedtexto;
         var sendedtexto = Request.Form("COobs") + "";
         re = /</g;             
         sendedtexto = sendedtexto.replace(re,"&lt;");    
         re = />/g;             
         sendedtexto = sendedtexto.replace(re,"&gt;");    
         oRec.Fields.Item("obs").value = sendedtexto;
         oRec.Fields.Item("calif").value = Request.Form("COcalif");
         oRec.Update();
         oRec.Close();
         Response.Redirect("verCalifOfic.asp?uid=" + uid + "&courseOwner=" + Request.QueryString.Item("courseOwner") + "&curso=" + curso + "&alumn=" + alumn);
       }
     else
       {  
         Response.Redirect("../errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);     
       }
   }
%>

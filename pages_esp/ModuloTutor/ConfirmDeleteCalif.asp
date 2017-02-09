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
      if ((Session("isTutor") > -1) || (Session("PermissionType") == ADMINISTRATOR) || (Session("userID") == Session("tutcursocordinador"))  || (Session("userID") == Request.QueryString.Item("courseOwner")))
       {
         var curso = Request.Form("curso");
         var alumn = Request.Form("alumn");
         var i = 0;
         var j = 0;
         var IDs = new Array();
         var dataPath = Application('dataPath');
         var oRec  = Server.CreateObject("ADODB.Recordset");
         var oComm = Server.CreateObject("ADODB.Command");
         var oConn = MakeConnection( oConn, dataPath );

          Sql = "SELECT * " + 
                "FROM Calif_Ofic " +
                "WHERE (curso = " + curso + ") and (alumn = " + alumn + ")";
          oRec = Query( Sql, oRec, oConn  );		

        while (oRec.EOF == false)
        {
          i = i + 1;
          if (Request.Form("check" + i) == "on") 
            { 
              j = j + 1;
              IDs[j] = oRec.Fields.Item("ID").Value;
            }
          oRec.Move(1);
        }
        oConn.Close();
         if (j >= 1)
           {
             for (i=1; i<=j; i++) 
               {
                 oConn.Open(dataPath);
                 oComm.ActiveConnection = oConn; 
                 oComm.CommandText = "delete from Calif_Ofic where ID=" + IDs[j];
                 oComm.Execute();
                 oConn.Close();          
               }
           } 
           Response.Redirect("verCalifOfic.asp?uid=" + uid + "&courseowner=" + Request.QueryString.Item("courseOwner") + "&curso=" + curso + "&alumn=" + alumn);
       }
     else
       {  
         Response.Redirect("../errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);     
       }
   }
%>

<%@ Language=JavaScript %>

<%
 Response.Expires = -1;
%>
<!-- #include file="../js/Adolibrary.inc" -->
<!-- #include file="../js/user.inc" -->


<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       
//Página para para el cambio del password....

  
   
  if (Request.Form.Count >= 2)
    { 
     //Se entra a la pagina desde la pagina de cambiar contraseña......
     var oldpassword            = Request.Form("useroldpassword") + "";
     var newpassword            = Request.Form("userpassword") + "";    
     //Response.Write(newpassword  + " == " +  oldpassword);     
     Session("userNewPassword") = newpassword;
     Session("userOldPassword") = oldpassword;
    }
  else
    {
     Response.Redirect("errorpage.asp?tipo=Error&short=" + BAD_PARAMS_SHORT  + "&desc=" + BAD_PARAMS_TEXT);    
    }   
    
    
  password = Session("userPassword");
  
  var filePath = Application('dataPath');
  var oConn;
  var oRec;
 
  oConn = MakeConnection( oConn, filePath );
  Sql = "select * from Usuarios where (ID=" + Session("userID") + ")";
  


  //Response.Write(password ==  oldpassword );
  if (password == oldpassword)
       {     
        oRec = Query( Sql, oRec, oConn  );
        oRec.Fields.Item("Password").value = newpassword;
        oRec.Update();
        Session("userPassword") = newpassword;
        DestroyAdoObjects( oConn, oRec );
        
        Response.Redirect("changepassword.asp?uid=" + uid);        
       }
     else
       {
        //No coincide el password anterior entrado con el de la Base de Datos........
        oConn.Close();  
        Response.Redirect("errorpage.asp?tipo=Error&short=" + INVALID_PASSWORD_SHORT  + "&desc=" + INVALID_PASSWORD_TEXT);            
       }   
     
  

%>

<html>
<head>
<title>Contraseña cambiada satisfactoriamente</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<%
  
  
  
 %>

</BODY>
</HTML>
<%
  }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
  //Response.Write(Request.QueryString.Item("uid"));
%>
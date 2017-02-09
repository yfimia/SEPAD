<%@ Language=JScript %>
<%
  Response.Buffer = true;
%>
<!-- #include file='../js/user.inc' -->
<!-- #include file='../js/adolibrary.inc' -->
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

<p>
<%

  var group = -1;
  if (Session("PermissionType") == ADMINISTRATOR)
    {
    if (Request.Form("passw") + "" == Request.Form("confirm") + "") 
      {
        var  filePath = Application("dataPath");   
        var  oConn    = Server.CreateObject("ADODB.Connection");
        var  oComm    = Server.CreateObject("ADODB.Command");
        var  oRec     = Server.CreateObject("ADODB.Recordset");
        oConn.Errors.Clear();
        oConn.Open( filePath );
  //      if (Request.Form("nick").Count > 0) 
  //        {
             oRec.Open("select * from Usuarios where (name = '" + Request.Form("nick") + "')",oConn,3,3);
             if (Request.Form("fullname") + "" != oRec.Fields.Item("FullName").Value)
               { 
                 oRec.Fields.Item("Fullname").Value = Request.Form("fullname") + ""; 
               }
               
             if ((Request.Form("passw").Count > 0) && (Request.Form("passw") != ""))
               {
                 oRec.Fields.Item("Password").Value = Request.Form("passw");
                 if (Session("name") == Request.Form("nick") ) 
                   {Session("userPassword") = Request.Form("passw") + "";}
               } 
             if ((Request.Form("email") != "") || (Request.Form("email").Count > 0)) 
               {
                 oRec.Fields.Item("email").Value = Request.Form("email") + "";
               }
               
             if (Session("name") == Request.Form("nick") ) 
              {
                Response.Write ("<font color=red size=2>El cambio de privilegio del usuario activo no es permitido</font><br>");  
              }
             else
               { 
                 if (Request.Form("TipoU") == "Administrador")
                   {
                     oRec.Fields.Item("PermissionType").Value = ADMINISTRATOR;
                     group = ADMIN_GROUP;
                   }
                 else  
                   {oRec.Fields.Item("PermissionType").Value = USER;}  
               }    
             
              //var fecha = new Date (parseInt(Request.Form("anio")), parseInt(Request.Form("mes")), parseInt(Request.Form("dia")));
	      oRec.Fields.Item("fechaNac") = Request.Form("mes") + "/" + Request.Form("dia") + "/"  + Request.Form("anio");
              
              if (Request.Form("sexo") == "m") {oRec.Fields.Item("sexo").Value = true;}
              else {oRec.Fields.Item("sexo").Value = false;} 
           
              oRec.Fields.Item("Institucion").Value = Request.Form("institucion") + "";                
              oRec.Fields.Item("telefono").Value = Request.Form("tel") + "";
              oRec.Fields.Item("direccion").Value = Request.Form("direccion") + "";
              oRec.Fields.Item("cpostal").Value = Request.Form("cp") + "";
              oRec.Fields.Item("titulo").Value = Request.Form("titulo") + "";   
              oRec.Fields.Item("Pais").Value = Request.Form("pais") + "";
              oRec.Fields.Item("Provincia").Value = Request.Form("prov") + "";
              
              oRec.Fields.Item("ci").Value = Request.Form("ci") + "";
              oRec.Fields.Item("profesion").Value = Request.Form("profesion") + "";
              oRec.Fields.Item("especialidad").Value = Request.Form("especialidad") + "";
              
              if (!isNaN(parseInt(Request.Form("annosgrad") + ""))) 
              {
                oRec.Fields.Item("annosgrad").Value = Request.Form("annosgrad") + "";
              }
              else
              {
                Response.Redirect("ModifyUser.asp?uid=" + uid + "&user=" + Request.QueryString.Item("user")) 
              }  
              
              oRec.Fields.Item("trabactual").Value = Request.Form("trabactual") + "";
              oRec.Fields.Item("cargoactual").Value = Request.Form("cargoactual") + "";
              oRec.Fields.Item("otros").Value = Request.Form("otros") + "";
              oRec.Fields.Item("catDoc").Value = Request.Form("catDoc") + "";
              oRec.Fields.Item("catInv").Value = Request.Form("catInv") + "";
              oRec.Fields.Item("catCient").Value = Request.Form("catCient") + "";
              
              if (Request.Form("forw") == "on") 
                {oRec.Fields.Item("forward").Value = true;}
              else
                {oRec.Fields.Item("forward").Value = false;}
              
             oRec.Update(); 
             oConn.Close();
             Response.Redirect("DeleteUser.asp?uid=" + uid);      
       //   }   
        Response.Write ('<script languaje=javascript>function pclick(){window.location="SelectUserToModify.asp";}</script>');      
        Response.Write ("<center><input type=Button Value=Regresar onclick='return pclick()' id=Button1 name=Button1></center>" ); 
      }
    else 
    {
      %><p><%=Request.Form("passw")%></p>
        <p><%=Request.Form("confirm")%></p>
        <p><%=Request.Form("nick")%></p><%
        
    }  
    }
  else  
    {
    Response.Redirect("errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);         
    }
  var  filePath = Application("dataPath");   
  var  oConn    = Server.CreateObject("ADODB.Connection");
  var  oComm    = Server.CreateObject("ADODB.Command");
  var  oRec     = Server.CreateObject("ADODB.Recordset");
  oConn.Errors.Clear();
  oConn.Open( filePath );
  oRec.Open("select * from Usuarios where (Name = '" + Request.Form("nick") + "')",oConn,3,3); 
    var IID = oRec.Fields.Item("ID").Value;
    if (group != -1) {AddUserToGroup(IID, group)};  
  oConn.Close();  
%>
<html>
<head>
<title>Matrículas satisfactorias</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/Main.css" type="text/css">
</head>
<body>
    
</body>
</html>
<%
  }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
  //Response.Write(Request.QueryString.Item("uid"));
%>
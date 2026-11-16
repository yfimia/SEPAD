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
function CloseWindow() {
	close();
}

 // opener.parent.location.reload(true);
</script>


<p>
<%

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
                  
             
              var fecha = new Date (parseInt(Request.Form("anio")), parseInt(Request.Form("mes")), parseInt(Request.Form("dia")));
	      oRec.Fields.Item("fechaNac") = Request.Form("dia") + "/" + Request.Form("mes") + "/"  + Request.Form("anio");
              
              if (Request.Form("sexo") == "m") {oRec.Fields.Item("sexo").Value = true;}
              else {oRec.Fields.Item("sexo").Value = false;} 

              if (Request.Form("forw") == "on") {oRec.Fields.Item("forward").Value = true;}
              else {oRec.Fields.Item("forward").Value = false;} 
           
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
                Response.Redirect("DatosPersonales.asp?uid=" + uid) 
              }  
              
              oRec.Fields.Item("trabactual").Value = Request.Form("trabactual") + "";
              oRec.Fields.Item("cargoactual").Value = Request.Form("cargoactual") + "";
              oRec.Fields.Item("otros").Value = Request.Form("otros") + "";
              oRec.Fields.Item("catDoc").Value = Request.Form("catDoc") + "";
              oRec.Fields.Item("catInv").Value = Request.Form("catInv") + "";
              oRec.Fields.Item("catCient").Value = Request.Form("catCient") + "";
              oRec.Fields.Item("skin").Value = Request.Form("skin") + "";
              Session("skin") = Request.Form("skin") + "";

              
             oRec.Update(); 
             oConn.Close();
             Response.Redirect("DatosPersonales.asp?uid=" + uid);      
       //   }   
        Response.Write ('<script languaje=javascript>function pclick(){window.location="DatosPersonales.asp?uid=" + uid;}</script>');      
        Response.Write ("<center><input type=Button Value=Regresar onclick='return CloseWindow()' id=Button1 name=Button1></center>" ); 
      }
    else 
    {
      %><p><%=Request.Form("passw")%></p>
        <p><%=Request.Form("confirm")%></p>
        <p><%=Request.Form("nick")%></p><%
    } 
    }           

%>
<html>
<head>
<title>Matrículas satisfactorias</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
</head>
<body>
    
</body>
</html>

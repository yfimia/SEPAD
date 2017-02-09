<%@ Language=JScript %>
<%
  Response.Expires = -1;
%>
<SCRIPT language=vbscript RUNAT=Server>
 function getAtualTime
  getAtualTime = Now()
 End function 

</SCRIPT>
<!-- #include file="../../js/adolibrary.inc" -->
<!-- #include file="../../js/library.inc" -->
<!-- #include file='../../js/user.inc' -->
<!-- #include file='inc/AceptaMtrCur.inc' -->
<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

%>

<%  

  if ((Session("PermissionType") == ADMINISTRATOR) || (Session("userID") == Session("admcordinador")))
    {
      var  filePath = Application("dataPath");   
      var  oRec  = Server.CreateObject("ADODB.Recordset");
      var  oRec2 = Server.CreateObject("ADODB.Recordset");
      var  oConn = Server.CreateObject("ADODB.Connection"); 
      var  oComm    = Server.CreateObject("ADODB.Command");
      
      //Cojo el grupo del curso
      groupid = Session("admgrupo");
      
      var grupo = new Array();
      oRec  = Session("UserList");
      oConn = Session("Conection");
      oComm = Session("Command");
      var i = 0;
      var j = 0;
      var IDMA = new Array();
      var IDMB = new Array();
      fullname = new Array();
      userid   = new Array();
      email	   = new Array();
      passw    = new Array();
      
      while (oRec.EOF == false)
        {
          i = i + 1;
          if (Request.Form("Check" + i) == "on") 
            {
              j = j + 1;
              IDMA[j] = oRec.Fields.Item("ID").Value;
              fullname[j] = oRec.Fields.Item("FullName").Value;
              userid[j] = oRec.Fields.Item("User").Value;
            }
          oRec.Move(1);
        }

      

			
   		oRec2.Open("SELECT * FROM Grupos_de_Usuarios",oConn,3,3);              
      // Acepto las subscripciones
      if (j >= 1)
        {
          for (i=1; i<=j; i++) 
            {
             if (empty("select * from Grupos_de_Usuarios where ([User] = " + userid[i] + ") and ( [Group] =" + groupid + ")")) 
             { 
              oRec2.AddNew();
              oRec2.Fields.Item("User").Value = userid[i];
              //Response.Write(oRec2.Fields.Item("User").Value);

              oRec2.Fields.Item("Group").Value = groupid;


              //Response.Write("<br>" + groupid[i]);

              oRec2.Update(); 
              
              SendMsg( 'SEPAD', userid[i], MatriculaAprobada + Session("admmoduloname"), MatModAcademica + Session("admmoduloname") + HaSidoAceptada + '\n' + GraciasPorSuInteres + '\n\n' + Mensaje + Session("admmoduloname") + ".");                    
             } 
              //borro las matriculas 

              oComm.ActiveConnection = oConn; 
              oComm.CommandText = "delete  from TMP_Matriculas_Cursos where (ID = " + IDMA[i] + ")";
              oRec = oComm.Execute();
//   Response.Write(IDMA[i] + "<br>");

            }
        }    
  Response.Redirect('MtrCurManager.asp?uid=' + uid);        
%>

<html>
<head>
<title><%=TituloPagina%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/Main.css" type="text/css"></head>

<body bgcolor="#FFFFFF" text="#000000">
<%

%>

<%          
        
       oRec2.Close();
      Session("UserList") = null;
      Session("Conection") = null;
      Session("Command") = null;
            //<script languaje=javascript>function pclick(){window.location="MtrCurManager.asp?course=' +course + '&coursename='+coursename + '" ;}</script>
%>                                                                                            

<%      
    }
  else
    {  
     Response.Redirect("../errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);    
    }

%>
    
</body>
</html>
<%
  }
 else
   Response.Redirect("../errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>

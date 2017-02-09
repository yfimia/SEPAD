b<%@ Language=JScript %>
<%
  Response.Expires = -1;
  Response.Buffer = true;  
%>

<!-- #include file='../../js/user.inc' -->
<!-- #include file='../../js/adolibrary.inc' -->
<!-- #include file='inc/ConfirmAgregaUserCur.inc' -->
<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       
  if ((Session("PermissionType") == ADMINISTRATOR) || (Session("userID") == Session("admcordinador")))
    {

      var  filePath = Application("dataPath");   
      var  oRec  = Server.CreateObject("ADODB.Recordset");
      var  oRec1  = Server.CreateObject("ADODB.Recordset");
      var  oConn = Server.CreateObject("ADODB.Connection"); 
      var  oComm    = Server.CreateObject("ADODB.Command");
      var grupo = new Array();
      oRec  = Session("UserList");
      oConn = Session("Conection");
      oComm = Session("Command");
      var i = 0;
      var j = 0;
      var fname = new Array();
      var IDU = new Array();
      var nick = new Array();
      while (oRec.EOF == false)
        {
          i = i + 1;
          if (Request.Form("Check" + i) == "on") 
            {
              j = j + 1;
              fname[j] = oRec.Fields.Item("FullName").Value + "";
              IDU[j] = oRec.Fields.Item("ID").Value;
              nick[j] = oRec.Fields.Item("Name").Value;
            }
          oRec.Move(1);
        }
       oRec.Close();
       oConn.Close();          
       
       if (j >= 1)
         {
           var IDG = Session("admgrupo");

           for (i=1; i<=j; i++) 
             {
               AddUserToGroup(IDU[i], IDG);
             }
               
             
         }  
      
 Response.Redirect('MtrUsrCour.asp?uid=' + uid + '&course=' + Session("course") + '&coursename=' + Session("admcourseName"));

%>

<html>
<head>
<title><%=TituloPagina%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/Main.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000">

<%
%>
<SCRIPT LANGUAGE=javascript>
<!--
function muOnclick()
 {
   location = 'MtrUsrCour.asp?uid=<%=uid%>&course=<%=Session("admcourse")%>&coursename=<%=Session("admcourseName")%>' ;
 }
 
//-->
</SCRIPT>

 <center><b><%=AlumnosMatriculados%></b></center>
 <center>    <INPUT align=center id=Buton name=Buton type=Button value=<%=Regresar%>  onclick=muOnclick()></center>
<%        
         
       Session("UserList") = null;
       Session("Conection") = null;
       Session("Command") = null;
     
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
  //Response.Write(Request.QueryString.Item("uid"));
%>
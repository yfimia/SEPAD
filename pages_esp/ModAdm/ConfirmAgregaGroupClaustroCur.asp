b<%@ Language=JScript %>
<%
  Response.Expires = -1;
  Response.Buffer = true;  
%>

<!-- #include file='../../js/user.inc' -->
<!-- #include file='../../js/adolibrary.inc' -->
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
      var k = 0;
      var fname = new Array();
      var IDU = new Array();
      var IDG = new Array();
      var nick = new Array();
      while (oRec.EOF == false)
        {
          i = i + 1;
          if (Request.Form("Check" + i) == "on") 
            {
              j = j + 1;
              
              IDG[j] = oRec.Fields.Item("ID").Value;
              
            }
          oRec.Move(1);
        }
       oConn.Close();

       if (j >= 1)
         {
           k = 0;
           for (i=1; i<=j; i++)
           {
            oConn.Open(filePath); 
            oRec.Open("select * from Grupos_de_Usuarios where ([Group] = " + IDG[i] + ")",oConn,3,3);
            while (oRec.EOF == false)
            {
             k = k + 1;
             IDU[k] = oRec.Fields.Item("User").Value;
             oRec.Move(1);
            }
            oConn.Close();   
           }
         
         if (k >= 1)
          {  
           oConn.Open(filePath);
           oRec.Open("select * from Grupos where (ID = " + Session("admclaustro") + ")",oConn,3,3);
           var IDGrupo = oRec.Fields.Item("ID").Value;
           oConn.Close();
                    
           oConn.Open(filePath);
           oRec.Open("select * from Grupos_de_Usuarios",oConn,3,3);
                 
           for (i=1; i<=k; i++) 
             {
               if (empty("select * from Grupos_de_Usuarios where ([User] = " + IDU[i] + ") and ([Group] =" + IDGrupo + ")")) 
               {
                 oRec.AddNew();
                 oRec.Fields.Item("User").Value = IDU[i];
                    
                 //Seleccionar el ID del grupo
                      
                 oRec.Fields.Item("Group").Value = IDGrupo;
                 oRec.Update();
                Response.Write(IDU[i] + " ");       
               }
          //    Response.Write(IDU[i] + " ");  
             }
                                  
               oConn.Close();          
          }   
         }
            
      
 Response.Redirect('MtrGClaustroCour.asp?uid=' + uid);

%>

<html>
<head>
<title><%=TituloPagina%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/Main.css" type="text/css">

<body bgcolor="#FFFFFF" text="#000000">

<%
%>
<SCRIPT LANGUAGE=javascript>
<!--
function muOnclick()
 {
   location = 'MtrGClaustroCour.asp?uid=<%=uid%>' ;
 }
 
//-->
</SCRIPT>

 <center><b><%=GruposAsignados%></b></center>
 <center>    <INPUT align=center id=Buton name=Buton type=Button value=<%=Regresar%>  onclick=muOnclick()></center>
<%        
         
       Session("UserList") = null;
       Session("Conection") = null;
       Session("Command") = null;
     
    }
  else
    {  
      Response.Redirect("../../errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);     
    }
%>
</body>
</html>
<%
  }
 else
   Response.Redirect("../../errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
  //Response.Write(Request.QueryString.Item("uid"));
%>
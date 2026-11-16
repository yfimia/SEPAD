<%@ Language=Jscript%>
<%

/*
  Esta pagina pone en la base de datos un nuevo correo que proviene de 
  los Foros de discusión
  
  - Los datos vienen en una forma que se postea desde la pagina NewForumEmail.asp
*/  

Response.Expires = -1;
%>

<!-- #include file='../js/user.inc' -->
<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

%>

<SCRIPT language=vbscript RUNAT=Server>
 function getDate
  getDate = Now()
 End function 
</SCRIPT>

<% 

 function Init()
  {
   var i = 0;         
      
   while (MailTo.indexOf(",") != -1)
    {     
     To[i]  = parseInt( MailTo.substr(0, MailTo.indexOf(",")) );
     MailTo = MailTo.slice(MailTo.indexOf(",") + 1);                
     i++;
    }
   
   To[i]    = parseInt( MailTo );      
  }
   
 MailTo       = new String(Request.Form("MailToText") + " ");
 MailBody     = new String(Request.Form("MailBody") + " ");  
 MailSubject  = new String(Request.Form("MailSubject") + " ");
 MailPriority = Request.Form("MailPriority")
 
 
 var oConn    = Server.CreateObject("ADODB.Connection");  
 var oRec     = Server.CreateObject("ADODB.Recordset");
 var filePath = Application("filePath");  
 var Sql      = "SELECT * FROM Mensajeria";
 var To       = new Array();
                 
 Init();
 
 oConn.Open( filePath );   
 oRec.Open(Sql, oConn, 3, 3);    
 
 //Entrada de datos de la mensajeria
 
 for (i=0;i <= To.length-1;i++)
  {
   oRec.AddNew();
   oRec.Fields.Item("FromID").value     = Session("userid");
   oRec.Fields.Item("Readed").value     = false;
   oRec.Fields.Item("From").value     = Session("fullName")
   oRec.Fields.Item("To").value       = To[i]; 
   oRec.Fields.Item("Subject").value  = MailSubject; 
   oRec.Fields.Item("Priority").value = MailPriority;
   oRec.Fields.Item("Content").value  = MailBody;
   oRec.Fields.Item("Date").value     = getDate();      
  }
  
  oRec.Update();
  
 //*********************************
 
 oConn.Close(); 
 Response.Redirect ("inbox.asp?uid=" + uid);
%>
<HTML>
<HEAD>
</HEAD>
<BODY>
</BODY>
</HTML>
<%
 }
 else //Response.Write(Request.QueryString.Item("uid"));
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>

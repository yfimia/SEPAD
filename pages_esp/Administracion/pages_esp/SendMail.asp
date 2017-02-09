<%@ Language=Jscript%>
<%
Response.Expires = -1;
%>
<!-- #include file='../js/adolibrary.inc' -->
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
   if (MailTo.charAt(0) == '0') {   // Un grupo
     var filePath = Application('dataPath');
     var oConn;
     var oRec;
     var result = "";
 
     oConn = MakeConnection( oConn, filePath );
     Sql = "SELECT Grupos_de_Usuarios.[User], Usuarios.Name, Usuarios.FullName " +
           "FROM Usuarios INNER JOIN Grupos_de_Usuarios ON Usuarios.ID = Grupos_de_Usuarios.[User] " +
           "WHERE (((Grupos_de_Usuarios.[Group])= " + MailTo.substr(2) + "))"
     oRec = Query( Sql, oRec, oConn  );		
     while (oRec.EoF == false) 
       {     
        //se codifican los valores de los option de la forma [0 grupo/1 usuario] + indice   
        ToID[i] = oRec.Fields("User").Value;
        i++;
        oRec.Move(1);
       }

       
     DestroyAdoObjects( oConn, oRec );
   
   }
   else { //Un usuario
     ToID[0] = MailTo.substr(2);

   }
   
  }
   
 MailTo       = Request.Form("MailTo") + "";
 MailToText   = Request.Form("MailToText") + "";
 MailBody     = Request.Form("MailBody") + "";  
 MailSubject  = Request.Form("MailSubject") + "";
 MailPriority = Request.Form("MailPriority");
 
 //Response.Write(MailTo);
 var oConn    = Server.CreateObject("ADODB.Connection");  
 var oRec     = Server.CreateObject("ADODB.Recordset");
 var filePath = Application("filePath");  
 var ToID       = new Array();
                 
 Init();
 
 oConn.Open( filePath );   
 var Sql      = "SELECT * FROM Mensajeria";
 oRec.Open(Sql, oConn, 3, 3);    
 
 //Entrada de datos de la mensajeria
 var IN = '';
 
 for (i=0;i <= ToID.length-1;i++)
  {
   //Response.Write(ToID[i]);
   oRec.AddNew();
 
   if (i != ToID.length - 1 ) IN = IN + ToID[i] + ','; else IN = IN + ToID[i];
     
   oRec.Fields.Item("FromID").value     = Session("userid");
   
   oRec.Fields.Item("Readed").value     = false;
   oRec.Fields.Item("From").value     = Session("fullName")
   oRec.Fields.Item("To").value       = parseInt(ToID[i], 10); 
   oRec.Fields.Item("ToText").value       = MailToText; 
   oRec.Fields.Item("Subject").value  = MailSubject; 
   oRec.Fields.Item("Priority").value = MailPriority;
   oRec.Fields.Item("Content").value  = MailBody;
   oRec.Fields.Item("Date").value     = getDate();      
  }
  
  oRec.Update();
  oRec.Close();
  //********************************
  
  var Mail = Server.CreateObject("RWMS.RWMailSender");

  var Sql      = "SELECT email FROM Usuarios WHERE (ID IN (" + IN + ")) and (forward = " + Application("dtrue") + ") and  (NOT(email = ''))";
  oRec.Open(Sql, oConn, 3, 3);    

  while (oRec.EOF == false) {
    //Envio correo
    try { 
  
      Mail.Send('SEPAD', Application("MailAddress"), oRec.Fields.Item("email").value, 'SEAPD: ' + MailSubject,
                 "Mensaje en SEPAD de: " + Session("fullName") + "\n" + 
                 "Asunto : "  + MailSubject + "\n" + MailBody , 
                 Application("MailServer"),
                Application("MailUser"), Application("MailPassword"), parseInt(Application("Port"),10), parseInt(Application("Auth"),10));
    } 
    catch(e) {		
      Response.Write("*** No pudo enviarse correo de confirmación al usuario *** <br>"); 
   }                           
   
   oRec.Move(1);
  }
  oRec.Close();
  Mail = null;

  
 
 oConn.Close(); 

 
%>


<HTML>
<HEAD>
</HEAD>

<BODY>

<SCRIPT LANGUAGE=JSCRIPT>
  //alert(opener);
  //OJO
  if  ((opener  != null) && (opener + "" != "undefined") && (opener + ""  != "null") && (opener  != 0))
    opener.location = "inbox.asp?uid=<%=uid%>";
  close();
</SCRIPT>

</BODY>
</HTML>
<%
 }
 else //Response.Write(Request.QueryString.Item("uid"));
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>

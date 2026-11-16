<%@ Language=JScript %>
<%
  Response.Expires = -1;
  Response.Buffer = true;  
%>
<!-- #include file='../js/user.inc' -->
<!-- #include file='inc/confirmaceptamatricula.inc' -->

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
<body>

<p>
<%
if (Session("PermissionType") == ADMINISTRATOR)
    {
      var  filePath = Application("dataPath");   
      var  oRec  = Server.CreateObject("ADODB.Recordset");
      var  oConn = Server.CreateObject("ADODB.Connection"); 
      var  oComm    = Server.CreateObject("ADODB.Command");
      var grupo = new Array();
      oRec  = Session("UserList");
      oConn = Session("Conection");
      oComm = Session("Command");
      var i = 0;
      var j = 0;
      var k = 0;
      var IDMA = new Array();
      var IDMB = new Array();
      fullname = new Array();
      nick	   = new Array();
      email	   = new Array();
      passw    = new Array();

      while (oRec.EOF == false)
        {
          i = i + 1;
          if (Request.Form("Check" + i) == "on") 
            {
              j = j + 1;
              IDMA[j] = oRec.Fields.Item("ID").Value;
              fullname[j] = oRec.Fields.Item("Fullname").Value;
              nick[j] = oRec.Fields.Item("Name").Value;
              email[j] = oRec.Fields.Item("Email").Value;
              passw[j] = oRec.Fields.Item("Password").Value;
            }
          else 
            {
              k = k + 1;
              IDMB[k] = oRec.Fields.Item("ID").Value;
            } 
          oRec.Move(1);
        }
      oConn.Close();   
      
  
      // Acepto las subscripciones
      if (j >= 1)
        {
          for (i=1; i<=j; i++) 
            {
              //Inserto el usuario
              oConn.Open( filePath );
              
              oRec.Open("select * from usuarios where (name = '" + nick[i] + "')",oConn,3,3);
              if (oRec.EOF == false)
                    { oRec.Close();
                      oConn.Close();
                      Response.Redirect("errorpage.asp?tipo=Error&short=" + LOGIN_EXIST_SHORT + "(" + nick[i] + ")"  + "&desc=" + LOGIN_EXIST_TEXT);     
                    }
              oRec.Close();
              oRec.Open("select * from Usuarios",oConn,3,3);
              oRec.AddNew();
              oRec.Fields.Item("FullName").Value = fullname[i] + "";
              oRec.Fields.Item("Name").Value = nick[i] + "";
              oRec.Fields.Item("password").Value = passw[i] + "";
              if ((email[j] == "") || (email[j] == null)) 
                //{oRec.Fields.Item("email").Value = nick[i] + "@Sepad.cu";}
                {oRec.Fields.Item("email").Value = nick[i] + " ";}
              else
                {oRec.Fields.Item("email").Value = email[i] + "";}  
              oRec.Fields.Item("Logins").Value = 0;
              oRec.Fields.Item("PermissionType").Value = 0;  
              oRec.Update(); 
              oConn.Close(); 
              
              //Envio correo
               try { 

                 var Mail = Server.CreateObject("RWMS.RWMailSender");
                 Mail.Send('SEPAD', Application("MailAddress"), email[i], Application("Subject"),
                            stmp1 + nick[i] + ":\n" + Application("Body") , Application("MailServer"),
                           Application("MailUser"), Application("MailPassword"), parseInt(Application("Port"),10), parseInt(Application("Auth"),10));
                 Mail = null;
               } 
               catch(e) {		
                 Response.Write(msg1); 
               }                           
              
              //Recojo el Id que obtuvo
              oConn.Open( filePath );
              oRec.Open("select * from Usuarios where (Name = '" + nick[i] + "')",oConn,3,3);
              var IID = 0;
              if (oRec.EOF == false)
                { IID = oRec.Fields.Item("ID").Value; } 
              oConn.Close();   
              //Inserto la pertenecia a SEPAD
              oConn.Open( filePath );
              oRec.Open("select [User] as urs, [Group] as grp from Grupos_de_Usuarios",oConn,3,3);
              oRec.AddNew();
              oRec.Fields.Item("urs").Value = IID;
              oRec.Fields.Item("grp").Value = SEPAD_GROUP;
              oRec.Update();      
              oConn.Close();   
              //borro las matriculas 
              oConn.Open(filePath);
              oComm.ActiveConnection = oConn; 
              oComm.CommandText = "delete  from matricula where (ID = " + IDMA[i] + ")";
              oRec = oComm.Execute();
              oConn.Close();          
            }
          
      
        }  
      Session("UserList") = null;
      Session("Conection") = null;
      Session("Command") = null;
      
      Response.Redirect("MatriculasManager.asp?uid=" + uid);      
      
      Response.Write ('<script languaje=javascript>function pclick(){window.location="MatriculasManager.asp";}</script>');      
      Response.Write ("<center><input type=Button Value=" + stmp2 + " onclick='return pclick()' id=Button1 name=Button1></center>" );     
    }
  else
    {  
     Response.Redirect("errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);                             
    }

%>
<%
  }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
  //Response.Write(Request.QueryString.Item("uid"));
%>

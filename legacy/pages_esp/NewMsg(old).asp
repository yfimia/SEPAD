<%@ Language=jScript %>

<%
Response.Expires = -1;
%>

<script language=vbscript  runat=Server>
function getDate
getDate = Now()
End function 
</script>

<!-- #include file='../js/user.inc' -->
<%
   var      MailTo = "";
   var      MailToText = "";
   var      MailBody = "";
   var      MailSubject = "";
   var	    reply = false;
   var      forward = false;

    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       
      if (Request.Form.Count >= 4) {
         MailToText = Request.Form("MailToName");
         MailTo = Request.Form("MailToText");
         MailBody = Request.Form("MailBody");
         if (MailTo == '-1') {
           forward = true;
         if (forward)
           MailSubject = "Fw: " + Request.Form("MailSubject");
           
         }
         reply = true;
      }
         if (reply && !forward)
           MailSubject = "Re: " + Request.Form("MailSubject");


%>



<%

  function List()
   {     
    function addItem( Obj )
     {
      this.Count++;
      this.Items[ this.Count ] = Obj; 
     }
       
     this.Count   = -1;
     this.Items   = new Array();
     this.addItem = addItem;
    }
    
  function TUser( Id, Name )
   {
    this.Id   = Id;
    this.Name = Name;
   }  
  
  function TGroup( Id, Name, Description )
   {
    this.Id          = Id;
    this.Name        = Name;
    this.Description = Description;
    this.Users       = new List();            
   }
    
    
  function FindGroup( ind )
   {
    var i = 0, ret = -1;   

    for (i=0;i <= GroupList.Count;i++)
     {
      if (GroupList.Items[i].Id == ind) 
       {ret = i;i = GroupList.Count}
     }    
         
    return ret
   }  
   
   
  var HTML     = "";    
  var oConn    = Server.CreateObject("ADODB.Connection");  
  var oRec     = Server.CreateObject("ADODB.Recordset");
  var filePath = Application("filePath");
  
  var Sql      = "SELECT A.Id as GroupId, A.Name as GroupName, A.Description as GroupDesc,  C.Id as UserId, C.FullName as UserName " + 
                 "FROM grupos AS A INNER JOIN (grupos_de_usuarios AS B INNER JOIN Usuarios AS C ON B.[User] = C.Id) ON A.id = B.[group] " +
                 "WHERE (C.ID <> " + GUEST_USER + ")" + 
                 "ORDER BY C.Fullname";
        //Response.Write(Sql);         
  var GroupList = new List();
  var UserList  = new List();                 
  
  
  
  oConn.Open( filePath );  
  oRec = oConn.Execute(Sql);  

  var lastUserId = -1;  
  
  while (oRec.EoF == false) 
   {     
    UserId     = oRec.Fields("UserId").Value;    
    GroupId    = oRec.Fields("GroupId").Value;    
    
    
    GroupIndex = FindGroup( GroupId );    
           
    if (UserId != lastUserId)
     {
      //************* Encuentro con un nuevo usuario      
      User       = new TUser( UserId, oRec.Fields("UserName").Value);         
      lastUserId = UserId;
      UserList.addItem ( User );
     }
     
    if (GroupIndex == -1)
     {
      //************* Encuentro nuevo grupo
      Group      = new TGroup(GroupId, oRec.Fields("GroupName").Value, oRec.Fields("GroupDesc").Value);
      GroupList.addItem ( Group );  
      GroupIndex = GroupList.Count;
     }
            
    GroupList.Items[ GroupIndex ].Users.addItem( UserId );        
    oRec.Move(1);
   }
    
  oConn.Close();  
  
  var HTML_SELECT = "";
  var SCRIPT      = "var GroupList = new List();" + 
                    "var UserList  = new List();" + 
                    "var Date      = '" + getDate() + "';" //Script del cliente
                    
  
  
  //se codifican los valores de los option de la forma [0 grupo/1 usuario] + indice
  for (i=0;i <= GroupList.Count;i++)
   {
    HTML_SELECT = HTML_SELECT + "<OPTION VALUE='0"+ i +"'>" + GroupList.Items[i].Name + " (" + GroupList.Items[i].Description + ")";
    
    // Se mandan a crear los grupos y los miembros de los grupos
    // en el cliente
    
    SCRIPT      = SCRIPT + "Group = new TGroup("+ GroupList.Items[i].Id +");";            
    
    for (j=0;j <= GroupList.Items[i].Users.Count; j++)
     {
      SCRIPT    = SCRIPT + 
                  "User = new TUser(" + GroupList.Items[i].Users.Items[j] +");" +     
                  "Group.Users.addItem( User );";
     } 
        
     SCRIPT     = SCRIPT + "GroupList.addItem( Group );";
   }
   
  for (i=0;i <= UserList.Count;i++)
   {
    // Creacion de los usuarios del cliente
    HTML_SELECT = HTML_SELECT + "<OPTION VALUE='1"+ i +"'>" + UserList.Items[i].Name;
    SCRIPT      = SCRIPT + "User = new TUser("+ UserList.Items[i].Id +");";
    SCRIPT      = SCRIPT + "UserList.addItem( User );";
    
   } 
%>
<html>
<head>
<title>Nuevo Mensaje</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<script language=jscript src ="../js/library.js"></script>
<script language=jscript>
<!--
  function List()
   {     
    function addItem( Obj )
     {
      this.Count++;
      this.Items[ this.Count ] = Obj; 
     }
       
     this.Count   = -1;
     this.Items   = new Array();
     this.addItem = addItem;
    }
    
  function TUser( Id )
   {
    this.Id   = Id;    
   }  
  
  function TGroup( Id )
   {
    this.Id          = Id;
    this.Users       = new List();            
   }
  
  function _Send()
   {  
    if (notIsScript(Subject.value))
     {        
<%
if (!reply || forward)  {
%>     

    St      = new String( Select.value );          
    Ind     = parseInt( St.substr(1, St.length ));
    MailTo  = "";
            
    if (St.charAt(0) == '0') 
     {      
      if (GroupList.Items[Ind].Users.Count > -1) 
       {        
        for (i=0;i <= GroupList.Items[Ind].Users.Count-1;i++)
        {MailTo = MailTo + GroupList.Items[Ind].Users.Items[i].Id + ","}
         MailTo  =  MailTo + GroupList.Items[Ind].Users.Items[ GroupList.Items[Ind].Users.Count ].Id;
       } 
     }
    else
     {
      MailTo = UserList.Items[Ind].Id;
     }
    Mail.MailToText.value   = MailTo;    
    Mail.MailBody.value     = Texto.value;      
    Mail.MailSubject.value  = Subject.value;
    Mail.MailPriority.value = Priority.checked == true?1:0
<%
  }
  else {
%>  
    Mail.MailBody.value     = Texto.value;      
    Mail.MailSubject.value  = Subject.value;
    Mail.MailPriority.value = Priority.checked == true?1:0
<%      
  }  
%>    
    Mail.submit();      
   }
   else
     alert('Asunto no válido, revise la presencia de caráteres especiales');
   }
   
  function _SendToAll()
   {
    if (notIsScript(Subject.value))
     {
    MailTo = "";
    
    for (i=0;i <= UserList.Count-1;i++)
     {MailTo = MailTo + UserList.Items[i].Id + ","}
      MailTo = MailTo + UserList.Items[ UserList.Count ].Id; 
              
    
    Mail.MailToText.value   = MailTo;  
    Mail.MailBody.value     = Texto.value;          
    Mail.MailSubject.value  = Subject.value;
    Mail.MailPriority.value = Priority.checked == true?1:0;    
    Mail.submit();   
    }
   else
     alert('Asunto no válido, revise la presencia de caráteres especiales');
   } 
   
  <%=SCRIPT%>   
//-->
</script>

</head>

<body bgcolor="#FFFFFF" text="#000000">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="MessageBarTD"> 
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td><a href="javascript:_Send()" class="ToolLink">&nbsp;Enviar&nbsp;</a></td>
          <td>|</td>
          <!--td><a href="javasript:_SendToAll()" class="ToolLink">&nbsp;Enviar&nbsp;a&nbsp;todos&nbsp;</a></td>
          <td>|</td-->
          <td><a href="javascript:close()" class="ToolLink">&nbsp;Cerrar&nbsp;</a></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td class="MessageTR1"> 
            <table border="0" cellspacing="0" cellpadding="0" width="100%">
              <tr> 
                <td width="1%">Para:</td>
                <td width="99%"> 
<%
  if (reply && !forward) {	           
%> 
                     <%=MailToText %>            
<%
  }     
  else {
%>      
                  <select id="Select" name="select1" class="ComboBox">
                     <%=HTML_SELECT %>
                  </select>
<%
  }
%>                  
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td class="MessageTR1"> 
            <table border="0" cellspacing="0" cellpadding="0" width="100%">
              <tr> 
                <td width="1%">Asunto:</td>
                <td width="99%"> 
                  <input type="text" id=Subject name="textfield" class="Edit" size="50" value="<%=MailSubject%>">
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td class="MessageTR1"> 
            <textarea name="textfield2" ID=Texto rows="12" class="TextArea"><%=MailBody %></textarea>
          </td>
        </tr>
        <tr> 
          <td class="MessageTR1"> 
            <table border="0" cellspacing="0" cellpadding="1">
              <tr> 
                <td> 
                  <input type="checkbox" ID=Priority name="checkbox" value="checkbox">
                </td>
                <td>Establecer nivel de prioridad.</td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td class="MessageBarTD"> 
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td>&nbsp;&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<FORM ID=Mail NAME=Mail METHOD=POST ACTION="SendMail.asp?uid=<%=uid%>" >
  <INPUT TYPE=hidden NAME=MailToText ID=MailToText VALUE="<%=MailTo%>">
  <INPUT TYPE=hidden NAME=MailBody ID=MailBody VALUE="<%=MailBody%>">  
  <INPUT TYPE=hidden NAME=MailSubject ID=MailSubject VALUE="<%=MailSubject%>">
  <INPUT TYPE=hidden NAME=MailPriority ID=MailPriority VALUE=0>      
</FORM>
</body>
</html>
<%
 }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>

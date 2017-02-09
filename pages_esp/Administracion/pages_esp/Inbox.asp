<%@ Language=jScript %>
<%  
 Response.Expires = -1;
%>
<!-- #include file="../js/Adolibrary.inc" --> 
<!-- #include file='../js/user.inc' -->
<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       
    first = "-1";
    where = "";
    if (Request.QueryString.Count >= 2) {
      first = Request.QueryString.Item("first") + "";
      
      if (first != "-1") {
        where = " and (Date < '" +  first + "')";
      } 
    } 

%>


<% var  filePath = Application("dataPath");   
   var oConn      = Server.CreateObject ("ADODB.Connection");
   var oRec       = Server.CreateObject ("ADODB.Recordset");
   var SQL        = "DELETE FROM Mensajeria WHERE ";
   
   oConn.Open(filePath);

   if (Request.Form.Count > 0)  
     {
      //Procesamiento para el borrado de mensajes uso del parametro flag
      var Tot       = Request.Form("Count");
      var i;

            
      if (Request.Form ("flag") == 0) 
        {
         //Busqueda de los mensajes a borrar      
         
         
                 
         for (i=0;i <= Tot - 1; i++)                                         
           if ((Request.Form("INP" + i).Count > 0) && (Request.Form("ID" + i).Count > 0))
             {
              SQL = SQL + "(ID=" + Request.Form("ID" + i) + ") OR ";
             }              

            SQL = SQL + "(1=0)";
            oConn.Execute (SQL);


        }
        //**********************************            
      else  
        {      
         SQL = SQL + "(To=" + Session("UserID")+ ")";
         oConn.Execute (SQL);
        } 
            
     }
       //Response.Write(SQL);
   
   
   
   var Count         = 0;   
   var SQL           = "SELECT Top " + SHOW_CANT + " ID, Priority, Readed, [From], Subject, Date  " +
                       "FROM Mensajeria WHERE ([To]="+ Session("UserID")+") " +  where + " ORDER BY Date DESC, Readed DESC, Priority DESC";   
%>
<html>
<head>
<title>Bandeja de Entrada</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<script type="text/javascript" language="javascript" src="../js/Windows.js"></script>
<script src="../js/CheckBoxes.js" language="JavaScript"></script>

</head>

<body bgcolor="#FFFFFF" text="#000000">
<form id="Messages" method="POST" action="inbox.asp?uid=<%=uid%>">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="MessageBarTD"> 
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td><a href="javascript:abreVentana('NewMsg',400,253,'NewMsg.asp?uid=<%=uid%>', 'no', 'no')" class="ToolLink">&nbsp;Nuevo&nbsp;mensaje&nbsp;</a></td>
          <td>|</td>
          <td><a href="javascript:onDeleteMessage()" class="ToolLink">&nbsp;Borrar&nbsp;mensaje(s)&nbsp;</a></td>
          <td>|</td>
          <td><a href="#" class="ToolLink" onClick="CheckAll(GetObjectCollection( document, 'FORM'))">&nbsp;Seleccionar&nbsp;todos&nbsp;los&nbsp;mensajes&nbsp;</a></td>
          <td>|</td>
          <td><a href="javascript:onNext()" class="ToolLink">&nbsp;Próximos&nbsp;<%=SHOW_CANT%>&nbsp;</a></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td> 
      <table width="100%" cellpadding="0" cellspacing="1" border="0">
        <tr> 
          <td width="1%" class="MessageInfo">&nbsp;</td>
          <td width="1%" class="MessageInfo">&nbsp;</td>
          <td width="30%" class="MessageInfo"> Enviado por</td>
          <td width="48%" class="MessageInfo">Asunto</td>
          <td width="20%" class="MessageInfo">Recibido</td>
        </tr>
<%
oRec = oConn.Execute( SQL );
var Alert = new Array("","!"); 
var ChechInput = "";
var trclass = 'MessageTR';
//
var  last = "-1";          

while (oRec.EoF == false) {     
    last = oRec.Fields.Item("Date").value;          

    bold = '';
    if (oRec.Fields("Readed").Value == false) {
      bold = 'Style="font-weight:bolder"';
    }
	ChechInput = '<input type="checkbox"  name="INP' + Count + '" id="INP' + Count + '">  <INPUT TYPE=hidden NAME="ID' + Count + '" VALUE="' + oRec.Fields("ID").Value  + '">';
     	var fecha = new Date();
     	var Vinculo = '<a href="#" onclick="onDblclick(' + oRec.Fields ("ID").Value +')">';
     	fecha.setTime( Date.parse ( oRec.Fields.Item("Date").Value + "" ) );
     	var date = fecha.toLocaleString() + "";
     	if (trclass == 'MessageTR') {
     		trclass = 'MessageTR1'; 
     	} else {
		trclass = 'MessageTR';
	}	
%>

        <tr  <%=bold%> class="<%=trclass%>"  valign="middle"> 
          <td align=center width="1%"><%=ChechInput %></td>
          <td align=center width="1%" class="PriorTD"><%=Alert[oRec.Fields("Priority").Value] %></td>
          <td align=center width="30%"><%=Vinculo%><%=oRec.Fields("From").Value %></a></td>
          <td align=left width="48%"><%=oRec.Fields("Subject").Value%></td>
          <td align=center width="20%"><%=date%></td>
        </tr>
<%
    	Count++;       
    	oRec.Move(1); 
}  
oConn.Close();    
%>
  <INPUT TYPE=hidden NAME="Count" VALUE="<%=Count%>">

     </table>
    </td>
  </tr>
</table>
<table border="0" cellspacing="0" cellpadding="0" class="ToolBar" width="100%">
  <tr>
    <td>
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td><a href="javascript:abreVentana('NewMsg',400,253,'NewMsg.asp?uid=<%=uid%>', 'no', 'no')" class="ToolLink">&nbsp;Nuevo&nbsp;mensaje&nbsp;</a></td>
          <td>|</td>
          <td><a href="javascript:onDeleteMessage()" class="ToolLink">&nbsp;Borrar&nbsp;mensaje(s)&nbsp;</a></td>
          <td>|</td>
          <td><a href="#" class="ToolLink" onClick="CheckAll(GetObjectCollection( document, 'FORM'))">&nbsp;Seleccionar&nbsp;todos&nbsp;los&nbsp;mensajes&nbsp;</a></td>
          <td>|</td>
          <td><a href="javascript:onNext()" class="ToolLink">&nbsp;Próximos&nbsp;<%=SHOW_CANT%>&nbsp;</a></td>
          
        </tr>
      </table>
    </td>
  </tr>
</table>
<INPUT TYPE=hidden NAME='flag' ID='flag' VALUE=1>  
</form>
<script language="javascript" >
<!--
   
function onDblclick( id ) {    
     event.srcElement.parentElement.parentElement.style.fontWeight  = "lighter";
     abreVentana('',400,240,'ViewMsg.asp?uid=<%=uid%>&ID=' + id + '&Flag=1&adm=1', 'no', 'no')
}  

function onNext() {    
     location.href = "inbox.asp?uid=<%=uid%>&first=<%=last%>";
}  
    
function onDeleteMessage() {      
     	Messages.all("flag").value = 0;
     	Messages.action = "inbox.asp?uid=<%=uid%>&first=<%=first%>";
     	Messages.submit();
}
    
function onDeleteAll() {    
    	Messages.all("flag").value = 1;
    	Messages.submit();
}    
-->          
</script>

</body>
</html>
<%
 }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>

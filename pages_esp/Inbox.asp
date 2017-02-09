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

 var   displayRecs = 12;
 var   recRange = 10
 // Get table name
 var  tablename = "[Mensajeria]";


      if (Request.QueryString("id").Count > 0) {
	    id=Request.QueryString("id");
	    Session("id") = id;
	  }  
      else {
	   id = Session("id");
	  } 
      

      //Load Default Order
      DefaultOrder = "Date";
      DefaultOrderType = "DESC";

      // Check for an Order parameter
      OrderBy = "";
      if (Request.QueryString("order").Count > 0) {
	    OrderBy = Request.QueryString("order");
	    // Check if an ASC/DESC toggle is required
	    if (Session("OrderBy") == OrderBy) {
		  if (Session("OrderType") == "ASC")
			Session("OrderType") = "DESC";
		  else
			Session("OrderType") = "ASC";
	    }  
	    else {
		  Session("OrderType") = "ASC";
	    }
	    Session("tablename") = tablename;
	    Session("OrderBy") = OrderBy + "";
	    Session("startRec") = 1;
	  }   
      else {
	    if (tablename == Session("tablename")) 
		  OrderBy = Session("OrderBy");
	    else {
		  OrderBy = DefaultOrder;
		  Session("OrderBy") = OrderBy + "";
		  Session("OrderType") = DefaultOrderType + "";
	    }
     }

     // Check for a START parameter
     if (Request.QueryString("start").Count > 0) {
	    startRec = parseInt(Request.QueryString("start"));
	    Session("tablename") = tablename + "";
	    Session("startRec") = startRec + "";
	 }    
     else {
	   if (tablename == Session("tablename")) 
		startRec =parseInt(Session("startRec") + "");
	   else {
		//reset start record counter
		startRec = 1;
		Session("startRec") = startRec;
	   }
     }

    //Set the last record to display
    stopRec = startRec + displayRecs - 1;

    // Build Query


    strsql = "select Mensajeria.ID, Priority, Readed, [From], Subject, Date   from " + 
             " Mensajeria  " ;
    dbwhere = " ([To]="+ Session("UserID")+") ";
    if (dbwhere != "") strsql = strsql + " WHERE " + dbwhere;
    if (OrderBy != "") strsql = strsql + " ORDER BY [" + OrderBy + "] " + Session("OrderType");

    var  filePath = Application("dataPath");   
    var  oConn    = Server.CreateObject("ADODB.Connection");
    var  oRec     = Server.CreateObject("ADODB.Recordset");
    oConn.Open(filePath);
    //Response.Write(strsql);
    //Response.End();

   if (Request.Form.Count > 0)  
     {
      var SQL        = "DELETE FROM Mensajeria WHERE ";
      //Procesamiento para el borrado de mensajes uso del parametro flag
      var Tot       = Request.Form("Count");
      var i;

            
      if (Request.Form ("flag") == 0) 
        {
         //Busqueda de los mensajes a borrar      
         
         
         for (i=1; i <= Tot; i++)                                         
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
   
    oRec.Open(strsql, oConn, 1, 2);
    totalRecs = oRec.RecordCount;

%>
<html>
<head>
<title>Bandeja de Entrada</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<script type="text/javascript" language="javascript" src="../js/Windows.js"></script>
<script src="../js/CheckBoxes.js" language="JavaScript"></script>

</head>

<body bgcolor="#FFFFFF" text="#000000" style="overflow-y: auto;">
<form id="Messages" method="POST" action="inbox.asp?uid=<%=uid%>">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="MessageBarTD"> 
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td><a href="javascript:abreVentana('NewMsg',800,440,'NewMsg.asp?uid=<%=uid%>', 'no', 'no')" class="ToolLink">&nbsp;Nuevo&nbsp;mensaje&nbsp;</a></td>
          <td>|</td>
          <td><a href="javascript:onDeleteMessage()" class="ToolLink">&nbsp;Borrar&nbsp;mensaje(s)&nbsp;</a></td>
          <td>|</td>
          <td><a href="#" class="ToolLink" onClick="CheckAll(GetObjectCollection( document, 'FORM'))">&nbsp;Seleccionar&nbsp;todos&nbsp;los&nbsp;mensajes&nbsp;</a></td>
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
          <td width="30%" class="MessageInfo"><b><a  class="ToolLink" href="Inbox.asp?uid=<%=uid%>&order=<%= Server.URLEncode("From") %>">Enviado por</a></b></td>
          <td width="48%" class="MessageInfo"><b><a  class="ToolLink" href="Inbox.asp?uid=<%=uid%>&order=<%= Server.URLEncode("Subject") %>">Asunto</a></b></td>
          <td width="20%" class="MessageInfo"><b><a  class="ToolLink" href="Inbox.asp?uid=<%=uid%>&order=<%= Server.URLEncode("Date") %>">Recibido</a></b></td>

        </tr>
<%
var Alert = new Array("","!"); 
var ChechInput = "";
var trclass = 'MessageTR';
//

  recCount = 0;
  recActual = 0;
  while ((! oRec.EOF) && (recCount < stopRec)) {
	recCount = recCount + 1;
	if (recCount >= startRec) {	 
	  recActual = recActual + 1;

    bold = '';
    if (oRec.Fields("Readed").Value == false) {
      bold = 'Style="font-weight:bolder"';
    }
	ChechInput = '<input type="checkbox"  name="INP' + recCount + '" id="INP' + recCount + '">  <INPUT TYPE=hidden NAME="ID' + recCount + '" VALUE="' + oRec.Fields("ID").Value  + '">';
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
    }
          oRec.Move(1);
 }
%>
  <input type=hidden name="Count" value="<%=recCount%>">

     </table>
    </td>
  </tr>
</table>

     <p class="txtregy">

<%

var dx1 = 0;
var dy1 = 0;
var dx2 = 0;
var dy2 = 0;

if (totalRecs > 0) {

	// Find out if there should be Backward or Forward Buttons on the table.
	if 	(startRec == 1) {
		isPrev = false;
        }		
	else {
		isPrev = true;
		PrevStart = startRec - displayRecs;
		if (PrevStart < 1 ) { PrevStart = 1; }
%>
	<center>
 	 <span class="txtboldy"><a href="Inbox.asp?uid=<%=uid%>&start=<%=PrevStart%>">&lt;&lt;&nbsp;Prev</a></span>
<%
	}
	// Display Page numbers
	if (isPrev || (! oRec.EOF)) {
		if (! isPrev) { 
%>
		 <center>
<%		}
		 x = 1;
		 y = 1;
	
		dx1 = Math.floor((startRec-1)/(displayRecs*recRange))*displayRecs*recRange+1;
		dy1 = Math.floor((startRec-1)/(displayRecs*recRange))*recRange+1;
		if ((dx1+displayRecs*recRange-1) > totalRecs) {
			dx2 = Math.floor(totalRecs/displayRecs)*displayRecs+1;
			dy2 = Math.floor(totalRecs/displayRecs)+1;
		}	
		else {
		
			dx2 = dx1+displayRecs*recRange-1;
			dy2 = dy2+recRange-1;
		}
	
		while (x <= totalRecs) {
			if ((x >= dx1) && (x <= dx2)) {
				if (startRec == x) {
%>
	<span class="txtboldy"><strong><%=y%></strong></span>
<% 			    } else {
%>
	<span class="txtboldy"><strong><a href="Inbox.asp?uid=<%=uid%>&start=<%=x%>"><%=y%></a></strong></span>
<%
			    }
				x = x + displayRecs;
				y = y + 1;
			}	
			else { 
			  if ((x >= (dx1-displayRecs*recRange)) && (x <= (dx2+displayRecs*recRange))) {
				if (x+recRange*displayRecs < totalRecs) { %>
	<span class="txtboldy"><strong><a href="Inbox.asp?uid=<%=uid%>&start=<%=x%>"><%=y%>-<%=y+recRange-1%></a></span>
<%			    }
                else {
					ny=Math.floor((totalRecs-1)/displayRecs)+1;
						if (ny == y) {
%>
	<span class="txtboldy"><a href="Inbox.asp?uid=<%=uid%>&start=<%=x%>"><%=y%></a></span>
<%					    } else {
%>
	<span class="txtboldy"><a href="Inbox.asp?uid=<%=uid%>&start=<%=x%>"><%=y%>-<%=ny%></a></span>
<%
					    }
				}
				x=x+recRange*displayRecs;
				y=y+recRange;
			  }	
			  else {
				x=x+recRange*displayRecs;
				y=y+recRange;
			  }
			}
		}
	}
	
	// Next link
	if (! oRec.EOF) {
		NextStart = startRec + displayRecs;
		isMore = true; 
%>
	<span class="txtboldy"><a href="Inbox.asp?uid=<%=uid%>&start=<%=NextStart%>">Next&nbsp;&gt;&gt;</a></span>
<%  }
	else {
	  isMore = false;
	}  
	
%>
	
    <br>
    <span class="txtboldy">
<%
   if (stopRec > recCount) { 
     stopRec = recCount; 
%>
	Records <%= startRec %> a <%= stopRec %> de <%= totalRecs%>
   </span>
<% }
 }
 else {
%>
<br>
        <p class="txtregy" align="center"><strong>La carpeta esta vacía</strong></p>
        <br>
<% 
 }
%>
<%
// Close recordset and connection
oRec.Close();
oConn.Close();
%>
<br>
<table border="0" cellspacing="0" cellpadding="0" class="ToolBar" width="100%">
  <tr>
    <td>
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td><a href="javascript:abreVentana('NewMsg',800,440,'NewMsg.asp?uid=<%=uid%>', 'no', 'no')" class="ToolLink">&nbsp;Nuevo&nbsp;mensaje&nbsp;</a></td>
          <td>|</td>
          <td><a href="javascript:onDeleteMessage()" class="ToolLink">&nbsp;Borrar&nbsp;mensaje(s)&nbsp;</a></td>
          <td>|</td>
          <td><a href="#" class="ToolLink" onClick="CheckAll(GetObjectCollection( document, 'FORM'))">&nbsp;Seleccionar&nbsp;todos&nbsp;los&nbsp;mensajes&nbsp;</a></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<input type=hidden name='flag' id='flag' value=1>  
</form>
<script language="javascript" >
<!--
   
function onDblclick( id ) {    
     event.srcElement.parentElement.parentElement.style.fontWeight  = "lighter";
     abreVentana('',800,484,'ViewMsg.asp?uid=<%=uid%>&ID=' + id + '&Flag=1&adm=1', 'no', 'no')
}  

    
function onDeleteMessage() {      
     	Messages.all("flag").value = 0;
     	Messages.action = "inbox.asp?uid=<%=uid%>";
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

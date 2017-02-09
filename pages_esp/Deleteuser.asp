<%@ Language=JScript %>
<%
  Response.Expires = -1;
%>
<!-- #include file='../js/user.inc' -->

<%
 var   displayRecs = 12;
 var   recRange = 10
 // Get table name
 var  tablename = "[Usuarios]";

    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       


  if (Session("PermissionType") == ADMINISTRATOR)
    {
    
      if (Request.QueryString("id").Count > 0) {
	    id=Request.QueryString("id");
	    Session("id") = id;
	  }  
      else {
	   id = Session("id");
	  } 
      

      //Load Default Order
      DefaultOrder = "name";
      DefaultOrderType = "ASC";

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
    strsql = "select * from " + tablename;
    dbwhere = " not ID in(" + Session("userID") + "," + ADMIN_USER + "," + GUEST_USER + ") ";
    if (dbwhere != "") strsql = strsql + " WHERE " + dbwhere;
    if (OrderBy != "") strsql = strsql + " ORDER BY " + tablename +".[" + OrderBy + "] " + Session("OrderType");

    var  filePath = Application("dataPath");   
    var  oConn    = Server.CreateObject("ADODB.Connection");
    var  oRec     = Server.CreateObject("ADODB.Recordset");
    oConn.Open(filePath);
    oRec.Open(strsql, oConn, 1, 2);
    totalRecs = oRec.RecordCount;
%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<script src="../js/CheckBoxes.js" language="JavaScript"></script>
</HEAD>
<body bgcolor="#FFFFFF" text="#000000">
<form name=Delnews action="ConfirmDeleteUser.asp?uid=<%=uid%>" method="post">
<table border="0" cellspacing="0" cellpadding="0" class="ToolBar" width="100%">
  <tr>
    <td>
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td nowrap><a href="AgregaUser.asp?uid=<%=uid%>" class="ToolLink">&nbsp;Crear&nbsp;</a></td>
          <td>|</td>
          <td nowrap><a href="javascript:Delnews.submit()" class="ToolLink">&nbsp;Eliminar&nbsp;</a></td>
          <td>|</td>
          <td><a href="#" class="ToolLink" onClick="CheckAll(GetObjectCollection( document, 'FORM'))">&nbsp;Seleccionar&nbsp;Todo&nbsp;</a></td>
        </tr>
      </table>
    </td>
  </tr>
</table>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar" align="center"><b>Listado de usuarios</b></td>
  </tr>
  <tr> 
    <td> 
      <table width="100%" border="0" cellspacing="1" cellpadding="0">
        <tr> 
          <td width="4%" class="ToolBar" align="center"></td>
          <td width="20%" class="ToolBar" align="center"><b><a href="Deleteuser.asp?uid=<%=uid%>&order=<%= Server.URLEncode("name") %>">Identificador</a></b></td>
          <td width="30%" class="ToolBar" align="center"><b><a href="Deleteuser.asp?uid=<%=uid%>&order=<%= Server.URLEncode("fullname") %>">Nombre y apellidos</a></b></td>
          <td width="30%" class="ToolBar" align="center"><b><a href="Deleteuser.asp?uid=<%=uid%>&order=<%= Server.URLEncode("email") %>">Correo eléctronico</a></b></td>
          <td width="8%" class="ToolBar" align="center"><b><a href="Deleteuser.asp?uid=<%=uid%>&order=<%= Server.URLEncode("fechaIngreso") %>">Alta</a></b></td>
          <td width="8%" class="ToolBar" align="center"><b><a href="Deleteuser.asp?uid=<%=uid%>&order=<%= Server.URLEncode("lastLogin") %>">Último acceso</a></b></td>
        </tr>
      
<%    
  var clase = "MessageTR1";
      
  recCount = 0;
  recActual = 0;
  while ((! oRec.EOF) && (recCount < stopRec)) {
	recCount = recCount + 1;
	if (recCount >= startRec) {	 
	  recActual = recActual + 1;
      if (clase == "MessageTR1") clase = "MessageTR"; else clase = "MessageTR1";

%>
        <tr> 
          <td width="4%" class="<%=clase%>">
             <input type="checkbox" id=Check<%=recCount%> name=check<%=recCount%> >
             <input type="hidden" id=ID<%=recCount%> name="ID<%=recCount%>" value="<%=oRec.Fields.Item("ID").value%>">
             <input type="hidden" id=Name<%=recCount%> name="Name<%=recCount%>"  value="<%=oRec.Fields.Item("Name").value%>" >
          </td>
          <td width="20%" class="<%=clase%>"><a href="ModifyUser.asp?uid=<%=uid%>&user=<%=oRec.Fields.Item("ID").value%>"><%=oRec.Fields.Item("name").value%></a></td>
          <td width="30%" class="<%=clase%>"><%=oRec.Fields.Item("FullName").value%></td>
          <td width="30%" class="<%=clase%>"><a href="mailto:<%=oRec.Fields.Item("email").value%>" ><%=oRec.Fields.Item("email").value%></a></td>

          <td width="8%" class="<%=clase%>"><%=oRec.Fields.Item("fechaIngreso").value%></td>
          <td width="8%" class="<%=clase%>"><%=oRec.Fields.Item("lastLogin").value%></td>
          
        </tr> 

<%  }
    oRec.Move(1);
  }  
%>        

<input type="hidden"  name="Count" value="<%=recCount%>" >

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
 	 <span class="txtboldy"><a href="Deleteuser.asp?uid=<%=uid%>&start=<%=PrevStart%>">&lt;&lt;&nbsp;Prev</a></span>
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
	<span class="txtboldy"><strong><a href="Deleteuser.asp?uid=<%=uid%>&start=<%=x%>"><%=y%></A></strong></span>
<%
			    }
				x = x + displayRecs;
				y = y + 1;
			}	
			else { 
			  if ((x >= (dx1-displayRecs*recRange)) && (x <= (dx2+displayRecs*recRange))) {
				if (x+recRange*displayRecs < totalRecs) { %>
	<span class="txtboldy"><strong><a href="Deleteuser.asp?uid=<%=uid%>&start=<%=x%>"><%=y%>-<%=y+recRange-1%></a></span>
<%			    }
                else {
					ny=Math.floor((totalRecs-1)/displayRecs)+1;
						if (ny == y) {
%>
	<span class="txtboldy"><a href="Deleteuser.asp?uid=<%=uid%>&start=<%=x%>"><%=y%></a></span>
<%					    } else {
%>
	<span class="txtboldy"><a href="Deleteuser.asp?uid=<%=uid%>&start=<%=x%>"><%=y%>-<%=ny%></a></span>
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
	<span class="txtboldy"><a href="Deleteuser.asp?uid=<%=uid%>&start=<%=NextStart%>">Next&nbsp;&gt;&gt;</a></span>
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
        <p class="txtregy"><strong>Ningún usuario fue encontrado!</strong></p>
        <br>
<% 
 }
%>
<%
// Close recordset and connection
oRec.Close();
oConn.Close();
%>
</center>
<br>
<table border="0" cellspacing="0" cellpadding="0" class="ToolBar" width="100%">
  <tr>
    <td>
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td nowrap><a href="AgregaUser.asp?uid=<%=uid%>" class="ToolLink">&nbsp;Crear&nbsp;</a></td>
          <td>|</td>
          <td nowrap><a href="javascript:Delnews.submit()" class="ToolLink">&nbsp;Eliminar&nbsp;</a></td>
          <td>|</td>
          <td><a href="#" class="ToolLink" onClick="CheckAll(GetObjectCollection( document, 'FORM'))">&nbsp;Seleccionar&nbsp;Todo&nbsp;</a></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
 </form>
<%
    }
  else  
    {
      Response.Redirect("errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);         
    }
%>

</BODY>
</HTML>
<%
  }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
  //Response.Write(Request.QueryString.Item("uid"));
%>
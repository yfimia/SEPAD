<%

function download_Modules() {


%>
<table border="0" width="100%" cellpadding="3" cellspacing="0" class="ToolBar">
    <tr> 
      <td><b>Modalidades Académicas:</b></td>
	</tr>
</table>

<%

  var filePath = Application('dataPath');
  var oConn;
  var oRec;

  oConn = MakeConnection( oConn, filePath );
  Sql = "select * FROM Modulo Order By Name";

  oRec = Query( Sql, oRec, oConn  );		


  while (!oRec.EOF)
    {
     ID               = oRec.Fields.Item("ID").Value;
     Name	      = oRec.Fields.Item("Name").Value;
     state            = oRec.Fields.Item("State").Value;
     cordinador       = oRec.Fields.Item("cordinador").Value;
     grupo            = oRec.Fields.Item("grupo").Value;
     claustro         = oRec.Fields.Item("claustro").Value;
     //Chequeos de los derechos del usuario actual respecto a este curso
     var isowner = ((Session("UserID") == cordinador) || (Session("permissionType") == ADMINISTRATOR));
     var isprof = ( isowner || (isUserInGroupByID(Session("UserID"), claustro)));

if (((state == MOD_ACA_DEVELOP) && (isprof))  || (Session("PermissionType") == ADMINISTRATOR) || (Session("userID") == cordinador) || ((state != MOD_ACA_NOVISIBLE) && (state != MOD_ACA_DEVELOP)) || (state == MOD_ACA_FREE) ) {
%>

<ul>
  <li class="clsDocument" > 
    <a   href="javascript:goGenMod(<%=ID%>, '<%=Name%>', <%=state%>, <%=cordinador%>, <%=grupo%>, <%=claustro%>)" >
     <IMG SRC="../images/<%=Session("skin")%>/TreeViewImg.gif"   class="IconClass" /> 
    <%=Name%></a> 
    
  </li>             
</ul>

<%
   }
     oRec.Move(1);  
     
    } 
   DestroyAdoObjects( oConn, oRec );

 }        

%>
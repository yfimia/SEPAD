<%@ Language=JScript %>
<%
  Response.Expires = -1;
%>
<!-- #include file='../js/user.inc' -->
<!-- #include file='../js/adolibrary.inc' -->
<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       
      var scpt = '';


  if (Session("PermissionType") == ADMINISTRATOR)
    {
    first = "-1";
    where = "";
    /*    
    if (Request.QueryString.Count >= 2) {
      first = Request.QueryString.Item("first") + "";
      
      if (first != "-1") {
        where = " and (Name > '" +  first + "')";
      } 

    } 
    */      
    
%>

<HTML>
<HEAD>
<title>Seleccionar grupos</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<script src="../js/CheckBoxes.js" language="JavaScript"></script>

</HEAD>
<body bgcolor="#FFFFFF" text="#000000">
<%    
      var  filePath = Application("dataPath");   
      var  oConn    = Server.CreateObject("ADODB.Connection");
      var  oComm    = Server.CreateObject("ADODB.Command");
      var  oRec     = Server.CreateObject("ADODB.Recordset");
      oRec.CursorLocation = 3;
      oRec.CursorType     = 3;
      oConn.Open(filePath);
      oComm.ActiveConnection = oConn;
      oComm.CommandText = "SELECT * FROM grupos Order By Name";
      oRec = oComm.Execute();    
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar" align="center"><b>Listado de grupos</b></td>
  </tr>
  <tr> 
    <td> 
      <table width="100%" border="0" cellspacing="1" cellpadding="0">
        <tr> 
          <td width="5%" class="ToolBar" align="center"></td>
          <td width="40%" class="ToolBar" align="center"><b>Nombre</b></td>
          <td width="50%" class="ToolBar" align="center"><b>Descripción</b></td>
        </tr>
      
<%    
      var clase = "MessageTR1";
      
      var i = 0;
      var last = "-1"
      while (oRec.EOF == false)
        { 
          i = i + 1;
          if (clase == "MessageTR1") clase = "MessageTR"; else clase = "MessageTR1";
          last = oRec.Fields.Item("Name").value;          
          var chain = new String(oRec.Fields.Item("description").value);
          var str = new String(DoubleQuotationToSimpleQuotation(chain.toString()));
          scpt = scpt +  ' groups[groups.length] = new Group(' + oRec.Fields.Item("id").value + ',"' + oRec.Fields.Item("name").value + '","' + str.toString() + '");\n';
%>
        <tr> 
          <td width="5%" class="<%=clase%>"><input type="checkbox" id=Check<%=i%> name=check<%=i%> ></td>
          <td width="40%" class="<%=clase%>"><%=oRec.Fields.Item("name").value%></td>
          <td width="50%" class="<%=clase%>"><%=oRec.Fields.Item("description").value%></td>
        </tr> 

<%
          oRec.Move(1);
        }  
      if (i > 0) 
        {
          oRec.MoveFirst();
//          Response.Write('<tr><td><INPUT id=Buton name=Buton type=Submit value=Borrar></td><td><INPUT id=Buton name=Buton type=Submit value=Regresar onclick="history.back(-1)"></td></tr>');
        }
      else
        {
%>        
          <tr><td align="center" width="100%" colspan="4" class="MessageTR">No quedan más elementos.</td></tr>        
<%          
        }  
%>        
<table border="0" cellspacing="0" cellpadding="0" class="ToolBar" width="100%">
  <tr>
    <td>
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td nowrap><!--a onclick="SelectUser()" class="ToolLink">&nbsp;Aceptar&nbsp;</a-->
            <input name=Mname Id=Mname type=button value="Seleccionar" onclick="SelectGroup()">
          </td>
<!--          <td>|</td>
          <td><a target="_self" href="findUser.asp?uid=<%=uid%>&first=<%=last%>" class="ToolLink">&nbsp;Próximos&nbsp;<%=SHOW_CANT%>&nbsp;</a></td>
-->          
        </tr>
      </table>
    </td>
  </tr>
</table>

<script language="JavaScript">
  function Group(id, name, desc) {
    this.id = id;
    this.name = name;
    this.desc = desc;
  }

  groups = new Array();
  resgroups = new Array();
  <%=scpt%>

  function SelectGroup() {
        
	var CheckList = GetObjectCollection( document.body, "INPUT" );
	var i=0;
  	for ( i=0; i < CheckList.length; i++ ) {
	  	if ( (CheckList[i].type == "checkbox") && (CheckList[i].checked == true) ) {	  	
	  		resgroups[resgroups.length] = groups[i];
                }
         }       
         returnValue = resgroups; 
         close();       
  }         
</script>

<%      
   DestroyAdoObjects( oConn, oRec );       
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
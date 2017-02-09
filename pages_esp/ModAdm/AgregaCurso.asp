<%@ Language=JScript %>
<%Response.Expires = -1;%>
<!-- #include file='../../js/adoLibrary.inc' -->
<!-- #include file='../../js/user.inc' -->
<!-- #include file='inc\AgregaCurso.inc' -->

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

%>

<%
  if ((Session("PermissionType") == ADMINISTRATOR) || (Session("userID") == Session("admcordinador")))
    {
      if ((Request.Form("Mname").Count == 0) || (Request.Form("Mname") == ""))
        {
%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/EvalResult.css" type="text/css">

<script src="../../js/windows.js" language="JavaScript"></script>
</head>
<script language=jscript>
  function User(id, name, fullName, email) {
    this.id = id;
    this.name = name;
    this.fullName = fullName;
    this.email = email;
  }
  
  
  function findUser() {
    var user = new Array();
    user = showModalDialog("findTutor.asp?uid=<%=uid%>", "", "");
    if (user != null) {
      newgroup.Mcord.value = user[0].fullName + ' (' + user[0].name + ')';
      newgroup.Mcordid.value = user[0].id;
    }  
  }
  
</script>
<body bgcolor="#FFFFFF" text="#000000">
  <form name=newgroup id=newgroup action="agregaCurso.asp?uid=<%=uid%>" method=post>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td class="HeaderTable"  align=center><b><%=AModulo%></b></td>
        </tr>
      </table>
    </td>
  </tr>
</table>

    <table border="0" cellspacing="1" cellpadding="1" align="center">
      <tr>
        <td  align=center Class="MessageTR1">
          <%=Nombre%>
        </td>
        <td  Class="MessageTR1"align=center >
          <input name=Mname Id=Mname type=text maxlength=250 size=65 >
        </td>
      </tr>
      <tr>
        <td align=center   Class="MessageTR">
          <%=Estado%>
        </td>
        <td align=center   Class="MessageTR">
          <%=CUR_SELECT_STATE%>
        </td>
      </tr>
      <tr>
      <tr>
        <td align=center   Class="MessageTR">
          <%=ProfPrincipal%>
        </td>
        <td align=center   Class="MessageTR">
          <input disabled name=Mcord Id=Mcord type=text maxlength=250 size=45 ><input name=Mname Id=Mname type=button value=<%=Seleccionar%> onclick="findUser()">
          <input name=Mcordid Id=Mcordid type=hidden maxlength=250 size=45 >
          
        </td>
      </tr>



      
      <tr>
        <td colspan=2 align=center Class="Toolbar">
          <input type=submit value=<%=Adicionar%>>
        </td>      
      </tr>
    </table>
  </form>
<%
        }
      else
        {
           var filePath = Application('dataPath');
           var oConn;
           var oRec;
 
           oConn = MakeConnection( oConn, filePath );
           Sql = "select * from Cursos where (modulo = " + Session("admmodulo") + ") and (Name = '" + Request.Form("Mname") + "')";

           oRec = Query( Sql, oRec, oConn  );	
          

          if (oRec.EOF == false)
            { 
              oConn.Close();   
              
              Response.Redirect("../errorpage.asp?tipo=Error&short=" + CUR_NAME_EXIST_SHORT  + "&desc=" + CUR_NAME_EXIST_TEXT); 
            }  
          else
            {
              oConn.Close();   
              oConn.Errors.Clear();
              oConn.Open( filePath );
              oRec.Open("select * from Cursos",oConn,3,3);
              oRec.AddNew();
              
              oRec.Fields.Item("modulo").Value = Session("admmodulo");
              if ((Request.Form("MName").Count != 0) && (Request.Form("Mname") + "" != ""))
                { 
                  oRec.Fields.Item("Name").Value = Request.Form("Mname") + "";
                }
              else Response.Redirect("../errorpage.asp?tipo=Error&short=Nombre no válido&desc=El nombre del módulo introducido no es válido o se ha dejado en blanco.");

              oRec.Fields.Item("state").Value = Request.Form("Mstate") + "";

              if ((Request.Form("Mcordid").Count != 0) && (Request.Form("Mcordid") + "" != ""))
                { 
                  oRec.Fields.Item("owner").Value = Request.Form("Mcordid") + "";
                }
              else Response.Redirect("../errorpage.asp?tipo=Error&short=Profesor principal&desc=El profesor principal no ha sido seleccionado apropiadamente. En la página de crear módulo para seleccionar el profesor principal, precione el botón Seleccionar y escójalo de la lista.");


              oRec.Update();      
              oConn.Close();   
              
              Response.Redirect("agregaCurso.asp?uid=" + uid);      
              
            }  
        }
    }    
  else
    {  
     Response.Redirect("../errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);                   
    }
%>

</BODY>
</HTML>
<%
  }
 else
   Response.Redirect("../errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
   
  //Response.Write(Request.QueryString.Item("uid"));
%>



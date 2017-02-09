<%@ Language=JScript %>
<%Response.Expires = -1;%>
<!-- #include file='../js/user.inc' -->

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

%>

<%
if (Session("PermissionType") == ADMINISTRATOR)
    {
      if ((Request.Form("Gname").Count == 0) || (Request.Form("Gname") == ""))
        {
%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/EvalResult.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000">
  <form name=newgroup id=newgroup action="agregagroup.asp?uid=<%=uid%>" method=post>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td class="HeaderTable"  align=center><b>Adicionar grupo</b></td>
        </tr>
      </table>
    </td>
  </tr>
</table>

    <table border="0" cellspacing="1" cellpadding="1" align="center">
      <tr>
        <td  align=center Class="MessageTR1">
          Nombre del nuevo grupo:
        </td>
        <td  Class="MessageTR1"align=center >
          <input name=Gname Id=Gname type=text maxlength=250 size=50 >
        </td>
      </tr>
      <tr>
        <td align=center   Class="MessageTR">
          Descripción del grupo:
        </td>
        <td align=center   Class="MessageTR">
          <textarea name=GDescript Id=GDescript" rows="8" cols="50" maxlength=250></textarea>
        </td>
      </tr>
      <tr>
        <td colspan=2 align=center Class="Toolbar">
          <input type=submit value="Adicionar">
        </td>      
      </tr>
    </table>
  </form>
<%
        }
      else
        {
          var  filePath = Application("dataPath");   
          var  oConn    = Server.CreateObject("ADODB.Connection");
          var  oComm    = Server.CreateObject("ADODB.Command");
          var  oRec     = Server.CreateObject("ADODB.Recordset");
          oConn.Open( filePath );
          oRec.Open("select * from grupos where (Name = '" + Request.Form("Gname") + "')",oConn,3,3);
          //Response.Write("select * from grupos where (Name = '" + Request.Form("Gname") + "')");
          if (oRec.EOF == false)
            { 
              oConn.Close();   
              
              Response.Redirect("errorpage.asp?tipo=Error&short=" + GRUOP_NAME_EXIST_SHORT  + "&desc=" + GRUOP_NAME_EXIST_TEXT); 
              
              Response.Write ("<center><font color = red> El grupo " + Request.Form("Gname") + " ya existe. Seleccione otro nombre para su grupo</font><br>" );
              Response.Write ('<script languaje=javascript>function pclick(){window.location="usermanager.asp";}</script>');      
              Response.Write ("<input type=Button Value=Regresar onclick='return pclick()' id=Button1 name=Button1></center>" );      ;
            }  
          else
            {
              oConn.Close();   
              oConn.Errors.Clear();
              oConn.Open( filePath );
              oRec.Open("select * from grupos",oConn,3,3);
              oRec.AddNew();
              oRec.Fields.Item("Name").Value = Request.Form("Gname") + "";
              if ((Request.Form("GDescript").Count != 0) && (Request.Form("GDescript") + "" != ""))
                { 
                  oRec.Fields.Item("Description").Value = Request.Form("GDescript") + ""; 
                }
              else oRec.Fields.Item("Description").Value = " ";   
              oRec.Update();      
              oConn.Close();   
              
              Response.Redirect("agregagroup.asp?uid=" + uid);      
              
              Response.Write ("<center><font color = red>Grupo creado satisfactoriamente</font><br>" );
              Response.Write ('<script languaje=javascript>function pclick(){window.location="usermanager.asp";}</script>');      
              Response.Write ("<input type=Button Value=Regresar onclick='return pclick()' id=Button1 name=Button1></center>" );      ;
            }  
        }
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



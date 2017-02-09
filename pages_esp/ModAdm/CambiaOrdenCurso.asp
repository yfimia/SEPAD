<%@ Language=JScript %>
<%
  Response.Expires = -1;  
%>
<!-- #include file='../../js/user.inc' -->
<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       
    first = "-1";
    where = "";
    if (Request.QueryString.Item("first").Count > 0) {
      first = Request.QueryString.Item("first") + "";
      
      if (first != 0) {
        where = " and (Cursos.modOrder > '" +  first + "')";
      }
      
    if ((Request.QueryString.Item("check").Count > 0) && (Request.QueryString.Item("check") + "" == "TRUE"))  
    {
      var  filePath = Application("dataPath");
      var  oConn    = Server.CreateObject("ADODB.Connection");
      var  oComm    = Server.CreateObject("ADODB.Command");
      var  oRec     = Server.CreateObject("ADODB.Recordset");
      oRec.CursorLocation = 3;
      oRec.CursorType     = 3;
      oConn.Open(filePath);
      oComm.ActiveConnection = oConn;
      
      oRec.Open("SELECT Top " + SHOW_CANT + "  Cursos.modOrder, Cursos.prologue, Cursos.ID, Cursos.state, Cursos.owner, Cursos.Name, Usuarios.FullName, Usuarios.Email, Usuarios.ID AS UID FROM Usuarios INNER JOIN Cursos ON Usuarios.ID = Cursos.Owner WHERE (Cursos.modulo = " + Session("admmodulo") + ")"  +  where + " Order By Cursos.modOrder", oConn, 3, 3);
	  conta=oRec.RecordCount;
      oRec.MoveFirst();
      for (i=1; i<=conta; i++)
      {       
        oRec.Fields.Item("modOrder").value = Request.Form("hidden" + i);
        oRec.Move(1);
      }
      
    }
    else var counter=0;   
    } 
    

%>

<%  
 /*   coursename = Request.QueryString.Item('coursename') + '';
  course     = Request.QueryString.Item('course') + '';

  Session('course')  = course;
  Session('courseName') = coursename;
*/

  if ((Session("PermissionType") == ADMINISTRATOR) || (Session("userID") == Session("admcordinador")))
    {
%>

<html>
<head>
<title>Ver/Eliminar alumnado</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/EvalResult.css" type="text/css">
<script src="../../js/CheckBoxes.js" language="JavaScript"></script>
<script src="../../js/windows.js" language="JavaScript"></script>

<script language="JScript">
<!--
function close_window() {
    window.close();
}
//-->

function change_value(ident, cont, valor)
{
  i=1;  
  
  
 // alert(document.all.item(26+(4+cont)*i).name);
  while (document.all.item("hidden" + i).value != ident)
  {
    i++;
  }
  
  indexa = 26+(4+cont)*i;
  
  document.all.item("select" + i).selectedIndex = valor - 1;
  document.all.item("hidden" + i).value = valor;
//  document.all.item("select" + i).value = valor;
}



function select_onchange(ident, cont)
{  
  change_value(document.all.item("select" + ident).selectedIndex + 1, cont, document.all.item("hidden" + ident).value);  
  document.all.item("hidden" + ident).value = document.all.item("select" + ident).selectedIndex + 1;
//  document.all.item("select" + ident).value = document.all.item("select" + ident).selectedIndex + 1;
}


</script>
</head>

<body bgcolor="#FFFFFF" text="#000000">

<form id="orden" name="orden" action="CambiaOrdenCurso.asp?uid=<%=uid%>&first=-1&check=TRUE&cont<%=counter%>" method="post">

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td class="HeaderTable"  align=center><b>Módulos</b></td>
        </tr>
      </table>
    </td>
  </tr>
</table>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
<%    
      var  filePath = Application("dataPath");
      var  oConn    = Server.CreateObject("ADODB.Connection");
      var  oComm    = Server.CreateObject("ADODB.Command");
      var  oRec     = Server.CreateObject("ADODB.Recordset");
      oRec.CursorLocation = 3;
      oRec.CursorType     = 3;
      oConn.Open(filePath);
      oComm.ActiveConnection = oConn;

      
%>
  <tr> 
    <td> 
      <table width="100%" border="0" cellspacing="1" cellpadding="0">
        <tr> 
          <td width="1%" class="ToolBar" align="center">&nbsp; </td>
          <td width="99%" class="ToolBar" align="center"><b>Nombre</b></td>
        </tr>
        <tr> 
<%    
      var clase = "MessageTR1";
      

      oRec.Open("SELECT Top " + SHOW_CANT + "  Cursos.modOrder, Cursos.prologue, Cursos.ID, Cursos.state, Cursos.owner, Cursos.Name, Usuarios.FullName, Usuarios.Email, Usuarios.ID AS UID FROM Usuarios INNER JOIN Cursos ON Usuarios.ID = Cursos.Owner WHERE (Cursos.modulo = " + Session("admmodulo") + ")"  +  where + " Order By Cursos.modOrder Asc", oConn, 3, 3);
      var i = 0;
      last = 0;
      
      counter=oRec.RecordCount;
      
      if (counter > 0) oRec.MoveFirst();
      while (oRec.EOF == false)
        {                    
          i++;
          last = i;
          if (oRec.Fields.Item("modOrder").value != i)
          {
      //      Response.Write('Si');
            oRec.Fields.Item("modOrder").value = i;
            oRec.UpDate();
          }
          if (clase == "MessageTR1") clase = "MessageTR"; else clase = "MessageTR1";
          
          
%>            
          <td width="5%" class="<%=clase%>">
            <input type=hidden id="hidden<%=i%>" name="hidden<%=i%>" value=<%=i%>>
            <select id="select<%=i%>" name="select<%=i%>" onchange="select_onchange(<%=i%>,<%=counter%>)">
              <%
                for (j=1;j<=counter;j++)
                {
                  if (j==i)
                  {
                    %><option value="<%=j%>" selected><%=j%></option><%
                  }
                  else
                  {
                    %><option value="<%=j%>"><%=j%></option><%
                  }
                }
              %>
            </select>
          </td>
          <td width="20%" class="<%=clase%>"><%=oRec.Fields.Item("name").value%></td>
                                  
<%              
          oRec.Move(1);
%>
        </tr>
<%
          
        }  
      if (i > 0) 
        {
          oRec.MoveFirst();
//          Response.Write('<tr><td><INPUT id=Buton name=Buton type=Submit value=Borrar></td><td><INPUT id=Buton name=Buton type=Submit value=Regresar onclick="history.back(-1)"></td></tr>');
        }
      else
        {
%>        
          <tr><td align="center" width="100%" colspan="7" class="MessageTR">No hay cursos por ordenar.</td></tr>        
<%          
        }  
      
      oRec.Close();
      oConn.Close();
%>
      </table>
    </td>
  </tr>
  <tr>
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td><a href="CambiaOrdenCurso.asp?uid=<%=uid%>&first=<%=last%>" class="ToolLink">&nbsp;Próximos&nbsp;<%=SHOW_CANT%>&nbsp;</a></td>
          <td class="ToolLink">|</td>

          <td><a href="javascript:orden.submit()" class="ToolLink">&nbsp;Aceptar&nbsp;</a></td>

        </tr>
      </table>
    </td>
  </tr>
<%
    }
  else  
    {
     Response.Redirect("../errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);         
    } 
%>
  
</table>
</form>

</body>
</html>
<%
  }
 else
   Response.Redirect("../errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>
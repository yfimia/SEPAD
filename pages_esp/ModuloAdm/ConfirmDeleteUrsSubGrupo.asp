<%@ Language=JScript %>
<%
  Response.Expires = -1;
  Response.Buffer = true;  
%>

<!-- #include file='../../js/user.inc' -->
<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

%>


<%
  if ((Session("PermissionType") == ADMINISTRATOR) || (Session("userID") == Session("admcursocordinador"))  || (Session("userID") == Session("admcursoOwner")))
    {
%>
<%
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
      var fname = new Array();
      var IDU = new Array();
      var nick = new Array();
      while (oRec.EOF == false)
        {
          i = i + 1;
          if (Request.Form("Check" + i) == "on") 
            {
              j = j + 1;
              fname[j] = oRec.Fields.Item("FullName").Value + "";
              IDU[j] = oRec.Fields.Item("ID").Value;
              nick[j] = oRec.Fields.Item("Name").Value;
            }
          oRec.Move(1);
        }
      oConn.Close();   
      if (j >= 1)
        {
          for (i=1; i<=j; i++) 
            {
              oConn.Open(filePath);
              oComm.ActiveConnection = oConn; 
              oComm.CommandText = "delete from UsuariosSubGrupo where (Usuario = " + IDU[i] + ") and (Subgrupo = "  + Request.QueryString.Item("sgrupo") + ")";
             // Response.Write(oComm.CommandText);
              oRec = oComm.Execute();
              oConn.Close();          
            }
        } 
  Response.Redirect('modifySubGrupo.asp?uid=' + uid + "&sgrupo=" + Request.QueryString.Item("sgrupo")); 
%>
<SCRIPT LANGUAGE=javascript>
<!--
function muOnclick()
 {
   location = 'verEliminaralumnado.asp?uid=<%=uid%>&course=<%=Session("course")%>&coursename=<%=Session("courseName")%>' ;
 }
 
//-->
</SCRIPT>

 <center><b>Los alumnos fueron dado de baja satisfactoriamente.</b></center>
 <center>    <INPUT align=center id=Buton name=Buton type=Button value="Regresar"  onclick=muOnclick()></center>
<%        
      Session("UserList") = null;
      Session("Conection") = null;
      Session("Command") = null;
    }
  else
    {  
      Response.Redirect("../../errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);     
    }

%>
    
</body>
</html>
<%
  }
 else
   Response.Redirect("../../errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
  //Response.Write(Request.QueryString.Item("uid"));
%>
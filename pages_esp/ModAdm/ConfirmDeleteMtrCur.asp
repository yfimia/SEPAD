<%@ Language=JScript %>
<%
  Response.Expires  = -1;
  Response.Buffer = true;  
%>
<SCRIPT language=vbscript RUNAT=Server>
 function getAtualTime
  getAtualTime = Now()
 End function 

</SCRIPT>

<!-- #include file='../../js/user.inc' -->
<!-- #include file="../../js/adolibrary.inc" -->
<!-- #include file="../../js/library.inc" -->
<!-- #include file="inc/ConfirmDeleteMtrCur.inc" -->

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

%>


<%
  if ((Session("PermissionType") == ADMINISTRATOR) || (Session("userID") == Session("admcordinador")))
    {


      var  filePath = Application("dataPath");   
      var  oRec  = Server.CreateObject("ADODB.Recordset");
      var  oConn = Server.CreateObject("ADODB.Connection"); 
      var  oComm    = Server.CreateObject("ADODB.Command");
      oRec  = Session("UserList");
      oConn = Session("Conection");
      oComm = Session("Command");
      var i = 0;
      var j = 0;
      var IDMA = new Array();
      userid   = new Array();

      while (oRec.EOF == false)
        {
          i = i + 1;
          if (Request.Form("Check" + i) == "on") 
            {
              j = j + 1;
              IDMA[j] = oRec.Fields.Item("ID").Value;
              userid[j] = oRec.Fields.Item("User").Value;
              //Response.Write(oRec.Fields.Item("ID").Value + "<BR>" + oRec.Fields.Item("User").Value + "<BR>" );
            }
          oRec.Move(1);
        }
      oConn.Close();   
      
      
      if (j >= 1)
        {
          //borro las matriculas 
          oConn.Open(filePath);

          for (i=1; i<=j; i++) 
            {
              oComm.ActiveConnection = oConn; 
              oComm.CommandText = "delete  from TMP_Matriculas_Cursos where (ID = " + IDMA[i] + ")";
              oRec = oComm.Execute();
              SendMsg( 'SEPAD', userid[i], MatriculaDenegada + Session("admmoduloname"), MatModalidadAcadem + Session("admmoduloname") + HaSidoDenegada + '\n\n' + Mensaje + Session("admmoduloname") + ".");                    
              
            }  
          oConn.Close();          
            
        }     
  Response.Redirect('MtrCurManager.asp?uid=' + uid );                  
%>            
<SCRIPT LANGUAGE=javascript>
<!--
function muOnclick()
 {
   location = 'MtrCurManager.asp?uid=<%=uid%>' ;
 }
 
//-->
</SCRIPT>

 <center><b><%=SolicitudesBaja%></b></center>
 <center>    <INPUT align=center id=Buton name=Buton type=Button value=<%=Regresar%>  onclick=muOnclick()></center>
<%          

      Session("UserList") = null;
      Session("Conection") = null; 
      Session("Command") = null;
    }
  else
    {  
    Response.Redirect("../errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);     
    }
 }
 else
   Response.Redirect("../errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
  // Response.Write(Request.QueryString.Item("uid"));
%>

    
</body>
</html>

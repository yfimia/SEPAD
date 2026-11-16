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

%>


<%
  if ((Session("PermissionType") == ADMINISTRATOR) || (Session("PermissionType") == PUBLICATOR))
    {

%>

<SCRIPT language=vbscript RUNAT=Server>
 function getAtualTime
  getAtualTime = Now()
 End function 

</SCRIPT>

<%
    var dataPath = Application('dataPath');
    
    var oConn;
    var oRec;  
    oConn = MakeConnection( oConn, dataPath );
    Sql ="select * from Noticias";

    oRec = Query( Sql, oRec, oConn  );		

  
          oRec.AddNew();
          oRec.Fields.Item("Titulo").value  = Request.Form('Titulo');
          oRec.Fields.Item("Cuerpo").value  = Request.Form('Cuerpo');
          oRec.Fields.Item("Imagen").value  = Request.Form('Imagen');
          oRec.Fields.Item("Url").value  = Request.Form('Url');
          oRec.Fields.Item("Fecha_Publicada").value      = getAtualTime();
          oRec.Fields.Item("Publisher").value = Session("userID");
          oRec.Update();      
           
    DestroyAdoObjects( oConn, oRec );
    Response.Redirect("news.asp?uid=" + uid);
 %>
<SCRIPT LANGUAGE=javascript>
<!--
function muOnclick()
 {
   location = 'news.asp?uid=<%=uid%>';
 }
 
//-->
</SCRIPT>

 <center><b>Las noticia fue publicada satisfactoriamente.</b></center>
 <center>    <INPUT align=center id=Buton name=Buton type=Button value="Regresar"  onclick=muOnclick()></center>
<%        
    }
  else
    {  
      Response.Redirect("errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);     
    }
%>
</body>
</html>
<%
  }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
  //Response.Write(Request.QueryString.Item("uid"));
%>



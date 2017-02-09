<%@language=Jscript%>
<Response.Expires = -1>

<head>

   <title> Modificar Anucio </title>
  
  
     <script language=Jscript >
   
         function make() {
           
        
         }

     </script>   
 
  

</head>



 
<html>
  
 <table width="100%" border=0 cellspacing=0 cellpadding=0>
  <tr> 
    <td class="ToolBar">
      <table border=0 cellspacing=0 cellpadding=0 class="ToolBar">
        <tr> 
          <td class="HeaderTable" align=center><b>Modificar Anuncios</b></td>
        </tr>
      </table>
    </td>
  </tr>
</table>

 <%
   
 
   
          var  filePath = Application("dataPath");   
          var  oConn    = Server.CreateObject("ADODB.Connection");
          var  oComm    = Server.CreateObject("ADODB.Command");
          var  oRec     = Server.CreateObject("ADODB.Recordset");
          oConn.Open( filePath );
         
          oRec.Open("select * from Anuncios",oConn,3,3);
       
       Response.write("<h4> Titulos : "); 
       Response.write("<form name=ModAnuncio Id=ModAnuncio Method=Post Action=Mod_Anun.asp>");
       Response.write("<select name=Anuncio id=Anuncio  onchange=make()>"); 
    
      while  (!oRec.EOF)  {
      
       var Ident =oRec.Fields.Item("Titulo").Value + "";  
       Response.write("<option value=" + Ident + ">"  + Ident);
       
     oRec.Move(1);       
     }
    Response.write("</Select>");
    Response.write("<Input Type=Submit value=Modicar></form>");  
    oConn.close();
 %>
 
 
 
 
 
</html>
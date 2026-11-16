<%@ Language=JavaScript%>
<%
  Response.Expires = -1;
%>
<!-- #include file="../js/adoLibrary.inc" -->
<!-- #include file="../js/gettime.inc" -->
<!-- #include file='../js/user.inc' -->

<%  
  var MAXNEW = Application("NewNews"); //minimo de dias para considerar una noticia como nueva.

  var filePath = Application('dataPath');
  
  var oConn;
  var oRec;

%>

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       
    first = "-1";
    where = "";
    if (Request.QueryString.Count >= 2) {
      first = Request.QueryString.Item("first") + "";
      
      if (first != "-1") {
        where = " and where (Fecha_Publicada < " + Application("dchar") +  first + Application("dchar") + ")";
      } 
    } 
      

%>

<html>
<head>
<title>Noticias</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/EvalResult.css" type="text/css">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<script src="../js/windows.js" language="JavaScript"></script>
</head>

<body bgcolor="#FFFFFF" text="#000000" style="overflow-y:scroll">
<table width="100%" border="0" cellspacing="0" cellpadding="2">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td><img src="../images/<%=Session("skin")%>/NoticiasIMG.gif" width="80" height="54"></td>
          <td class="HeaderTable">Noticias</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td>
      <hr noshade>
    </td>
  </tr>  
</table>

<%
    oConn = MakeConnection( oConn, filePath );
  	Sql = "SELECT Top " + SHOW_CANT + " Noticias.*, Usuarios.Name, Usuarios.ID as PID  FROM Noticias, Usuarios where (Noticias.Publisher = Usuarios.ID) " +  where + " ORDER BY Fecha_Publicada DESC";

    oRec = Query( Sql, oRec, oConn  );		
    last = "-1";          

    if (oRec.EOF == true)
       {
%>          
<table width="100%" border="0" cellspacing="2" cellpadding="0">
  <tr> 
    <td align=center> 
       <b> No hay más noticias publicadas</b>
    </td>
  </tr>
</table>           

<%
          }
        else
        
     	while (oRec.EOF == false)
    	   
     	    {
              last = oRec.Fields.Item("Fecha_Publicada").value;          

     	      url = oRec.Fields.Item('URL') + "";

     	      im = oRec.Fields.Item('Imagen') + "";

     	     
%>
<table width="100%" border="0" cellspacing="2" cellpadding="0">
  <tr> 
    <td width="15%"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td align="center">
<%

             if (oRec.Fields.Item('Imagen') + "" != "") 
               {
%>             
                 <img src="../news/<%=im %>" width="60" height="60">               
<%
		}
             else
                {
%>
                 <img src="../news/No_Image.gif" width="60" height="60">               
<%
                }

%>
             
         </td>
        </tr>

        <tr> 
          <td align="center"><%=oRec.Fields.Item('Fecha_Publicada') %></td>
        </tr>

<% 
   
    delta = CompareDate(oRec.Fields.Item('Fecha_Publicada'),Today());
    if ((delta <=  MAXNEW) && (delta >= 0))
    
      {
%>                
        <tr> 
          <td align="center"><img src="../images/<%=Session("skin")%>/NewNews.gif" width="40" height="15"></td>
        </tr>

<%
      }
%>                  
      </table>
    </td>
    <td valign="top"> 
      <table width="100%" border="0" cellspacing="2" cellpadding="0">
        <tr> 
          <td width="1%"><b>T&iacute;tulo: </b></td>
          <td width="84%">
<%
	     if (oRec.Fields.Item('URL') != '')	
	        {
%>		
                 <a target=_blank href="<%=url %>" class="ModifyBtnLnk"><%=oRec.Fields.Item('Titulo') %></a>
<%
                }
             else
                {   
%>              
                 <%=oRec.Fields.Item('Titulo') %>
<%
                }

%>       
        <tr>
          <td width="1%" colspan=2><b>Publicado&nbsppor: </b></td>
        </tr>  
        <tr>
          <td width="1%" colspan=2>
            <a href="javascript:abreVentana('',350,280,'userDetails.asp?uid=<%=uid%>&id=<%=oRec.Fields.Item('PID').value%>', 'yes', 'yes')"><%=oRec.Fields.Item('Name')%></a>
          </td>
        </tr>            
        <tr> 
          <td colspan="2"><%=oRec.Fields.Item('Cuerpo') %></td>
        </tr>
      </table>
    </td>
  </tr>
<%        
   if ((Session("PermissionType") == ADMINISTRATOR) || (Session("PermissionType") == PUBLICATOR))
     {
%>
  <tr>
   <td colspan=2 align=right> 
     <a href="modNews.asp?nid=<%=oRec.Fields.Item('ID')%>&uid=<%=uid%>">Modificar</a>
     &nbsp
     <a href="delNews.asp?nid=<%=oRec.Fields.Item('ID')%>">Eliminar</a>
   </td>
   <td> 
   </td>
  </tr>
  <%}%>
  <tr> 
    <td colspan="2"> 
      <hr noshade>
    </td>
  </tr>
</table>
            
<%            
              oRec.Move(1);
           }   
           
       DestroyAdoObjects( oConn, oRec );
%>    

                            
<%        
   if ((Session("PermissionType") == ADMINISTRATOR) || (Session("PermissionType") == PUBLICATOR))
     {
%>
<SCRIPT LANGUAGE=javascript>
<!--
function mysubmit() {
  if (courForm.titulo.value != "") { courForm.submit(); }
}
//-->
</SCRIPT>

 <form id=courForm name=courForm action="publicNews.asp?uid=<%=uid%>" method="post" >
<table border="0" cellspacing="5" cellpadding="5" align="center">
  <tr>
    <td>
      <table border="0" cellspacing="0" cellpadding="0" class="PrologueTable">
        <tr> 
          <td> 
            <table border="0" cellspacing="1" cellpadding="2">
              <tr align="left"> 
                <td colspan="4" class="ToolBar"> 
                  <table border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td class="ToolBar"><b>Titulo:</b></td>
                      <td class="ToolBar"> 
                        <input type="text" id="titulo" name="titulo" maxlength="60" size="80" class="Edit">
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr> 
                <td width="100%" colspan="4" class="MessageTR1"> 
                  <textarea name="cuerpo" rows="13" cols="100"></textarea>
                </td>
              </tr>
              <tr> 
                <td class="ToolBar"> 
                  <table border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td class="ToolBar"><b>Imagen:</b></td>
                      <td class="ToolBar"> 
                        <select name="imagen" class="ComboBox1">
                          <option value="Deportes.gif">Deportes</option>
                          <option value="Cultura.gif">Cultura</option>
                          <option value="Nacionales.gif">Ambito Nacional</option>
                          <option value="Internacionales.gif">Ambito Internacional</option>
                          <option value="others.gif" selected>Otras</option>
                          <option value="No_image.gif">Sin Imagen</option>                          
                        </select>
                      </td>
                    </tr>
                  </table>
                </td>
                <td class="ToolBar"> 
                  <table border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td class="ToolBar"><b>URL:</b></td>
                      <td class="ToolBar"> 
                        <input type="text" name="url" size="25" class="Edit">
                      </td>
                    </tr>
                  </table>
                </td>
                <td colspan="2" class="ToolBar"> 
                  <input type="button" name="Submit" value="Publicar" onclick="mysubmit()">
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table border="0" cellspacing="0" cellpadding="0" class="ToolBar" width="100%">
  <tr>
    <td>
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td><a href="news.asp?uid=<%=uid%>&first=<%=last%>" class="ToolLink">&nbsp;Próximos&nbsp;<%=SHOW_CANT%>&nbsp;</a></td>
        </tr>
      </table>
    </td>
  </tr>
</table>

 </form>
<%     
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

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


  if (Session("PermissionType") == ADMINISTRATOR)
    {
    first = "-1";
    where = "";
    if (Request.QueryString.Count >= 2) {
      first = Request.QueryString.Item("first") + "";
      
      if (first != "-1") {
        where = " where (Name > '" +  first + "')";
      } 
    } 
    
%>

<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<script src="../js/CheckBoxes.js" language="JavaScript"></script>
<script src="../js/windows.js" language="JavaScript"></script>
</HEAD>
<body bgcolor="#FFFFFF" text="#000000">
<form name=Delnews action="ConfirmDeleteMod.asp?uid=<%=uid%>" method="post">
<%    
      
           var filePath = Application('dataPath');
           var oConn;
           var oRec;
 
           oConn = MakeConnection( oConn, filePath );
           Sql = "SELECT Top " + SHOW_CANT + " * FROM Modulo " +  where + " Order By Name";

           oRec = Query( Sql, oRec, oConn  );	
      
      
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar" align="center"><b>Listado de modalidades acad�micas</b></td>
  </tr>
  <tr> 
    <td class="ToolBar" align="center">Nota: Para eliminar una modalidad deben primero eliminarse todos sus m�dulos.</td>
  </tr>
  
  <tr> 
    <td> 
      <table width="100%" border="0" cellspacing="1" cellpadding="0">
        <tr> 
          <td width="5%" class="ToolBar" align="center"></td>
          <td width="75%" class="ToolBar" align="center"><b>Nombre</b></td>
          <td width="10%" class="ToolBar" align="center"><b>Estado</b></td>
          <td width="10%" class="ToolBar" align="center"></td>
          
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

%>
        <tr> 
          <td width="5%" class="<%=clase%>">
<%
         if (empty("select id from Cursos where modulo = " + oRec.Fields.Item("ID").value)) {
%>          
             <input type="checkbox" id=Check<%=i%> name=check<%=i%> >
             <input type="hidden"  name="ID<%=i%>" value="<%=oRec.Fields.Item("ID").value%>">
             
<%
         }
%>          
             
          </td>
          <td width="75%" class="<%=clase%>">
              <%=oRec.Fields.Item("Name").value%>
          </td>
          <td width="10%" nowrap class="<%=clase%>">
   <%  
      switch (oRec.Fields.Item("state").value) {
       case MOD_ACA_NOVISIBLE: {%>No visible<% break;}
       case MOD_ACA_DISABLE  : {%>Deshabilitado<% break;}
       case MOD_ACA_ININSCRIPTION : {%>En inscripci�n<% break;}
       case MOD_ACA_INCOURINSC : {%>En inscripci�n en curso<% break;}
       case MOD_ACA_INCOURSE : {%>En curso<% break;}
       case MOD_ACA_FINISHED : {%>Finalizado<% break;}
       default : {%>Deshabilitado<% break;}
      }       
     %>        
          </td>
          <td width="10%" class="<%=clase%>">
            <a href="javascript:abreVentana('ModalidadAdm',700,400,'modadm/main.asp?uid=<%=uid%>&modulo=<%=oRec.Fields.Item("ID").value%>&courseName=<%=oRec.Fields.Item("Name").value%>', 'no', 'no')">
              <img title="Administrar" src="../images/<%=Session("skin")%>/admin.gif"  width="20" height="20" border="0" >
            </a>
          </td>
          
        </tr> 

<%
          oRec.Move(1);
        }  
        
%>
    <input type="hidden"  name="Count" value="<%=i%>">

<%        
      if (i > 0) 
        {
          oRec.MoveFirst();
//          Response.Write('<tr><td><INPUT id=Buton name=Buton type=Submit value=Borrar></td><td><INPUT id=Buton name=Buton type=Submit value=Regresar onclick="history.back(-1)"></td></tr>');
        }
      else
        {
%>        
          <tr><td align="center" width="100%" colspan="4" class="MessageTR">No quedan m�s elementos.</td></tr>        
<%          
        }  
%>        
<table border="0" cellspacing="0" cellpadding="0" class="ToolBar" width="100%">
  <tr>
    <td>
      <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
        <tr> 
          <td><a href="AgregaMod.asp?uid=<%=uid%>" class="ToolLink">&nbsp;Crear</a></td>
          <td>|</td>
          <td nowrap><a href="javascript:Delnews.submit()" class="ToolLink">&nbsp;Eliminar&nbsp;</a></td>
          <td>|</td>
          <td><a href="#" class="ToolLink" onClick="CheckAll(GetObjectCollection( document, 'FORM'))">&nbsp;Seleccionar&nbsp;Todo&nbsp;</a></td>
          <td>|</td>
          <td><a href="DeleteMod.asp?uid=<%=uid%>&first=<%=last%>" class="ToolLink">&nbsp;Pr�ximos&nbsp;<%=SHOW_CANT%>&nbsp;</a></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<%      
     oRec.Close();
     oConn.Close();   
%>
 </form>
<%
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
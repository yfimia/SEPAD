<%@ Language=JScript %>
<%Response.Expires = -1;%>
<!-- #include file='../../js/adoLibrary.inc' -->
<!-- #include file='../../js/user.inc' -->
<!-- #include file='../../js/course.inc' -->
<!-- #include file='../../js/change.inc' -->

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

%>

<%
 function myTrim(cad) {
   var p = cad.indexOf('>');
   cad = cad.substr(p + 1, cad.length);
   cad = cad.substring(0, cad.indexOf('</'));
   //Response.Write("cad :" + cad);
   //Response.Write("p :" + p);
   
   return cad;
 }

function quitCRLF(cad) {
  return (cad);  
}

 if ((Session("PermissionType") == ADMINISTRATOR) || (Session("userID") == Session("admcursocordinador"))  || (Session("userID") == Session("admcursoOwner")))
    {
      if ((Request.Form("Mname").Count == 0) || (Request.Form("Mname") == ""))
        {
        
           var filePath = Application('dataPath');
           var oConn;
           var oRec;
 
           oConn = MakeConnection( oConn, filePath );
           Sql = "select Cursos.* from Cursos where (Cursos.modulo = " + Session("admcursomodulo") + ") and (Cursos.ID = " + Session("admcurso")  + ")";

           oRec = Query( Sql, oRec, oConn  );	
          

          if (oRec.EOF == true)
            { 
              oConn.Close();   
              
              Response.Redirect("../errorpage.asp?tipo=Error&short=Módulo no válido&desc=El módulo que trata de modificar no existe o fue eliminado."); 
            }  
           
           var prol = GetFileUrl(oRec.Fields.Item("prologue").Value); 
           if (prol != "") {           
             var sXml = "../../Courses/course" + Session("admcurso") + "/" + prol;
             var oXmlDoc = Server.CreateObject("MICROSOFT.XMLDOM");
             oXmlDoc.async = false;
             oXmlDoc.load(Server.MapPath(sXml));
           }  
                      
        
%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../../css/<%=Session("skin")%>/EvalResult.css" type="text/css">

<style>
  LI {list-style-type: none}
</style>

<script src="../../js/windows.js" language="JavaScript"></script>
</head>
<script language=jscript>

function unchange(ss) {
    	var cnt = 0;
    	s = new String(ss);
    	var cl = "<br/>";
    	    brtext = "";
    	while (s.search(cl) != -1) {
         	s = s.replace(cl,brtext);
        } 
    	var cl = "<BR/>";
    	    brtext = "";
    	while (s.search(cl) != -1) {
         	s = s.replace(cl,brtext);
        } 
        
 return(s);        
}


 function change(ss) {
    	var cnt = 0;
    	s = new String(ss);
    	var cl = "\n";
    	    brtext = "<BR/>";
    	while (s.search(cl) != -1) {
         	s = s.replace(cl,brtext); 
        }
 return(s);        
}
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
  

 function addAutor() {
   var text = '';
   var params  = new Array("Adicionar Autor", "");

   text = change(showModalDialog("addText.htm", params, "dialogHeight:200px"));
   if ((text != '') && (text + '' != 'undefined')) {
     if (aut.children.length > 0) {
       var last = aut.children(aut.children.length - 1).id;
       var next = parseInt(last.substr(3, last.length), 10) + 1;
       var index = 'aut' + next;
     }
     else
       var index = 'aut0';   
     
     aut.insertAdjacentHTML("beforeEnd", '<li id="' + index + '"><table ><tr><td  Class="MessageTR" width="98%">' + text + '</td><td  Class="MessageTR" width="1%"><a href="javascript:modifyAutor(' + index + ')">Modificar...</a></td><td  Class="MessageTR" width="1%"><a href="javascript:removeAutor(' + index + ')">Eliminar</a></td><input type="hidden"  name="inaut" value="' + text + '"></tr></table></li>');   
                                             

   }  
 }

 function modifyAutor(obj) {
   var text = '';
   var params  = new Array("Modificar Autor", unchange(obj.getElementsByTagName("input").item(0).value));
   text = change(showModalDialog("addText.htm", params, "dialogHeight:200px"));
   if ((text != '') && (text + '' != 'undefined')) {
     obj.innerHTML = '<table><tr><td  Class="MessageTR" width="98%">' + text + '</td><td  Class="MessageTR" width="1%"><a href="javascript:modifyAutor(' + obj.id + ')">Modificar...</a></td><td  Class="MessageTR" width="1%"><a href="javascript:removeAutor(' + obj.id + ')">Eliminar</a></td><input type="hidden"  name="inaut" value="' + text + '"></tr></table>';   
                     
   }  
 } 
 
 function removeAutor(obj) {
   for (var i = 0; (i < aut.children.length) && (aut.children(i).id != obj.id); i++) {
   }
   
   aut.removeChild(aut.children(i));   
 } 



 function addPClave() {
   var text = '';
   var params  = new Array("Adicionar Palabra Clave", "");
   
   text = change(showModalDialog("addText.htm", params, "dialogHeight:200px"));
   if ((text != '') && (text + '' != 'undefined')) {
     if (pclave.children.length > 0) {
       var last = pclave.children(pclave.children.length - 1).id;
       var next = parseInt(last.substr(6, last.length), 10) + 1;
       var index = 'pclave' + next;
     }
     else
       var index = 'pclave0';   
     
     pclave.insertAdjacentHTML("beforeEnd", '<li id="' + index + '"><table><tr><td Class="MessageTR"  width="98%">' + text + '</td><td Class="MessageTR"  width="1%"><a href="javascript:modifyPClave(' + index + ')">Modificar...</a></td><td Class="MessageTR"  width="1%"><a href="javascript:removePClave(' + index + ')">Eliminar</a></td><input type="hidden"  name="inpclave" value="' + text + '"></tr></table></li>');   
   }  
 }

 function modifyPClave(obj) {
   var text = '';
   var params  = new Array("Modificar Palabra Clave", unchange(obj.getElementsByTagName("input").item(0).value));

   text = change(showModalDialog("addText.htm", params, "dialogHeight:200px"));
   if ((text != '') && (text + '' != 'undefined')) {
     obj.innerHTML = '<table><tr><td Class="MessageTR"  width="98%">' + text + '</td><td Class="MessageTR"  width="1%"><a href="javascript:modifyPClave(' + obj.id + ')">Modificar...</a></td><td Class="MessageTR"  width="1%"><a href="javascript:removePClave(' + obj.id + ')">Eliminar</a></td><input type="hidden"  name="inpclave" value="' + text + '"></tr></table>';   
   }  
 } 
 
 function removePClave(obj) {
   for (var i = 0; (i < pclave.children.length) && (pclave.children(i).id != obj.id); i++) {
   }
   
   pclave.removeChild(pclave.children(i));   
 } 

 function addobj() {
   var text = '';
   var params  = new Array("Adicionar Objetivo", "");

   text = change(showModalDialog("addText.htm", params, "dialogHeight:200px"));
   if ((text != '') && (text + '' != 'undefined')) {
     if (obje.children.length > 0) {
       var last = obje.children(obje.children.length - 1).id;
       var next = parseInt(last.substr(4, last.length), 10) + 1;
       var index = 'obj' + next;
     }
     else
       var index = 'obj0';   
     
     obje.insertAdjacentHTML("beforeEnd", '<li id="' + index + '"><table><tr><td Class="MessageTR"  width="98%">' + text + '</td><td Class="MessageTR"  width="1%"><a href="javascript:modifyobj(' + index + ')">Modificar...</a></td><td Class="MessageTR"  width="1%"><a href="javascript:removeobj(' + index + ')">Eliminar</a></td><input type="hidden"  name="inobj" value="' + text + '"></tr></table></li>');   
   }  
 }

 function modifyobj(obj) {
   var text = '';
   var params  = new Array("Modificar Objetivo", unchange(obj.getElementsByTagName("input").item(0).value));
   
   text = change(showModalDialog("addText.htm", params, "dialogHeight:200px"));
   if ((text != '') && (text + '' != 'undefined')) {
     obj.innerHTML = '<table><tr><td Class="MessageTR"  width="98%">' + text + '</td><td Class="MessageTR"  width="1%"><a href="javascript:modifyobj(' + obj.id + ')">Modificar...</a></td><td Class="MessageTR"  width="1%"><a href="javascript:removeobj(' + obj.id + ')">Eliminar</a></td><input type="hidden"  name="inobj" value="' + text + '"></tr></table>';   
   }  
 } 
 
 function removeobj(obj) {
   for (var i = 0; (i < obje.children.length) && (obje.children(i).id != obj.id); i++) {
   }
   
   obje.removeChild(obje.children(i));   
 } 

 function addcreq() {
   var text = '';
   var params  = new Array("Adicionar Conocimiento Requerido", "");

   text = change(showModalDialog("addText.htm", params, "dialogHeight:200px"));
   if ((text != '') && (text + '' != 'undefined')) {
     if (creq.children.length > 0) {
       var last = creq.children(creq.children.length - 1).id;
       var next = parseInt(last.substr(4, last.length), 10) + 1;
       var index = 'creq' + next;
     }
     else
       var index = 'creq0';   
     
     creq.insertAdjacentHTML("beforeEnd", '<li id="' + index + '"><table><tr><td Class="MessageTR"  width="98%">' + text + '</td><td Class="MessageTR"  width="1%"><a href="javascript:modifycreq(' + index + ')">Modificar...</a></td><td Class="MessageTR"  width="1%"><a href="javascript:removecreq(' + index + ')">Eliminar</a></td><input type="hidden"  name="increq" value="' + text + '"></tr></table></li>');   
   }  
 }

 function modifycreq(obj) {
   var text = '';
   var params  = new Array("Modificar Conocimiento Requerido", unchange(obj.getElementsByTagName("input").item(0).value));

   text = change(showModalDialog("addText.htm", params, "dialogHeight:200px"));
   if ((text != '') && (text + '' != 'undefined')) {
     obj.innerHTML = '<table><tr><td Class="MessageTR"  width="98%">' + text + '</td><td Class="MessageTR"  width="1%"><a href="javascript:modifycreq(' + obj.id + ')">Modificar...</a></td><td Class="MessageTR"  width="1%"><a href="javascript:removecreq(' + obj.id + ')">Eliminar</a></td><input type="hidden"  name="increq" value="' + text + '"></tr></table>';   
   }  
 } 
 
 function removecreq(obj) {
   for (var i = 0; (i < creq.children.length) && (creq.children(i).id != obj.id); i++) {
   }
   
   creq.removeChild(creq.children(i));   
 } 

 function addcal() {

   var params  = new Array("Adicionar Actividad", "Actividad", "Fechas", '', '');

   var text = showModalDialog("add2Text.htm", params, "dialogHeight:200px");
   if ((text[0] != '') && (text[0] + '' != 'undefined')) {
     if (cal.children.length > 0) {
       var last = cal.children(cal.children.length - 1).id;
       var next = parseInt(last.substr(3, last.length), 10) + 1;
       var index = 'cal' + next;
     }
     else
       var index = 'cal0';   
     cal.insertAdjacentHTML("beforeEnd", '<li id="' + index + '"><table><tr><td Class="MessageTR"  width="78%">' + change(text[0]) + '</td><td Class="MessageTR"  width="20%">' + change(text[1]) + '</td><td Class="MessageTR"  width="1%"><a href="javascript:modifycal(' + index + ')">Modificar...</a></td><td Class="MessageTR"  width="1%"><a href="javascript:removecal(' + index + ')">Eliminar</a></td><input type="hidden"  name="incal" value="' + change(text[0]) + '"><input type="hidden"  name="incalf" value="' + change(text[1]) + '"></tr></table></li>');   
   } 
 }

 function modifycal(obj) {

   var params  = new Array("Adicionar Actividad", "Actividad", "Fechas", unchange(obj.getElementsByTagName("input").item(0).value), unchange(obj.getElementsByTagName("input").item(1).value));

   text = showModalDialog("add2Text.htm", params, "dialogHeight:200px");
   if ((text[0] != '') && (text[0] + '' != 'undefined')) {
     obj.innerHTML = '<table><tr><td Class="MessageTR"  width="78%">' + change(text[0]) + '</td><td Class="MessageTR"  width="20%">' + change(text[1]) + '</td><td Class="MessageTR"  width="1%"><a href="javascript:modifycal(' + obj.id + ')">Modificar...</a></td><td Class="MessageTR"  width="1%"><a href="javascript:removecal(' + obj.id + ')">Eliminar</a></td><input type="hidden"  name="incal" value="' + change(text[0]) + '"><input type="hidden"  name="incalf" value="' + change(text[1]) + '"></tr></table>';   
     
   }  
 } 
 
 function removecal(obj) {
   for (var i = 0; (i < cal.children.length) && (cal.children(i).id != obj.id); i++) {
   }
   
   cal.removeChild(cal.children(i));   
 } 

 function addreq() {

   var params  = new Array("Adicionar Requerimiento Técnico", "Requerimiento", "Url", '', '');

   var text = showModalDialog("add2Text.htm", params, "dialogHeight:200px");
   if ((text[0] != '') && (text[0] + '' != 'undefined')) {
     if (req.children.length > 0) {
       var last = req.children(req.children.length - 1).id;
       var next = parseInt(last.substr(3, last.length), 10) + 1;
       var index = 'req' + next;
     }
     else
       var index = 'req0';   
     req.insertAdjacentHTML("beforeEnd", '<li id="' + index + '"><table><tr><td Class="MessageTR"  width="78%">' + change(text[0]) + '</td><td Class="MessageTR"  width="20%">' + change(text[1]) + '</td><td Class="MessageTR"  width="1%"><a href="javascript:modifyreq(' + index + ')">Modificar...</a></td><td Class="MessageTR"  width="1%"><a href="javascript:removereq(' + index + ')">Eliminar</a></td><input type="hidden"  name="inreq" value="' + change(text[0]) + '"><input type="hidden"  name="inrequ" value="' + change(text[1]) + '"></tr></table></li>');   
   } 
 }

 function modifyreq(obj) {

   var params  = new Array("Modificar Requerimiento Técnico", "Requerimiento", "Url", unchange(obj.getElementsByTagName("input").item(0).value), unchange(obj.getElementsByTagName("input").item(1).value));

   text = showModalDialog("add2Text.htm", params, "dialogHeight:200px");
   if ((text[0] != '') && (text[0] + '' != 'undefined')) {
     obj.innerHTML = '<table><tr><td Class="MessageTR"  width="78%">' + change(text[0]) + '</td><td Class="MessageTR"  width="20%">' + change(text[1]) + '</td><td Class="MessageTR"  width="1%"><a href="javascript:modifyreq(' + obj.id + ')">Modificar...</a></td><td Class="MessageTR"  width="1%"><a href="javascript:removereq(' + obj.id + ')">Eliminar</a></td><input type="hidden"  name="inreq" value="' + change(text[0]) + '"><input type="hidden"  name="inrequ" value="' + change(text[1]) + '"></tr></table>';   
     
   }  
 } 
 
 function removereq(obj) {
   for (var i = 0; (i < req.children.length) && (req.children(i).id != obj.id); i++) {
   }
   
   req.removeChild(req.children(i));   
 } 


</script>
<body bgcolor="#FFFFFF" text="#000000">
  <form name=newgroup id=newgroup action="modifyMod.asp?uid=<%=uid%>" method=post>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td class="HeaderTable"  align=center><b>Modificar generalidades del módulo</b></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<%
           if (prol == "") {           
%>           	
  <b>El módulo no ha sido publicado aún. Para publicarlo utilice la herramienta de publicación SEPADHP v. 1.5.</b>
<%  
           }  

%>
    <table border="0" cellspacing="1" cellpadding="1" align="center">
      <tr>
        <td  align=right Class="MessageTR1">
          Nombre:
        </td>
        <td  Class="MessageTR1"align=left >
          <input name=Mname Id=Mname type=text maxlength=250 size=40 value="<%=oRec.Fields.Item("name").Value%>">
        </td>
      </tr>
      <tr>
        <td align=right   Class="MessageTR">
          Estado:
        </td>
        <td align=left   Class="MessageTR">
                           <select name=Mstate>  
                              <option <%= ((oRec.Fields.Item("state").Value == MOD_ACA_NOVISIBLE) ? "selected" : "") %>  value=<%=MOD_ACA_NOVISIBLE%>>No Visible</option>
                              <option <%= ((oRec.Fields.Item("state").Value == MOD_ACA_DISABLE) ? "selected" : "") %> value=<%=MOD_ACA_DISABLE%>>Deshabilitado</option>
                              <option <%= ((oRec.Fields.Item("state").Value == MOD_ACA_INCOURSE) ? "selected" : "") %> value=<%=MOD_ACA_INCOURSE%>>En curso</option>
                              <option <%= ((oRec.Fields.Item("state").Value == MOD_ACA_FINISHED) ? "selected" : "") %> value=<%=MOD_ACA_FINISHED%>>Finalizado</option>
                              <option <%= ((oRec.Fields.Item("state").Value == MOD_ACA_FREE) ? "selected" : "") %> value=<%=MOD_ACA_FREE%>>Libre</option>
                              
                            </select>  
        </td>
      </tr>
      <tr>
      <tr>
        <td align=right   Class="MessageTR">
          Profesor principal:
        </td>
        <td align=left   Class="MessageTR">
          <input disabled name=Mcord Id=Mcord type=text maxlength=250 size=25 value="<% var user = GetUserData(oRec.Fields.Item("owner").Value); Response.Write(user.fullName); %>"><input name=Mname Id=Mname type=button size=20 value="Seleccionar..." onclick="findUser()">
          <input name=Mcordid Id=Mcordid type=hidden maxlength=250 size=25  value="<%=quitCRLF(oRec.Fields.Item("owner").Value)%>">
        </td>
      </tr>
<%
  if (prol != "") {           
%>
      <tr>
      <td colspan=2>
        <table width="90%"  cellpadding="0" cellspacing="0" align="center" class="PrologueTable">
         <tr cellspacing="0" cellspacing="0">
           <td class="ToolBar" width="99%"><b>Autor(es):</b></td>
           <td class="ToolBar" width="1%" align=right>
             <a class="ToolLink" href="javascript:addAutor()" >Adicionar...</a>
           </td>
         </tr>
         <tr>
          <td colspan=2>
            <ul id="aut">
<% 
   var nodeList =  oXmlDoc.getElementsByTagName("Autor");  
   for( var i=0; i < nodeList.length; i++ )  { 
   
%>
              <li id="aut<%=i%>" >
                  <table >
                    <tr>
                      <td Class="MessageTR"  width="98%" Class="MessageTR" >
                         <%=myTrim(nodeList.item(i).xml)%>
                      </td>   
                      <td Class="MessageTR"  width="1%" Class="MessageTR" >
                         <a href="javascript:modifyAutor(aut<%=i%>)">Modificar...</a>
                      </td>   
                      <td Class="MessageTR"  width="1%" Class="MessageTR" >
                         <a href="javascript:removeAutor(aut<%=i%>)">Eliminar</a>
                      </td>   
                      <input type="hidden"  name="inaut" value="<%=myTrim(nodeList.item(i).xml)%>">
                    </tr>
                  </table>
              </li>
<% 
   }
%>      
              
            </ul>
          </td>
         </tr>
        </table>
        </td>
      </tr>
      
      <tr>
      <td colspan=2>
        <table width="90%"  cellpadding="0" cellspacing="0" align="center" class="PrologueTable">
         <tr cellspacing="0" cellspacing="0">
           <td class="ToolBar" width="99%"><b>Palabras Clave:</b></td>
           <td class="ToolBar" width="1%" align=right>
             <a class="ToolLink" href="javascript:addPClave()" >Adicionar...</a>
          </td>
          
         </tr>
         <tr>
          <td colspan=2>
            <ul id="pclave">
<% 
   var nodeList =  oXmlDoc.getElementsByTagName("PalabraClave");  
   for( var i=0; i < nodeList.length; i++ )  { 
   
%>
              <li id="pclave<%=i%>">
                  <table >
                    <tr>
                      <td Class="MessageTR"  width="98%" Class="MessageTR" >
                         <%=myTrim(nodeList.item(i).xml)%>
                      </td>   
                      <td Class="MessageTR"  width="1%" Class="MessageTR" >
                         <a href="javascript:modifyPClave(pclave<%=i%>)">Modificar...</a>
                      </td>   
                      <td Class="MessageTR"  width="1%" Class="MessageTR" >
                         <a href="javascript:removePClave(pclave<%=i%>)">Eliminar</a>
                      </td>   
                      <input type="hidden"  name="inpclave" value="<%=myTrim(nodeList.item(i).xml)%>">
                    </tr>
                  </table>
              </li>
              
<% 
   }

%>      
              
            </ul>
          </td>
         </tr>
        </table>
        <td>
      </tr>

      <tr>
      <td colspan=2>
        <table width="90%"  cellpadding="0" cellspacing="0" align="center" class="PrologueTable">
         <tr cellspacing="0" cellspacing="0">
           <td class="ToolBar" width="99%"><b>Conocimientos Requeridos:</b></td>
           <td class="ToolBar" width="1%" align=right>
             <a class="ToolLink" href="javascript:addcreq()" >Adicionar...</a>
          </td>
          
         </tr>
         <tr>
          <td colspan=2>
            <ul id="creq">
<% 
   var nodeList =  oXmlDoc.getElementsByTagName("Conocimiento");  
   for( var i=0; i < nodeList.length; i++ )  { 
   
%>
              <li id="creq<%=i%>">
                  <table >
                    <tr>
                      <td Class="MessageTR"  width="98%" Class="MessageTR" >
                         <%=myTrim(nodeList.item(i).xml)%>
                      </td>   
                      <td Class="MessageTR"  width="1%" Class="MessageTR" >
                         <a href="javascript:modifycreq(creq<%=i%>)">Modificar...</a>
                      </td>   
                      <td Class="MessageTR"  width="1%" Class="MessageTR" >
                         <a href="javascript:removecreq(creq<%=i%>)">Eliminar</a>
                      </td>   
                      <input type="hidden"  name="increq" value="<%=myTrim(nodeList.item(i).xml)%>">
                    </tr>
                  </table>
              </li>
              
<% 
   }

%>      
              
            </ul>
          </td>
         </tr>
        </table>
        <td>
      </tr>


      <tr>
      <td colspan=2>
        <table width="90%"  cellpadding="0" cellspacing="0" align="center" class="PrologueTable">
         <tr cellspacing="0" cellspacing="0">
           <td class="ToolBar" width="99%"><b>Objetivos:</b></td>
           <td class="ToolBar" width="1%" align=right>
             <a class="ToolLink" href="javascript:addobj()" >Adicionar...</a>
          </td>
          
         </tr>
         <tr>
          <td colspan=2>
            <ul id="obje">
<% 
   var nodeList =  oXmlDoc.getElementsByTagName("Objetivo");  
   for( var i=0; i < nodeList.length; i++ )  { 
   
%>
              <li id="obj<%=i%>">
                  <table >
                    <tr>
                      <td Class="MessageTR"  width="98%" Class="MessageTR" >
                         <%=myTrim(nodeList.item(i).xml)%>
                      </td>   
                      <td Class="MessageTR"  width="1%">
                         <a href="javascript:modifyobj(obj<%=i%>)">Modificar...</a>
                      </td>   
                      <td Class="MessageTR" width="1%">
                         <a href="javascript:removeobj(obj<%=i%>)">Eliminar</a>
                      </td>   
                      <input type="hidden"  name="inobj" value="<%=myTrim(nodeList.item(i).xml)%>">
                    </tr>
                  </table>
              </li>
              
<% 
   }

%>      
              
            </ul>
          </td>
         </tr>
        </table>
       </td> 
      </tr>


      <tr>
      <td colspan=2>
        <table width="90%"   cellpadding="0" cellspacing="0" align="center" class="PrologueTable">
         <tr  cellspacing="0" cellspacing="0">
           <td  class="ToolBar" width="99%"><b>Calendario de Actividades:</b></td>
           <td  class="ToolBar" width="1%" align=right>
             <a class="ToolLink" href="javascript:addcal()" >Adicionar...</a>
          </td>
          
         </tr>
         <tr>
         
          <td colspan=2>
            <ul id="cal">
<% 
   var nodeList =  oXmlDoc.getElementsByTagName("Actividad");  
   for( var i=0; i < nodeList.length; i++ )  { 
   
%>
              <li id="cal<%=i%>">
                  <table >
                    <tr>
                      <td Class="MessageTR" width="78%">
                         <%=myTrim(nodeList.item(i).xml)%>
                      </td>   
                      <td Class="MessageTR" width="20%">
<%
      if (nodeList.item(i).attributes.length > 0) {
%>                      

                         <%=nodeList.item(i).attributes.getNamedItem("fecha").text%>
<%
     }
%>                         
                      </td>   
                      <td Class="MessageTR" width="1%">
                         <a href="javascript:modifycal(cal<%=i%>)">Modificar...</a>
                      </td>   
                      <td Class="MessageTR" width="1%">
                         <a href="javascript:removecal(cal<%=i%>)">Eliminar</a>
                      </td>   
                      <input type="hidden"  name="incal" value="<%=myTrim(nodeList.item(i).xml)%>">
                      <input type="hidden"  name="incalf" value="<%=nodeList.item(i).attributes.getNamedItem("fecha").text%>">
                      
                    </tr>
                  </table>
              </li>
              
<% 
   }

%>      
              
            </ul>
          </td>
         </tr>
        </table>
        </td>
      </tr>

      <tr>
        <td colspan=2>
        <table width="90%"   cellpadding="0" cellspacing="0" align="center" class="PrologueTable">
         <tr  cellspacing="0" cellspacing="0">
           <td  class="ToolBar" width="99%"><b>Requerimientos Técnicos:</b></td>
           <td  class="ToolBar" width="1%" align=right>
             <a class="ToolLink" href="javascript:addreq()" >Adicionar...</a>
          </td>
          
         </tr>
         <tr>
         
          <td colspan=2>
            <ul id="req">
<% 
   var nodeList =  oXmlDoc.getElementsByTagName("Requerimiento");  
   for( var i=0; i < nodeList.length; i++ )  { 
   
%>
              <li id="req<%=i%>">
                  <table >
                    <tr>
                      <td Class="MessageTR" width="78%">
                         <%=myTrim(nodeList.item(i).xml)%>
                      </td>   
                      <td Class="MessageTR" width="20%">
<%
      if (nodeList.item(i).attributes.length > 0) {
%>                      
                         <%=nodeList.item(i).attributes.getNamedItem("url").text%>
<%
     }
%>                         
                      </td>   
                      <td Class="MessageTR" width="1%">
                         <a href="javascript:modifyreq(req<%=i%>)">Modificar...</a>
                      </td>   
                      <td Class="MessageTR" width="1%">
                         <a href="javascript:removereq(req<%=i%>)">Eliminar</a>
                      </td>   
                      <input type="hidden"  name="inreq" value="<%=myTrim(nodeList.item(i).xml)%>">
                      <input type="hidden"  name="inrequ" value="<%=nodeList.item(i).attributes.getNamedItem("url").text%>">
                      
                    </tr>
                  </table>
              </li>
              
<% 
   }
%>      
              
            </ul>
          </td>
         </tr>
        </table>
        </td>
      </tr>
    </table>

    <table width=90% border="0" cellspacing="1" cellpadding="1" align="center">
      
      
<%
  var val = '';
  if (oXmlDoc.getElementsByTagName("OrientacionesGenerales").length > 0) 
    var val = oXmlDoc.getElementsByTagName("OrientacionesGenerales").item(0).xml; 
%>

      <tr>
        <td colspan=2 align=center   Class="MessageTR">
          Orientaciones Generales
        </td>
      </tr>
      <tr>
        <td colspan=2 align=left   Class="MessageTR">
          <textarea name=Mogen  rows="8" cols="78" maxlength=250><%=myTrim(unchange1(val))%></textarea>
        </td>
      </tr>
<%
  var val = '';
  if (oXmlDoc.getElementsByTagName("NotasAdicionales").length > 0) 
    var val = oXmlDoc.getElementsByTagName("NotasAdicionales").item(0).xml; 
%>

      <tr>
        <td colspan=2 align=center   Class="MessageTR">
          Sistema de Evaluación
        </td>
      </tr>
      <tr>  
        <td colspan=2  align=left   Class="MessageTR">
          <textarea name=Mscont  rows="8" cols="78" maxlength=250><%=unchange1(myTrim(val))%></textarea>
        </td>
      </tr>

<%
  var val = '';
  if (oXmlDoc.getElementsByTagName("Fundament").length > 0) 
    var val = oXmlDoc.getElementsByTagName("Fundament").item(0).xml; 
%>

      <tr>
        <td colspan=2 align=center   Class="MessageTR">
          Fundamentación
        </td>
      </tr>
      <tr>  
        <td colspan=2  align=left   Class="MessageTR">
          <textarea name=Mfund  rows="8" cols="78" maxlength=250><%=unchange1(myTrim(val))%></textarea>
        </td>
      </tr>
<%
  var val = '';
  if (oXmlDoc.getElementsByTagName("Programa").length > 0) 
    var val = oXmlDoc.getElementsByTagName("Programa").item(0).xml; 
%>

      <tr>
        <td colspan=2 align=center   Class="MessageTR">
          Programa
        </td>
      </tr>
      <tr>  
        <td colspan=2  align=left   Class="MessageTR">
          <textarea name=Mprog  rows="8" cols="78" maxlength=250><%=unchange1(myTrim(val))%></textarea>
        </td>
      </tr>

<%
  var val = '';
  if (oXmlDoc.getElementsByTagName("Recursos").length > 0) 
    var val = oXmlDoc.getElementsByTagName("Recursos").item(0).xml; 
%>


      <tr>
        <td colspan=2 align=center   Class="MessageTR">
          Recursos
        </td>
      </tr>
      <tr>  
        <td colspan=2  align=left   Class="MessageTR">
          <textarea name=Mrec  rows="8" cols="78" maxlength=250><%=unchange1(myTrim(val))%></textarea>
        </td>
      </tr>

<%
  var val = '';
  if (oXmlDoc.getElementsByTagName("SistemaTutorias").length > 0) 
    var val = oXmlDoc.getElementsByTagName("SistemaTutorias").item(0).xml; 
%>

     <tr>
        <td colspan=2 align=center   Class="MessageTR">
          Sistema de Tutorías
        </td>
      </tr>
      <tr>  
        <td colspan=2  align=left   Class="MessageTR">
          <textarea name=Mstut  rows="8" cols="78" maxlength=250><%=unchange1(myTrim(val))%></textarea>
        </td>
      </tr>

<%
  var val = '';
  if (oXmlDoc.getElementsByTagName("Bibliog").length > 0) 
    var val = oXmlDoc.getElementsByTagName("Bibliog").item(0).xml; 
%>

     <tr>
        <td colspan=2 align=center   Class="MessageTR">
          Bibliografía General
        </td>
      </tr>
      <tr>  
        <td colspan=2  align=left   Class="MessageTR">
          <textarea name=Mbibliog  rows="8" cols="78" maxlength=250><%=unchange1(myTrim(val))%></textarea>
        </td>
      </tr>

<%
 } //fin de if (prol != "")
%>

      <tr>
        <td align=right   Class="MessageTR">
          Cantidad de ejercicios por evaluaciones de la lección actual(numérico):
        </td>
        <td align=left   Class="MessageTR">
          <input name=MctdExAct type=text maxlength=5 size=5  value="<%=quitCRLF(oRec.Fields.Item("ctdExAct").Value)%>">
        </td>
      </tr>

     <tr>
        <td align=right   Class="MessageTR">
           Cantidad de ejercicios por evaluaciones de lecciones anteriores(numérico): 
        </td>
        <td align=left   Class="MessageTR">
          <input name=MctdExRet type=text maxlength=5 size=5  value="<%=quitCRLF(oRec.Fields.Item("ctdExRet").Value)%>">
        </td>
      </tr>


      
      <tr> 
         <td align=left   Class="MessageTR">Cantidad de evaluaciones a tener en 
          cuenta en la detecci&oacute;n de problemas(numérico): </td>
        <td align=left   Class="MessageTR"> 
          <input type="text" name="Matmost" size="5" maxlength=5 value="<%=quitCRLF(oRec.Fields.Item("atmost").Value)%>">
        </td>
      </tr>

      <tr>
        <td align=right   Class="MessageTR">
          Tamaño en kilobytes del directorio destinado para recibir los archivos que envien los alumnnos(numérico): 
        </td>
        <td align=left   Class="MessageTR">
<%
   if ((Session("PermissionType") == ADMINISTRATOR)) {
%>
          <input name=Mmaxmem  type=text maxlength=5 size=5  value="<%=quitCRLF(oRec.Fields.Item("maxmem").Value)%>">
<% }
   else {
%>
      <%=quitCRLF(oRec.Fields.Item("maxmem").Value)%>
<%
   }
%>
        </td>
      </tr>

     <tr>
        <td align=right   Class="MessageTR">
           Tamaño máximo en kilobytes de los archivos que pueden subir los alumnos(numérico):
        </td>
        <td align=left   Class="MessageTR">
<%
   if ((Session("PermissionType") == ADMINISTRATOR)) {
%>
        
          <input name=Mmaxfile type=text maxlength=5 size=5 value="<%=quitCRLF(oRec.Fields.Item("maxfile").Value)%>">
<% }
   else {
%>
      <%=quitCRLF(oRec.Fields.Item("maxfile").Value)%>
<%
   }
%>
          
        </td>
      </tr>
    
      <tr>
        <td widht="100%" colspan=2 align=center Class="Toolbar">
          <center><input type=button value="Actualizar" onclick="formSubmit()"></center>
        </td>      
      </tr>
    </table>
  </form>
<SCRIPT LANGUAGE=javascript>
<!--
function formSubmit()
{ 
   if ( isNaN(parseInt(newgroup.MctdExAct.value)) || isNaN(parseInt(newgroup.MctdExRet.value)) 
<% if (Session("PermissionType") == ADMINISTRATOR) {%>   
        ||  
        isNaN(parseInt(newgroup.Mmaxfile.value)) || isNaN(parseInt(newgroup.Mmaxmem.value))
<% }%>           
      )
    {
     alert('Error, chequee que los campos numéricos se hallan introducidos correctamente.');    
    }
  else newgroup.submit();
    
}
 
//-->
</SCRIPT>  
<%
              oRec.Close();   
              oConn.Close();   

        }
      else
        {
           var filePath = Application('dataPath');
           var oConn;
           var oRec;
 
           oConn = MakeConnection( oConn, filePath );
           Sql = "select Cursos.* from Cursos where (Cursos.modulo = " + Session("admcursomodulo") + ") and (Cursos.ID = " + Session("admcurso")  + ")";
           
           //Sql = "select Cursos.*, Ficheros.url from Cursos, Ficheros where (Cursos.prologue = Ficheros.ID) and (Cursos.modulo = " + Session("admcursomodulo") + ") and (Cursos.ID = " + Session("admcurso")  + ")";


           oRec = Query( Sql, oRec, oConn  );	
          
          
           var prol = GetFileUrl(oRec.Fields.Item("prologue").Value); 
          

          if (oRec.EOF == true)
            { 
              oConn.Close();   
              
              Response.Redirect("../errorpage.asp?tipo=Error&short=Módulo no válido&desc=El módulo que trata de modificar no existe o fue eliminado."); 
            }  
          else
            {
            
            
              if ((Request.Form("MName").Count != 0) && (Request.Form("Mname") + "" != ""))
                { 
                  oRec.Fields.Item("Name").Value = Request.Form("Mname") + "";
                }
              else Response.Redirect("../errorpage.asp?tipo=Error&short=Nombre no válido&desc=El nombre del módulo introducido no es válido o se ha dejado en blanco.");

              if ((Request.Form("MctdExAct").Count != 0) && (!isNaN(parseInt(Request.Form("MctdExAct"),10))))
                { 
                  oRec.Fields.Item("ctdExAct").Value = parseInt(Request.Form("MctdExAct")) + "";
                }
              else Response.Redirect("../errorpage.asp?tipo=Error&short=Campo no válido&desc=Chequee que los campos numéricos se hallan introducidos correctamente.");

              if ((Request.Form("MctdExRet").Count != 0) && (!isNaN(parseInt(Request.Form("MctdExRet") + ""))))
                { 
                  oRec.Fields.Item("ctdExRet").Value = parseInt(Request.Form("MctdExRet")) + "";
                }
              else Response.Redirect("../errorpage.asp?tipo=Error&short=Campo no válido&desc=Chequee que los campos numéricos se hallan introducidos correctamente.");

              if (Session("PermissionType") == ADMINISTRATOR) {
                if ((Request.Form("Mmaxmem").Count != 0) && (!isNaN(parseInt(Request.Form("Mmaxmem") + ""))))
                  { 
                    oRec.Fields.Item("MaxMem").Value = parseInt(Request.Form("Mmaxmem")) + "";
                  }
                else Response.Redirect("../errorpage.asp?tipo=Error&short=Campo no válido&desc=Chequee que los campos numéricos se hallan introducidos correctamente.");
  
                if ((Request.Form("Mmaxfile").Count != 0) && (!isNaN(parseInt(Request.Form("Mmaxfile") + ""))))
                  { 
                    oRec.Fields.Item("MaxFile").Value = parseInt(Request.Form("Mmaxfile")) + "";
                  }
                else Response.Redirect("../errorpage.asp?tipo=Error&short=Campo no válido&desc=Chequee que los campos numéricos se hallan introducidos correctamente.");
              }  

              if ((Request.Form("Matmost").Count != 0) && (!isNaN(parseInt(Request.Form("Matmost") + ""))))
                { 
                  oRec.Fields.Item("atmost").Value = parseInt(Request.Form("Matmost")) + "";
                }
              else Response.Redirect("../errorpage.asp?tipo=Error&short=Campo no válido&desc=Chequee que los campos numéricos se hallan introducidos correctamente.");

              oRec.Fields.Item("state").Value = parseInt(Request.Form("Mstate")) + "";

              if ((Request.Form("Mcordid").Count != 0) && (Request.Form("Mcordid") + "" != ""))
                { 
                  oRec.Fields.Item("owner").Value = Request.Form("Mcordid") + "";
                }
              else Response.Redirect("../errorpage.asp?tipo=Error&short=Profesor principal no válido&desc=El profesor principal no ha sido seleccionado apropiadamente. Para seleccionar el profesor principal, precione el botón Seleccionar y escójalo de la lista.");



              //actualizando el prologo xml....
                      
              
              var xmlText = '<Prologo>';
                  
              //Titulo
              xmlText = xmlText + '<Titulo>' + Request.Form("Mname") +  '</Titulo>';

              //Autores 
              for( var i=0; i < Request.Form("inaut").Count; i++ )  { 
                xmlText = xmlText + '<Autor>' + Request.Form("inaut")(i + 1) + '</Autor>';
              }

              //Palabras clave
              xmlText = xmlText + '<PalabrasClaves>';
              for( var i=0; i < Request.Form("inpclave").Count; i++ )  { 
                xmlText = xmlText + '<PalabraClave>' + (Request.Form("inpclave")(i + 1)) + '</PalabraClave>';
              }
                  xmlText = xmlText + '</PalabrasClaves>';
                  
              //Objetivos
              xmlText = xmlText + '<Objetivos>';
              for( var i=0; i < Request.Form("inobj").Count; i++ )  { 
                xmlText = xmlText + '<Objetivo>' + (Request.Form("inobj")(i + 1)) + '</Objetivo>';
              }
                  xmlText = xmlText + '</Objetivos>';

              //Conocimientos Requeridos
              xmlText = xmlText + '<ConocimientosRequeridos>';
              for( var i=0; i < Request.Form("increq").Count; i++ )  { 
                xmlText = xmlText + '<Conocimiento>' + (Request.Form("increq")(i + 1)) + '</Conocimiento>';
              }
                  xmlText = xmlText + '</ConocimientosRequeridos>';
                  
              //Calendario de actividades
              xmlText = xmlText + '<Calendario>';
              for( var i=0; i < Request.Form("incal").Count; i++ )  { 
                xmlText = xmlText + '<Actividad fecha="' + Request.Form("incalf")(i + 1)  + '">' + (Request.Form("incal")(i + 1)) +  '</Actividad>';
              }
                  xmlText = xmlText + '</Calendario>';


              //RequerimientosTecnicos
              xmlText = xmlText + '<RequerimientosTecnicos>';
              for( var i=0; i < Request.Form("inreq").Count; i++ )  { 
                xmlText = xmlText + '<Requerimiento url="' + Request.Form("inrequ")(i + 1)  + '">' + (Request.Form("inreq")(i + 1)) +  '</Requerimiento>';
              }
                  xmlText = xmlText + '</RequerimientosTecnicos>';

              //Orientaciones Generales
              xmlText = xmlText + '<OrientacionesGenerales>';
              xmlText = xmlText + change(Request.Form("Mogen") + "");
              xmlText = xmlText + '</OrientacionesGenerales>';

              //Notas Adicionales
              xmlText = xmlText + '<NotasAdicionales>';
              xmlText = xmlText + change(Request.Form("Mscont") + "");
              xmlText = xmlText + '</NotasAdicionales>';


              //Fundament
              xmlText = xmlText + '<Fundament>';
              xmlText = xmlText + change(Request.Form("Mfund") + "");
              xmlText = xmlText + '</Fundament>';

              //Recursos
              xmlText = xmlText + '<Recursos>';
              xmlText = xmlText + change(Request.Form("Mrec") + "");
              xmlText = xmlText + '</Recursos>';

              //Programa
              xmlText = xmlText + '<Programa>';
              xmlText = xmlText + change(Request.Form("Mprog") + "");
              xmlText = xmlText + '</Programa>';

              //SistemaTutorias
              xmlText = xmlText + '<SistemaTutorias>';
              xmlText = xmlText + change(Request.Form("Mstut") + "");
              xmlText = xmlText + '</SistemaTutorias>';

              //Bibliog
              xmlText = xmlText + '<Bibliog>';
              xmlText = xmlText + change(Request.Form("Mbibliog") + "");
              xmlText = xmlText + '</Bibliog>';

              xmlText = xmlText + '</Prologo>';
  
  
  
              //Salvando el prologo
           var prol = GetFileUrl(oRec.Fields.Item("prologue").Value); 
           if (prol != "") {           
             var sXml = "../../Courses/course" + Session("admcurso") + "/" + prol;
             var oXmlDoc = Server.CreateObject("MICROSOFT.XMLDOM");
             oXmlDoc.async = false;
             oXmlDoc.loadXML(xmlText);
             oXmlDoc.save(Server.MapPath(sXml));
           }  
              
               



              Session("admcurso")        = oRec.Fields.Item("ID").Value;
              Session("admcursoName")    = oRec.Fields.Item("Name").Value;
              Session("admcursoState")        = oRec.Fields.Item("state").Value;
              Session("admcursoOwner")    = oRec.Fields.Item("owner").Value;
             
              oRec.Update();      
              oConn.Close();   
              
              Response.Redirect("modifyMod.asp?uid=" + uid);      
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



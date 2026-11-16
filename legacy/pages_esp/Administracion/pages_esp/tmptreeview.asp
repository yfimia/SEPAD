<%@ Language=jScript %>
<%
 Response.Expires = -1;
 
 MAXLEVEL = 50; //Maximo # de niveles que puede tener el treeview.
%>
<!-- #include file='../js/user.inc' -->
<!-- #include file='inc/tmptreeview.inc' -->
<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

%>


<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/listspan.css" type="text/css">
<script type="text/javascript" language="javascript" src="../js/listspan.js"></script>
<script type="text/javascript" language="javascript" src="../js/Treeview.js"></script>  
</head>
<body bgcolor="#FFFFFF" text="#000000" class="TreeViewBody" style="overflow-y:scroll">

<%
    var niveles = new Array();  
    var lines = new Array();
    var lid = new Array();  
    var dir = new Array();
    var state = new Array();    
    
    var Ecant = 0;
    
    
    for (Ecant = 0;Ecant <= Session("lessonCant");Ecant++)
      {
        
        niveles[Ecant] = parseInt(Session("lessons").items.nivel[Ecant],10);
        lines[Ecant] = "" + Session("lessons").items.name[Ecant];
        lid[Ecant] = "" + Session("lessons").items.id[Ecant];
        dir[Ecant] = "" + Session("lessons").items.dir[Ecant];
        state[Ecant] = "" + Session("lessons").items.state[Ecant];
//        Response.Write(lines[Ecant] + " -- Nivel : " + niveles[Ecant] + "<br>"  + state[Ecant] + "<br>");
      }
    

    var levels = 0;
    var p;
    
    function setTree(pos) {
      var incOk = true;
      if ((levels < MAXLEVEL) && (Ecant > 0)) {
        levels = levels + 1;
        p = pos;
        while ((incOk == true) && (p < Ecant) && (niveles[p]  == niveles[pos] )) {
//            Response.Write(lines[p] + "  "+ niveles[p + 1]  +"  >  " + niveles[pos] + "<br>");
//            Response.Write(p + " != "+ (Ecant - 1) + "<br>");
//            Response.Write(((p == (Ecant - 1)) ||(niveles[p + 1]  == niveles[pos] )) + "<br>");            
            if ((p != (Ecant - 1) ) && (niveles[p + 1] == niveles[pos] + 1)) { //si la proxima no es la ultima y tiene hijos
%>          
      <li class="clsHasKidsCollapsed"> 
<%
      if ((Session("userID") == Session("courseOwner")) || (Session("userID") == Session("cordinador")) || (Session("PermissionType") == ADMINISTRATOR))  {
%>
        
<%      
                   if (state[p] == MOD_ACA_NOVISIBLE) {       
%>                 
            <a name="Lesson<%=p%>" href="javascript:click('docs.asp?uid=<%=uid%>&lesson=<%=lid[p]%>&lessonurl=<%=dir[p]%>&lessonindex=<%=p%>&lessonname=<%=lines[p]%>',<%=lid[p]%>&material=)"> 
              <span>

                   <IMG SRC="../images/<%=Session("skin")%>/No_Visible.gif"   class="IconClass" />  <%=lines[p]%>
              </span>
            </a>              
                   
<%      
                   }
                   else if (state[p] == MOD_ACA_DISABLE) {       
%>                   
            <a name="Lesson<%=p%>" href="javascript:click('docs.asp?uid=<%=uid%>&lesson=<%=lid[p]%>&lessonurl=<%=dir[p]%>&lessonindex=<%=p%>&lessonname=<%=lines[p]%>',<%=lid[p]%>)"> 
              <span>

                   <IMG  SRC="../images/<%=Session("skin")%>/No_Accesible.gif"   class="IconClass" />  <%=lines[p]%>
              </span>
            </a>              
                   
<%      
                   }
                   else {
%>      
            <a name="Lesson<%=p%>" href="javascript:click('docs.asp?uid=<%=uid%>&lesson=<%=lid[p]%>&lessonurl=<%=dir[p]%>&lessonindex=<%=p%>&lessonname=<%=lines[p]%>',<%=lid[p]%>)"> 
              <span>

                   <IMG SRC="../images/<%=Session("skin")%>/TreeViewImg.gif"   class="IconClass" />  <%=lines[p]%>
              </span>
            </a>              
                   
<%                   
                   }
%>      
        
        
                 
<%      
          } else
          if (state[p] == MOD_ACA_DISABLE) {              //IMAGEN para link desabilitado
%>      
              <span>
                 <IMG title="<%= TITLE1 %>" SRC="../images/<%=Session("skin")%>/No_Accesible.gif"   class="IconClass" />  
                 <%=lines[p]%>
              </span>
<%      
          }
          else  if (state[p] == MOD_ACA_INCOURSE) {
%>
            <a name="Lesson<%=p%>" href="javascript:click('docs.asp?uid=<%=uid%>&lesson=<%=lid[p]%>&lessonurl=<%=dir[p]%>&lessonindex=<%=p%>&lessonname=<%=lines[p]%>',<%=lid[p]%>)"> 
              <span>

                   <IMG SRC="../images/<%=Session("skin")%>/TreeViewImg.gif"   class="IconClass" />  <%=lines[p]%>
              </span>
            </a>              

<%          
          }
%>      
              <ul>
<%                 
                   incOk = true;
                   p = p + 1;
                   levels = levels + 1;
                   setTree(p);
                   levels = levels - 1;
%>      
              </ul>     
          </li>       
<%          
            }
             else if (niveles[p] == niveles[pos]) { //si se mantien en el mismo nivel
              incOk = true;
              if ((Session("userID") == Session("courseOwner")) || (Session("userID") == Session("cordinador")) || (Session("PermissionType") == ADMINISTRATOR))  {
%>            
            <li class="clsDocument" > 
<%         
                     if (state[p] == MOD_ACA_NOVISIBLE) {       
%>                   
              <a  name="Lesson<%=p%>" href="javascript:click('docs.asp?uid=<%=uid%>&lesson=<%=lid[p]%>&lessonurl=<%=dir[p]%>&lessonindex=<%=p%>&lessonname=<%=lines[p]%>',<%=lid[p]%>)">
                <span>

                     <IMG SRC="../images/<%=Session("skin")%>/No_Visible.gif"   class="IconClass" />  <%=lines[p]%>
                 </span>
               </a>
                     
<%         
                     }
                     else if (state[p] == MOD_ACA_DISABLE) {       
%>                     
              <a  name="Lesson<%=p%>" href="javascript:click('docs.asp?uid=<%=uid%>&lesson=<%=lid[p]%>&lessonurl=<%=dir[p]%>&lessonindex=<%=p%>&lessonname=<%=lines[p]%>',<%=lid[p]%>)">
                <span title="<%= TITLE1 %>" >

                     <IMG SRC="../images/<%=Session("skin")%>/No_Accesible.gif"   class="IconClass" />  <%=lines[p]%>
                 </span>
               </a>
                     
<%         
                     }
                     else {
%>         
              <a  name="Lesson<%=p%>" href="javascript:click('docs.asp?uid=<%=uid%>&lesson=<%=lid[p]%>&lessonurl=<%=dir[p]%>&lessonindex=<%=p%>&lessonname=<%=lines[p]%>',<%=lid[p]%>)">
                <span>

                     <IMG SRC="../images/<%=Session("skin")%>/TreeViewImg.gif"   class="IconClass" />  <%=lines[p]%>
                 </span>
               </a>
                     
<%                     
                     }
%>

             </li>             

<%                     
            
              } else {
                   if (state[p] == MOD_ACA_DISABLE) {              //IMAGEN para link desabilitado
%>                 
            <li class="clsDocument" > 
                <span title="<%= TITLE1 %>" >
                          <IMG title="<%= TITLE1 %>" SRC="../images/<%=Session("skin")%>/No_Accesible.gif"   class="IconClass" />  
                          <%=lines[p]%>
                 </span>
             </li>             
                       
<%                 
                } else {
                   if (state[p] == MOD_ACA_INCOURSE) {
%>
            <li class="clsDocument" > 
              <a  name="Lesson<%=p%>" href="javascript:click('docs.asp?uid=<%=uid%>&lesson=<%=lid[p]%>&lessonurl=<%=dir[p]%>&lessonindex=<%=p%>&lessonname=<%=lines[p]%>',<%=lid[p]%>)">
                <span>
                          <IMG SRC="../images/<%=Session("skin")%>/TreeViewImg.gif"   class="IconClass" />  
                          <%=lines[p]%>
                 </span>
               </a>
             </li>             


<%               } }   
               }
               
                 p = p + 1;
              }
            else incOk = false;   
          }            
      }  
    }  
%>    

<table border="0" celsspacing="0" cellspadding="0" width="100%" class="ToolBar">
  <tr>
    <td>
      <table border="0" celsspacing="0" cellspadding="0">
        <tr>
          <td><a class="ToolLink" href="javascript:Do_Expand_All()">&nbsp;<%= TEXT1 %>&nbsp;</a></td>
          <td>|</td>
          <td><a class="ToolLink" href="javascript:Do_Collapse_All()">&nbsp;<%= TEXT2 %>&nbsp;</a></td>
        </tr> 
      </table>
    </td>
  </tr>
  <tr>
    <td align=center>
      <h5><%= TEXT3 %></h5>
    </td>
  </tr>
  
</table>

 <ul>
   <%= setTree(0)    %>  
 </ul>
 
 <script language="jscript">

  function click(url, lessonid) {
   parent.frames("mainFrame").frames("mainFrame1").location.href = url;
   parent.frames("mainFrame").frames("bottomFrame").execScript(" fullCombo()                       ");   
   parent.frames("mainFrame").frames("bottomFrame").execScript(" lastLesson = " + lessonid);   
   
  }

 Do_Expand_All();
 LocateChild(<%=Session("lastLessonIndex")%>);
</script>
</body>
</html>
<%
 }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>

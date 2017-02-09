<%
function SelectCourses() {
          var dataPath1 = Application('dataPath');
          var oConn1;
          var oRec1;
          oConn1 = MakeConnection(oConn1, dataPath1);
          SQL = "SELECT ID, Name FROM Cursos order By Name";
          oRec1 = Query(SQL, oRec1, oConn1);
          while (oRec1.EOF == false)
            { 
%>
               <OPTION title="<%=oRec1.Fields.Item("Name")%>" value="<%=oRec1.Fields.Item("ID")%>"><%=oRec1.Fields.Item("Name")%></OPTION> 

<%            
              oRec1.Move(1);
            }
          DestroyAdoObjects(oConn1, oRec1);

}

%>


<SCRIPT LANGUAGE=javascript>
<!--
function search() {
  
  if ((pattern.value.indexOf("'") == -1) && (pattern.value.indexOf(">") == -1) && (pattern.value.indexOf("<") == -1)) {
<% 
  if (tipo == 0) {
%>
   location.href = "search.asp?uid=<%=uid%>&courseID=" + SearchBy.value + "&pattern=" + pattern.value;
<%
  }
  else {
%>  
    parent.frames("mainFrame").location.href = "search.asp?uid=<%=uid%>&courseID=" + SearchBy.value + "&pattern=" + pattern.value;    
<%
  }
%>   
 }
 else
   alert("Ha introducido caracteres no válidos para las búsquedas.");
}
//-->
</SCRIPT>

<TABLE border=0 cellPadding=0 cellSpacing=1 width=167>
  
  <TR class=MainTRTTI>
    <TD align=middle colSpan=2 class="ToolBar">Búsquedas</TD></TR>
  <TR class=OrdinalTR>
    <TD align=middle colSpan=2>
      <TABLE border=0 cellPadding=0 cellSpacing=2 width="100%">
        
        <TR align=middle class=LogTR>
          <TD colSpan=3 vAlign=center><INPUT value="<%=pattern%>" class=SearchEdit name="textfield" id="pattern"  
            size=26> </TD></TR>
        <TR align=middle class=LogTR>
          <TD width="34%">Módulo</TD>
          <TD align=left colSpan=2><SELECT class=SearchCombo name=SearchBy id=SearchBy > 
              <OPTION  value=-1 selected>Todos</OPTION>

              <%=SelectCourses()%>
              </SELECT> </TD></TR>
        <TR align=middle class=LogTR>
          <TD colSpan=2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
<INPUT class=OkButton name=SearchButton type=button value=Buscar onclick="search()"> 
</TD>
          <TD width="23%"><!--IMG align=absMiddle height=35 
            src="../images/Lens.gif" 
  width=40--></TD></TR></TABLE></TD></TR></TABLE>

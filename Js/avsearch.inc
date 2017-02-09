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
   location.href = "avsearch.asp?uid=<%=uid%>&courseID=" + SearchBy.value + "&pattern=" + pattern.value;
<%
  }
  else {
%>  
    location.href = "avsearch.asp?uid=<%=uid%>&courseID=" + SearchBy.value + "&pattern=" + pattern.value;    
<%
  }
%>   
 }
 else
   alert("Ha introducido caracteres no válidos para las búsquedas.");
}
//-->
</SCRIPT>

<TABLE border=0 align=center cellPadding=0 cellSpacing=0 class="PrologueTable" width=167>
  
  <TR>
    <TD class="ToolBar" align="center" >Búsquedas</TD>
  </TR>
  <TR >
    <TD align=center>
      <TABLE border=0 cellPadding=0 cellSpacing=2 width="100%">
        
        <TR align=center class=MessageTR1>
          <TD colSpan=2 vAlign=middle><INPUT value="<%=pattern%>" class=SearchEdit name="textfield" id="pattern"  size=26> 
          </TD>
        </TR>
        <TR align=center class=MessageTR>
          <TD width="34%">Módulo</TD>
          <TD align=left>
              <SELECT class=SearchCombo name=SearchBy id=SearchBy > 
                 <OPTION selected title="<%=Session("courseName")%>" value="<%=Session("course")%>"><%=Session("courseName")%></OPTION> 
              </SELECT> 
          </TD>
        </TR>
        <TR align=center class=MessageTR1>
          <TD colSpan=2 align=center>
             <INPUT class=OkButton name=SearchButton type=button value=Buscar onclick="search()"> 
          </TD>

        </TR>
       </TABLE>
      </TD>
   </TR>
</TABLE>

<%@ Language=JavaScript %>

<%
 Response.Expires = -1;
%>
<!-- #include file='../../js/user.inc' -->
<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       


 if ((Session("isTutor") > -1) || (Session("PermissionType") == ADMINISTRATOR) || (Session("userID") == Session("tutcursocordinador"))  || (Session("userID") == Request.QueryString.Item("courseOwner")))
    {

%>


<%

   var HTML  = "";

   if (Request.QueryString.Count > 0)
     {
      userID     = Request.QueryString.Item("userID") + "";
      courseID   = Session("tutcurso");
      Lesson     = Request.QueryString.Item("Lesson") + "";
      Exercise   = Request.QueryString.Item("Exercise") + "";
      Puntuation = Request.QueryString.Item("Puntuation") + "";
      ID         = Request.QueryString.Item("ID") + "";
      fecha      = Request.QueryString.Item("Date") + "";

      
      var fliePath,oConn,oRec;
           
      filePath = Application("dataPath"); 
      oConn    = Server.CreateObject("ADODB.Connection");
      oRec     = Server.CreateObject("ADODB.Recordset");
      oRec.CursorLocation = 3;
      oRec.CursorType     = 3;

      
      oConn.Open( filePath );
         
      var  resp = ""
              oRec.Open("select * from Evaluaciones_Pendientes where (ID="+ ID +")",oConn,3,3);          
              if (!oRec.EOF) {
                resp =      oRec.Fields.Item("response").value;
              }
              oRec.Close(); 

            
              //Adiciona la evaluacion a la Base de Datos.......
              oRec.Open("select * from Evaluaciones",oConn,3,3);          
              oRec.AddNew();
              oRec.Fields.Item("User").value        = userID;
              oRec.Fields.Item("Course").value      = courseID;
              oRec.Fields.Item("Lesson").value      = Lesson;           
              oRec.Fields.Item("Exercise").value    = Exercise;
              oRec.Fields.Item("Puntuation").value  = Puntuation;
              oRec.Fields.Item("Date").value        = fecha;
              oRec.Fields.Item("response").value = resp;
              oRec.Update(); 
                    
              //Elimina la evaluacion pendiente...........
              oConn.Execute("delete  from Evaluaciones_Pendientes where (ID="+ ID +")");
               
      oConn.Close();
      
      Response.Redirect('exerlist.asp?uid=' + uid + "&courseOwner=" + Request.QueryString.Item("courseOwner"));
%>
<SCRIPT LANGUAGE=javascript>
<!--
function muOnclick()
 {
   location = 'exerlist.asp?uid=<%=uid%>' ;
 }
 
//-->
</SCRIPT>

 <center><b>La calificación fue registrada satisfactoriamente.</b></center>
 <center>    <INPUT align=center id=Buton name=Buton type=Button value="Regresar"  onclick=muOnclick()></center>
<%        
     } 
%> 


 
<%

    }
  else
    {  
      Response.Redirect("../errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);     
    }

  }
 else
   Response.Redirect("../errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
  //Response.Write(Request.QueryString.Item("uid"));
%>
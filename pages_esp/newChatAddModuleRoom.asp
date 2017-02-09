<%@ Language=JScript %>
<!-- #include file='../js/user.inc' -->
<%
  Response.Expires = -1;
  
  uid = -1;
  if (Request.QueryString.Item("uid").count != 0){
    uid = Request.QueryString.Item("uid");
  }
  
  if (uid == Session("uid"))
    {    

      //Chequeo que el usuario actual pueda entrar
      if (Session("valid")){
        idcursoactual = -1;
        if (Request.QueryString.Item("course").count != 0){
          idcursoactual = Request.QueryString.Item("course");
        }
  
        nombrecursoactual = -1;
        if (Request.QueryString.Item("courseName").count != 0){
          nombrecursoactual = Request.QueryString.Item("courseName");
        }
  
      
        var  filePath = Application("dataPath");   
        var  oConn    = Server.CreateObject("ADODB.Connection");
        var  oRec     = Server.CreateObject("ADODB.Recordset");
        oConn.Open(filePath);

        if ((Request.Form("roomname").Count != 0) && (Request.Form("roomname") != "") &&
            (Request.Form("roomobj").Count != 0 ) && (Request.Form("roomobj")  != "")){
          oRec.Open("SELECT * From Salas", oConn, 3, 3);
          oRec.AddNew();
          var sendedtexto = Request.Form("roomname") + "";
          re = /</g;             
          sendedtexto = sendedtexto.replace(re,"&lt;");    
          re = />/g;             
          sendedtexto = sendedtexto.replace(re,"&gt;");    
          oRec.Fields.Item("Nombre").value           = sendedtexto;
       
          sendedtexto = Request.Form("roomobj") + "";
          re = /</g;             
          sendedtexto = sendedtexto.replace(re,"&lt;");    
          re = />/g;             
          sendedtexto = sendedtexto.replace(re,"&gt;");    
          oRec.Fields.Item("Objetivo").value         = sendedtexto;
          oRec.Fields.Item("Publica").value          = (Request.Form("roomtype") == "on");
          oRec.Fields.Item("Moderador").value        = Session("UserID");
          oRec.Update();
          oRec.MoveLast();
          salainsertada = oRec.Fields.Item("ID").value
          oRec.Close(); 
  
          oRec.Open("SELECT * From SalasCurso", oConn, 3, 3);
          oRec.AddNew();
          oRec.Fields.Item("Sala").value = salainsertada;
          oRec.Fields.Item("Curso").value = idcursoactual;
          oRec.Update();
          oRec.Close(); 
      
          oConn.Close();
          Response.Redirect("newChatModuleRoomList.asp?uid=" + uid + "&course=" + idcursoactual + "&courseName=" + nombrecursoactual );
        };  
%>
<HTML>
<HEAD>
  <title>Agregar sala de charlas al módulo actual</title>
    <link rel="stylesheet" href="../css/<%=Session("skin")%>/EvalResult.css" type="text/css">
    <link rel="stylesheet" href="../css/<%=Session("skin")%>/Chat.css" type="text/css">
  <script type="text/javascript" language="javascript" src="../js/Windows.js"></script>
</HEAD>

<body text="#000000" style="overflow-y:auto">
<table width="100%" border="0" cellspacing="0" cellpadding="2">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td><img src="../images/<%=Session("skin")%>/ChatIMG.gif" width="80" height="54"></td>
          <td class="HeaderTable">Agregar nueva sala al módulo</td>
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
  <table width="50%"  align=center>
    <tr>
      <td>
        <form id=newsala name=newsala  action="newChatAddModuleRoom.asp?uid=<%=uid%>&course=<%=idcursoactual%>&courseName=<%=nombrecursoactual%>" method=post>
          <table width="100%" class="BorderedTable">
            <tr bordercolor=gold>
              <td width="1%" valign=top><br><td>
              <td width="99%"><br></td>
            </tr>
            <tr bordercolor=gold>
              <td width="1%" valign=top><b>Nombre:</b><td>
              <td width="99%"><input id=roomname name=roomname maxlength=255 class="ChatEdit" ><br><br></td>
            </tr>
            <tr   bordercolor=gold>
              <td width="1%"  valign=top><b>Objetivo:</b><td>
              <td width="99%"><input id=roomobj name=roomobj maxlength=255 class="ChatEdit" ><br><br></td>
            </tr>
            <tr   bordercolor=gold>
              <td width="1%"><input id=roomtype name=roomtype type=checkbox LANGUAGE=javascript><b>Pública</b><td>
              <td width="99%"><input class="TButton" type=Submit id="Submit" name="Submit" value="Aceptar"></td>
            </tr>
          </table> 
        </form>
      </td>
    </tr>    
  </table>   
</BODY>
</HTML>
<%
    }
    else{
      Response.Redirect("errorpage.asp?tipo=Error&short=El usuario no tiene suficientes privilegios." + "&desc=" + INVALID_CHARACTERS_TEXT);
    }
  }  
  else{
    Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
  }  
%>

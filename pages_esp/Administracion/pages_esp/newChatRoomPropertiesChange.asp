<%@ Language=JScript %>
<%
 Response.Expires = -1;
%>
<!-- #include file='../js/user.inc' -->
<!-- #include file='../js/chat.inc' -->
<%
  sala = -1;
  if (Request.QueryString.Item("salaid").count != 0){
    sala = Request.QueryString.Item("salaid");
  }
 
  var  filePath = Application("dataPath");   
  var  oConn    = Server.CreateObject("ADODB.Connection");
  var  oRec     = Server.CreateObject("ADODB.Recordset");
  
  oConn.Open(filePath);
  
  //Verifico que la sala sea valida
  oRec.Open("SELECT * FROM Salas WHERE (ID = " + sala + ")", oConn, 3, 3);    
  if (oRec.EOF == false){
    salanombre = oRec.Fields.Item("Nombre").value;
    salaobjetivo = oRec.Fields.Item("Objetivo").value;
    salapublica = oRec.Fields.Item("Publica").value;
    salamoderador  = oRec.Fields.Item("Moderador").value;
    oRec.Close();

    //Verifico si el usuario actual puede entrar a la sala
    if ((Session("userID") == salamoderador) || (Session("PermissionType") == ADMINISTRATOR)){

      //Si llegan cambios los actualizo
      if ((Request.Form("nombre").Count != 0) && (Request.Form("nombre") != "")){
        //Elimino codigos scripts maliciosos del mensaje
        var sendedtexto = Request.Form("nombre") + "";
        re = /</g;             
        sendedtexto = sendedtexto.replace(re,"&lt;");    
        re = />/g;             
        sendedtexto = sendedtexto.replace(re,"&gt;");  
        
        sendedobj = Request.Form("objetivo") + "";
        re = /</g;             
        sendedobj = sendedobj.replace(re,"&lt;");    
        re = />/g;             
        sendedobj = sendedobj.replace(re,"&gt;");    
          
      
        if (Application("SQL") == false) {
          oRec.Open("UPDATE Salas SET Nombre = '" + sendedtexto + "', Objetivo = '" + sendedobj + "', Publica = " + (Request.Form("publica") == "on") + "  WHERE (ID = " + sala + ")", oConn, 3, 3);
        }
        else{
          if (Request.Form("publica") == "on"){
            oRec.Open("UPDATE Salas SET Nombre = '" + sendedtexto + "', Objetivo = '" + sendedobj + "', Publica = 1  WHERE (ID = " + sala + ")", oConn, 3, 3);
          }
          else{
            oRec.Open("UPDATE Salas SET Nombre = '" + sendedtexto + "', Objetivo = '" + sendedobj + "', Publica = 0  WHERE (ID = " + sala + ")", oConn, 3, 3);
          }      
        }  
        salanombre = Request.Form("nombre") + "";
        salaobjetivo = Request.Form("objetivo") + "";
        salapublica = (Request.Form("publica") == "on");
      }
      

%>
<HTML>
<HEAD>
<title>Sala de charlas</title>
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Chat.css" type="text/css">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/EvalResult.css" type="text/css">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
  function Actualiza(x){
    if (x == 1){
      window.parent.frames(1).location = "newChatAddURE.asp?salaid=<%=sala%>&pn=0";
      window.parent.frames(2).location = "newChatDeleteURE.asp?salaid=<%=sala%>&pn=0";
    }
    else{
      window.parent.frames(1).location = "newChatAddURS.asp?salaid=<%=sala%>&pn=0";
      window.parent.frames(2).location = "newChatDeleteURS.asp?salaid=<%=sala%>&pn=0";
    }
  }

//-->
</SCRIPT>
</HEAD>

<BODY class="Toolbar">
<table width="100%" border="0" cellspacing="0" cellpadding="2">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td><img src="../images/<%=Session("skin")%>/ChatIMG.gif" width="80" height="54"></td>
          <td class="HeaderTable">Administrar la sala</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width=100% align=left cellpadding="2" cellspacing="0">
  <tr>
    <td colspan=2 align=middle valign=top>
      <form name=saladata id=saladata method=post action="newChatRoomPropertiesChange.asp?salaid=<%=sala%>" >   
        <table align=center width="100%" cellpadding="0" cellspacing="2">
          <tr>
            <td width="1%">
              Sala:
            </td>
            <td width="99%">
              <input id=nombre name=nombre maxlength=255 class="ChatEdit" value="<%=salanombre%>">
            </td>
          </tr>
          <tr> 
            <td width="1%">
              Objetivo:
            </td>
            <td width="99%" valign=center> 
              <input id=objetivo name=objetivo maxlength=255 class="ChatEdit" value="<%=salaobjetivo%>">
            </td>
          </tr>      
          <tr>
            <td>
              <input id=publica name=publica type=checkbox <%if (salapublica){%>CHECKED<%}else{}%>>Pública
            </td>  
            <td> 
              <input class="TButton" type="Submit" id="Submit" name="Submit" value="Confirmar cambios"">
            </td>      
          </tr>
          <tr>
            <td colspan=2  align=center>
              <hr>
              <table align=center width="100%" cellpadding="0" cellspacing="2">
                <tr>
                  <td align=center>
                    <a href="javascript:Actualiza(1)" class="ToolLink">
                      Editar usuarios que pueden entrar
                    </a>
                  </td>
                  <td align=center>
                    <a href="javascript:Actualiza(2)" class="ToolLink">
                      Editar usuarios que pueden hablar
                    </a>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
      </form>  
    </td>
  </tr>
</table>  
</BODY>
</HTML>
<%  }
    else{
      oRec.Close();
      oConn.Close();
      Response.Redirect("errorpage.asp?tipo=Error&short=Acceso denegado." + "&desc=El usuario actual no tiene suficientes privilegios para efectuar la operación");    
    }//Puede o no puede administrar la sala
  }
  else{
    oRec.Close();
    oConn.Close();
    Response.Redirect("errorpage.asp?tipo=Error&short=Error con la sala seleccionada." + "&desc=" + INVALID_ROOM);
  }//Sala valida o no
  oConn.Close();
%>

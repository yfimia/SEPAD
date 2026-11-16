<%@ Language=JScript %>
<%
 Response.Expires = -1;
%>
<!-- #include file='../js/user.inc' -->
<!-- #include file='../js/chat.inc' -->
<%
  var sala = -1;
  if (Request.QueryString.Item("salaid").count != 0){
    sala = Request.QueryString.Item("salaid");
  }

  var usrid = -1;
  if (Request.QueryString.Item("userid").count != 0){
    usrid = Request.QueryString.Item("userid");
  }

  function deg2rad(x){
    return x * Math.PI / 180;  
  }
                   
  function dec2hex(x){
    gg = x.toString(16);
    if (gg.length == 1) {
      gg = "0" + gg;
    }
    return gg;
  }

 
  var  filePath = Application("dataPath");   
  var  oConn    = Server.CreateObject("ADODB.Connection");
  var  oRec     = Server.CreateObject("ADODB.Recordset");
  oRec.CursorLocation = 3;
  oRec.CursorType     = 3;
  
  oConn.Open(filePath);

  //Verifico si la sala es valida
  oRec.Open("SELECT * FROM Salas WHERE (ID = " + sala + ")", oConn, 3, 3);    
  salapublica = false;
  if (oRec.EOF == false){
    nameofsala = oRec.Fields.Item("Nombre").value;
    salapublica = oRec.Fields.Item("Publica").value;
    salamoderador = oRec.Fields.Item("Moderador").value;
    oRec.Close();

  
    //Si no llega configuracion por ser la primera vez
    if ((Request.Form("msgcolor").Count == 0) || (Request.Form("msgcolor") == "")){
      //Chequeo si existes configuracion
      oRec.Open("SELECT * From UserConfig Where (Usuario = " + Session("UserID") + ")", oConn, 3, 3);    
      if (oRec.EOF == false){
        //Si existe la cargo
        usercolor = oRec.Fields.Item("Color").value + "";
        usernegrita = oRec.Fields.Item("Negrita").value;
        usercursiva = oRec.Fields.Item("Cursiva").value;
        usersubrayada = oRec.Fields.Item("Subrayado").value;
        oRec.Close();
      }
      else{
        //Pongo los valores por defecto
        usercolor = "#000000";
        usernegrita = false;
        usercursiva = false;
        usersubrayada = false;
        oRec.Close();      
      }  
    }
    else{
      usercolor = Request.Form("msgcolor");
      usernegrita = (Request.Form("msgnegrita") == "on"); 
      usercursiva = (Request.Form("msgcursiva") == "on"); 
      usersubrayada = (Request.Form("msgsubrayada") == "on");
    }            
    
    //Actualizo configuraciones del envio de mensaje
    oRec.Open("SELECT * From UserConfig Where (Usuario = " + Session("UserID") + ")", oConn, 3, 3);    
    existeconfig = false;
    existeconfig = (oRec.EOF == false);
    oRec.Close();
    if (existeconfig){
      //Actualizo la configuracion con los valores actuales
      oRec.Open("UPDATE UserConfig SET color = '" + Request.Form("msgcolor") + "'" + 
                " , negrita = " + ((Request.Form("msgnegrita") == "on")?Application("dtrue"):Application("dfalse")) + 
                " , cursiva = " + ((Request.Form("msgcursiva") == "on")?Application("dtrue"):Application("dfalse")) + 
                " , subrayado = " + ((Request.Form("msgsubrayado") == "on")?Application("dtrue"):Application("dfalse")) + 
                " WHERE (Usuario = " + Session("UserID") + ")", oConn, 3, 3);
    }
    else{
      //Como la configuracion no existe la creo con los valores actuales         
      oRec.Open("SELECT * From UserConfig", oConn, 3, 3);    
      oRec.AddNew();
      oRec.Fields.Item("Color").value            = Request.Form("msgcolor");
      oRec.Fields.Item("Negrita").value          = (Request.Form("msgnegrita") == "on");
      oRec.Fields.Item("Cursiva").value          = (Request.Form("msgcursiva") == "on");
      oRec.Fields.Item("Subrayado").value        = (Request.Form("msgsubrayada") == "on");
      oRec.Fields.Item("Usuario").value          = Session("UserID");
      oRec.Update();
      oRec.Close(); 
    }

  
    //Verifico si el usuario actual puede entrar a la sala
    oRec.Open("SELECT * FROM ChatCanEnter WHERE (Sala = " + sala + ") and (Usuario = " + Session("userID") + ")", oConn, 3, 3);     
    if ((salapublica) || (salamoderador == Session("userID")) || (oRec.EOF == false) || (Session("PermissionType") == ADMINISTRATOR)){
      oRec.Close();

      //Verifico si el usuario actual puede enviar mensajes
      usercanspeak = false;
      oRec.Open("SELECT * FROM ChatCanSpeak WHERE (Sala = " + sala + ") and (Usuario = " + Session("userID") + ")", oConn, 3, 3);     
      if ((salapublica) || (salamoderador == Session("userID")) || (oRec.EOF == false) || (Session("PermissionType") == ADMINISTRATOR)){
        oRec.Close();
        usercanspeak = true;
        
        //Si el msg que se desea enviar no es vacio lo mando al canal publico de la sala 
        if ((Request.Form("msgtext").Count != 0 ) && (Request.Form("msgtext") + "" != "")){
          
          //Elimino codigos scripts maliciosos del mensaje
          var sendedtexto = Request.Form("msgtext") + "";
          re = /</g;             
          sendedtexto = sendedtexto.replace(re,"&lt;");    
          re = />/g;             
          sendedtexto = sendedtexto.replace(re,"&gt;");    
          
          hayalguien = false;
          for (k = 1; k <= Request.Form("cantidad"); k++){
            if (Request.Form("userto" + k) != ''){
              hayalguien = true;
            }
          }    
             
          if (hayalguien){
            oRec.Open("SELECT * From ChatMenssages", oConn, 3, 3);    
            oRec.AddNew();
            oRec.Fields.Item("Texto").value            = sendedtexto;
            oRec.Fields.Item("Color").value            = Request.Form("msgcolor");
            oRec.Fields.Item("Negrita").value          = (Request.Form("msgnegrita") == "on");
            oRec.Fields.Item("Cursiva").value          = (Request.Form("msgcursiva") == "on");
            oRec.Fields.Item("Subrayada").value        = (Request.Form("msgsubrayada") == "on");
            oRec.Fields.Item("Sala").value             = sala;
            oRec.Fields.Item("UserOrigen").value       = Session("UserID");
            oRec.Fields.Item("Publico").value          = false;
            oRec.Fields.Item("Carita").value           = Request.Form("msgemotion");
            oRec.Update();
            oRec.MoveLast();
            usermsgid = oRec.Fields.Item("ID").value;
            oRec.Close();   
            
            oRec.Open("SELECT * From ChatPrivateMsg", oConn, 3, 3);    
            for (k = 1; k <= Request.Form("cantidad"); k++){
              if (Request.Form("userto" + k) != ''){
                oRec.AddNew();
                oRec.Fields.Item("MsgSource").value        = usermsgid;
                oRec.Fields.Item("UserDestino").value      = Request.Form("userto" + k);
              }  
            }  
            oRec.Update();
            oRec.Close();   
          }  
        } 
      }  
      else{
        oRec.Close();
      };  
%>
<HTML>
<HEAD>
<title>Sala de charlas privadas</title>
<link rel="stylesheet" href="../css/<%=Session("skin")%>/chat.css" type="text/css">

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function msgkeypress(){  
  if (event.keyCode == 13){ 
    Konclick(); 
  }
}

function Konclick(){
  if (msgsender.msgtext.value != ""){
    for (l = 1; l <= msgsender.cantidad.value; l++){
      if (document.all("selusr" + l).checked){
        document.all("userto" + l).value = document.all("selusr" + l).title;
      }  
    }
  }
  else{
    alert("Mensanje no válido.");
  }
}

function button1_onclick(){
  if (msgsender.msgtext.value != ""){
    for (l = 1; l <= msgsender.cantidad.value; l++){
      if (document.all("selusr" + l).checked){
        document.all("userto" + l).value = document.all("selusr" + l).title;
      }  
    }
    msgsender.submit();
  }
  else{
    alert("Mensanje no válido.");
  }
}

function color_onclick(xx) {
  boxcolor.bgColor = xx
  msgsender.msgcolor.value = xx;
}

function emo_onclick(xx,ii) {
  msgsender.msgemotion.value = ii;
  emoselect.src = "../images/<%=Session("skin")%>/chat/emotions/" + xx;
}

function emoreset_onclick() {
  emoselect.src = "../images/<%=Session("skin")%>/chat/emotions/none.gif";
}

function clearCheck(cant){
  for (i = 1; i <= cant; i++ ){
    document.all("selusr" + i).checked = false;
  }
}

function selectCheck(cant){
  for (i = 1; i <= cant; i++ ){
    document.all("selusr" + i).checked = true;
  }
}

//-->
</SCRIPT>
</HEAD>

<BODY class="Toolbar">
<%
      if (usercanspeak){
%>      
<table>
  <tr>
    <td width=25%>
      <table class="BorderedTable">
        <tr>
          <td colspan=2  class="ServicesTable">
            <b>Lista de usuarios</b>
          </td>
        </tr>
        <%
          oRec.Open("SELECT ChatOnlineUsers.Usuario,Usuarios.[Name], ChatOnlineUsers.Sala " + 
                    "FROM Usuarios INNER JOIN ChatOnlineUsers ON Usuarios.ID = ChatOnlineUsers.Usuario " +
                    "WHERE (ChatOnlineUsers.Usuario <> " + Session("userID") + ") and (ChatOnlineUsers.Sala = " + sala + ")" , oConn, 3, 3);    
          k = 0;
          hiddenstr = "<input id=cantidad name=cantidad type=hidden value=" + oRec.RecordCount + ">";
          while (oRec.EOF == false){
            k++;
        %>
        <tr>
          <td colspan=2 class="">
            <%if ( (usrid != -1) && (oRec.Fields.Item("Usuario").value == usrid)){%>
            <input id=selusr<%=k%> name=selusr<%=k%> type=checkbox CHECKED title=<%=oRec.Fields.Item("Usuario").value%>>
            <%}
              else{
            %>  
            <input id=selusr<%=k%> name=selusr<%=k%> type=checkbox  title=<%=oRec.Fields.Item("Usuario").value%>>
            <%  
              }
              hiddenstr = hiddenstr + "<input id=userto" + k + " name=userto" + k + " type=hidden value=''>";
            %>
            [<b><%=oRec.Fields.Item("name").value%></b>]
          </td>
        </tr>
        <% 
            oRec.Move(1);
          }
          oRec.Close();
        %>
        <tr>
          <td>
            <input class="TButton" id=Clear name=Clear type=button value="Limpiar" onClick="clearCheck(<%=k%>)">
          </td>
          <td>
            <input class="TButton" id=Selecter name=Selecter type=button value="Seleccionar todos" onClick="selectCheck(<%=k%>)">
          </td>
        </tr>
      </table>
    </td>
    <td>
<table width=100% align=left cellpadding="2" cellspacing="0">
  <tr>
    <td colspan=2 align=middle valign=top>
      <form name=msgsender id=msgsender method=post action="newChatPrivateMessanger.asp?salaid=<%=sala%><% if (usrid != -1){%>&userid=<%=usrid%><%}%>" >   
        <table align=center width="100%" cellpadding="0" cellspacing="0">
          <tr>
            <td>
              <input id=msgnegrita name=msgnegrita type=checkbox <%if (usernegrita){%>CHECKED<%}else{}%> >Negrita
              <input id=msgcursiva name=msgcursiva type=checkbox <%if (usercursiva){%>CHECKED<%}else{}%>>Cursiva
              <input id=msgsubrayada name=msgsubrayada type=checkbox <%if (usersubrayada){%>CHECKED<%}else{}%>>Subrayada
            </td>
            <td>
              <%=hiddenstr%>
              <table  border=0 cellpadding=0 cellspacing=0>
                <tr>
                  <%
                    var TextColors = new Array();
   	                // Define los colores a picar solo para JavaScript1.1+ browsers 
		            TextColors[0] = '#000000';
		            TextColors[1] = '#ffffff';
		            p = 2;
		            for(i = 0; i < 360; i += 6){
		              Math.deg
			          r = Math.ceil(126 * (Math.cos(deg2rad(i)) + 1));
			          g = Math.ceil(126 * (Math.cos(deg2rad(i + 240)) + 1));
			          b = Math.ceil(126 * (Math.cos(deg2rad(i + 120)) + 1));
			          if (!(r > 128 && g < 128 && b < 128)){
				        rr = r.toString(16);
				        TextColors[p] = '#' + dec2hex(r) + dec2hex(g) + dec2hex(b);
				        p++; 
			          }//if
  		            }//for                    
                      
                    for (i = 0; i < p; i++){
                  %>
                  <td id=c<%=i%> name=c<%=i%> bgcolor=<%=TextColors[i]%> width=1px  height=10px LANGUAGE=javascript onclick="color_onclick('<%=TextColors[i]%>')">
                    &nbsp;
                  </td>
                  <%
                    } //for
                  %>
                  <td>
                    &nbsp;&nbsp;&nbsp;&nbsp;
                  </td>
                  <td id=boxcolor name=boxcolor bgcolor=<%=usercolor%>>
                    &nbsp;&nbsp;&nbsp;&nbsp;
                  </td>
                </tr> 
              </table>
            </td>
          </tr>
          
          <tr> 
            <td width="99%" valign=center colspan=2> 
              <table border=0 cellpadding=0 cellspacing=3>
                <tr>
                  <td width="99%">
                    <input id=msgtext name=msgtext maxlength=255 class="ChatEdit" >
                  <td>  
                  <td width="1%">
                    <input class="TButton" type="button" id="Submit" name="Submit" value=" Enviar" onClick="button1_onclick()">
                    <input id=msgcolor name=msgcolor type=hidden value=<%=usercolor%>>
                    <input id=msgemotion name=msgemotion type=hidden value="0">
                  </td>  
                </tr>  
              </table>  
            </td>
          </tr>
        </table>
      </form>  
    </td>
  </tr>
  <tr  valign=top>
    <td  valign=top>
      <table border=0>
        <tr  valign=top>
          <%
            oRec.Open("SELECT * From Caritas", oConn, 3, 3);    
            i = 0; 
            while (oRec.EOF == false){
              i++;
          %> 
          <td  valign=top>
            <img id=emo<%=i%> name=emo<%=i%> alt="<%=oRec.Fields.Item("texto").value%>" border=0 src="../images/<%=Session("skin")%>/chat/emotions/<%=oRec.Fields.Item("imagen").value%>" title="<%=oRec.Fields.Item("texto").value%>" LANGUAGE=javascript onclick="emo_onclick('<%=oRec.Fields.Item("imagen").value%>',<%=oRec.Fields.Item("ID").value%>)">
          </td>    
          <%
              oRec.Move(1);
            }
            oRec.Close();
          %>
          <td valign=top>
            <img src="../images/<%=Session("skin")%>/chat/emotions/arrow.gif">
          </td>
          <td valign=top>
            <img border=1 id=emoselect name=emoselect src="../images/<%=Session("skin")%>/chat/emotions/none.gif" title="Presiona click sobre mi para no poner ninguna" LANGUAGE=javascript onclick="emoreset_onclick()">
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>  
    </td>
  </tr>  
</table>
<script language=jscript>
  msgsender.msgtext.focus();
</script>
 
<%
  }
  else{
%>
  <hr>
  <marquee height=15><h3>En esta sala usted solo puede leer los mensajes.</h3></marquee>
  <hr>
<%
  }
%>
</BODY>
</HTML>
<%  }
    else{
      oRec.Close();
      oConn.Close();
      Response.Redirect("errorpage.asp?tipo=Error&short=El usuario no tiene suficientes privilegios." + "&desc=" + INVALID_ROOM_USER);
    }//Puede o no puede entrar a la sala
  }
  else{
    oRec.Close();
    oConn.Close();
    Response.Redirect("errorpage.asp?tipo=Error&short=Error con la sala seleccionada." + "&desc=" + INVALID_ROOM);
  }//Sala valida o no
  oConn.Close();
%>

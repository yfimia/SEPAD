<%@ Language=JScript %>
<%
  Response.Expires = -1;
%>
<!-- #include file='../js/user.inc' -->
<%
  sala = -1;
  if (Request.QueryString.Item("salaid").count != 0){
    sala = Request.QueryString.Item("salaid");
  }

  pagenumber = 0;
  if (Request.QueryString.Item("pn").count != 0){
    pagenumber = Request.QueryString.Item("pn");
  }

  DoMyAction = "";
  
  var  filePath = Application("dataPath");   
  var  oConn    = Server.CreateObject("ADODB.Connection");
  var  oRec     = Server.CreateObject("ADODB.Recordset");
  var  oRec1     = Server.CreateObject("ADODB.Recordset");
  
  oConn.Open(filePath);
  
  //Verifico que la sala sea valida
  oRec.Open("SELECT * FROM Salas WHERE (ID = " + sala + ")", oConn, 3, 3);    
  if (oRec.EOF == false){
    salamoderador = oRec.Fields.Item("Moderador").value;
    oRec.Close();

    //Verifico si el usuario actual puede entrar a la sala
    if ((Session("userID") == salamoderador) || (Session("PermissionType") == ADMINISTRATOR)){
  
  
      if (Request.Form("cantidad").Count != 0){
        oRec.Open("SELECT * From ChatCanSpeak", oConn, 3, 3);    
        for (k = 1; k <= Request.Form("cantidad"); k++){
          if ((Request.Form("nombre" + k).Count != 0) && (Request.Form("nombre" + k) != "")){
            oRec1.Open("SELECT * From ChatCanSpeak WHERE (Usuario = " + Request.Form("nombre" + k) + ") and (sala = " + sala + ")", oConn, 3, 3);
            if (oRec1.EOF == true){
              oRec.AddNew();
              oRec.Fields.Item("Usuario").value        = Request.Form("nombre" + k);
              oRec.Fields.Item("Sala").value           = sala;
            }  
            oRec1.Close();
          }  
          DoMyAction = "<script language=jscript>window.parent.frames(2).location = 'newChatDeleteURS.asp?salaid=" + sala + "&pn=0'; </script>";  
        }
        oRec.Update();
        oRec.Close(); 
      }  
  
  
      //oRec.Open("SELECT * FROM usuarios, WHERE (ID <> 580) ORDER BY Name", oConn, 3, 3);
      oRec.Open("SELECT Usuarios.* FROM Usuarios,ChatCanEnter WHERE (Usuarios.ID = ChatCanEnter.Usuario) and (ChatCanEnter.Sala = " + sala + ") ORDER BY Name", oConn, 3, 3);
      
%>
<HEAD>
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
  <link rel="stylesheet" href="../css/<%=Session("skin")%>/Chat.css" type="text/css">
  <script src="../js/CheckBoxes.js" language="JavaScript"></script>
  <script language=jscript>
    function Adiciona(){
      for (l = 1; l <= usuarios.cantidad.value; l++){
        if (!document.all("Check" + l).checked){
          document.all("nombre" + l).value = "";
        }
      }
      usuarios.submit();
    }
  </script>
  <%=DoMyAction%>
</HEAD>

<body text="#000000" style="overflow-y:auto">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td colspan=3 class="ToolBar" align="center">
        <b>Listado de usuarios</b>
      </td>
    </tr>
    <tr> 
      <td> 
        <form id="usuarios" name="usuarios" action="newChatAddURS.asp?salaid=<%=sala%>&pn=<%=pagenumber%>" method=post>
          <table width="100%" border="0" cellspacing="1" cellpadding="0">
            <tr> 
              <td width="5%" class="ToolBar" align="center"></td>
              <td width="20%" class="ToolBar" align="center">
                <b>Identificador</b>
              </td>
              <td width="75%" class="ToolBar" align="center">
                <b>Nombre y apellidos</b>
              </td>
            </tr>
<%    
      
      var clase = "MessageTR1";
      
      creg = oRec.RecordCount;
      
      if (pagenumber * SHOW_CANT < creg){
        oRec.Move(pagenumber * SHOW_CANT); 
      }
      else{
        pagenumber = 0;
      }
      
      var i = 0;
      while ((oRec.EOF == false) && (i < SHOW_CANT)){ 
        i = i + 1;
        if (clase == "MessageTR1"){clase = "MessageTR";} else {clase = "MessageTR1";};
%>
            <tr> 
              <td width="5%" class="<%=clase%>">
                <input type="checkbox" id=Check<%=i%> name=check<%=i%> >
                <input type="hidden" id=nombre<%=i%> name=nombre<%=i%> value="<%=oRec.Fields.Item("ID").value%>">
              </td>
              <td width="20%" class="<%=clase%>">
                <%=oRec.Fields.Item("name").value%>
              </td>
              <td width="75%" class="<%=clase%>">
                <%=oRec.Fields.Item("FullName").value%>
              </td>
            </tr> 
<%
        oRec.Move(1);
      }  
      oRec.Close();
%>        
          </table>
          <input type="hidden" id=cantidad name=cantidad value="<%=i%>">
          <table border="0" cellspacing="0" cellpadding="0" class="ToolBar" width="100%">
            <tr>
              <td>
                <table border="0" cellspacing="0" cellpadding="4" class="ToolBar">
                  <tr> 
                    <td nowrap>
                      <a href="javascript:Adiciona()" class="ToolLink">
                        &nbsp;Adicionar&nbsp;
                      </a>
                    </td>
                    <td>|</td>
                    <td>
                      <a href="#" class="ToolLink" onClick="CheckAll(GetObjectCollection( document, 'FORM'))">
                        &nbsp;Seleccionar&nbsp;Todo&nbsp;
                      </a>
                    </td>
                    <td>|</td>
                    <td>
                      <a href="newChatAddURS.asp?salaid=<%=sala%>&pn=<%=pagenumber + 1%>" class="ToolLink">
                        &nbsp;Próximos&nbsp;<%=SHOW_CANT%>&nbsp;
                      </a>
                    </td>
                  </tr>
                </table>
                <table border="0" cellspacing="0" cellpadding="0" class="ToolBar" width="100%">
                  <tr>
                    <td colspan=5  align=center>
                      <%
                        i = 0;
                        while (i * SHOW_CANT < creg){
                      %>    
                      <a class="numbers" href="newChatAddURS.asp?salaid=<%=sala%>&pn=<%=i%>">
                        <b><%=i + 1%>&nbsp;</b>
                      </a>  
                      <%
                          i = i + 1;
                        }
                      %>
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

<%@ Language=JScript %>
<%Response.Expires = -1;%>
<!-- #include file='../js/adoLibrary.inc' -->
<!-- #include file='../js/user.inc' -->

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

%>

<%
if (Session("PermissionType") == ADMINISTRATOR)
    {
      if ((Request.Form("Mname").Count == 0) || (Request.Form("Mname") == ""))
        {
%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/Main.css" type="text/css">
<link rel="stylesheet" href="../css/<%=Session("skin")%>/EvalResult.css" type="text/css">
<script src="../js/windows.js" language="JavaScript"></script>
</head>
<script language=jscript>
  function User(id, name, fullName, email) {
    this.id = id;
    this.name = name;
    this.fullName = fullName;
    this.email = email;
  }
  
  function Group(id, name, desc) {
    this.id = id;
    this.name = name;
    this.desc = desc;
  }
  
  function findUser() {
    var user = new Array();
    user = showModalDialog("findUser.asp?uid=<%=uid%>", "", "");
    if (user != null) {
      newgroup.Mcord.value = user[0].fullName + ' (' + user[0].name + ')';
      newgroup.Mcordid.value = user[0].id;
    }  
  }
  
  function findGroup() {
    var group = new Array();
    group = showModalDialog("findGroup.asp?uid=<%=uid%>", "", "");
    if (group != null) {
      newgroup.Mgrupo.value = group[0].name;
      newgroup.Mgrupoid.value = group[0].id;
    }  
  }

  function findClaustro() {
    var group = new Array();
    group = showModalDialog("findGroup.asp?uid=<%=uid%>", "", "");
    if (group != null) {
      newgroup.Mclaustro.value = group[0].name;
      newgroup.Mclaustroid.value = group[0].id;
    }  
  }
 
 var curText;
 function editHTML(textArea)  {
   curTex = textArea;
   abreVentana('',580,530,'../WordPad/wordpad.htm', 'no', 'no');
 }
</script>
<body bgcolor="#FFFFFF" text="#000000">
  <form name=newgroup id=newgroup action="agregaMod.asp?uid=<%=uid%>" method=post>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td class="HeaderTable"  align=center><b>Adicionar modalidad acad�mica</b></td>
        </tr>
      </table>
    </td>
  </tr>
</table>

    <table border="0" cellspacing="1" cellpadding="1" align="center">
      <tr>
        <td  align=center Class="MessageTR1">
          Nombre:
        </td>
        <td  Class="MessageTR1"align=center >
          <input name=Mname Id=Mname type=text maxlength=250 size=60 >
        </td>
      </tr>
      <tr>
        <td align=center   Class="MessageTR">
          Estado:
        </td>
        <td align=center   Class="MessageTR">
          <%=MOD_ACA_SELECT_STATE%>
        </td>
      </tr>
      <tr>
      
      <tr>
        <td align=center   Class="MessageTR">
          Coordinador:
        </td>
        <td align=center   Class="MessageTR">
          <input disabled name=Mcord Id=Mcord type=text maxlength=250 size=35 ><input name=Mname Id=Mname type=button value="Seleccionar..." onclick="findUser()">
          <input name=Mcordid Id=Mcordid type=hidden maxlength=250 size=35 >
        </td>
      </tr>

      <tr>
        <td align=center   Class="MessageTR">
          Claustro:
        </td>
        <td align=center   Class="MessageTR">
          <input name=Mclaustro Id=Mclaustro type=text maxlength=250 size=35 ><input name=Mname Id=Mname type=button value="Seleccionar..." onclick="findClaustro()">
          <input name=Mclaustroid Id=Mclaustroid type=hidden maxlength=250 size=35 >
          <P><input name=Mcrearclaustro Id=Mcrearclaustro type=checkbox  >Crear el grupo del claustro con el nombre especificado</P></td>
      </tr>

      <tr>
        <td align=center   Class="MessageTR">
          Grupo de alumnos:
        </td>
        <td align=center   Class="MessageTR">
          <input name=Mgrupo Id=Mgrupo type=text maxlength=250 size=35 ><input name=Mname Id=Mname type=button value="Seleccionar..." onclick="findGroup()">
          <input name=Mgrupoid Id=Mgrupoid type=hidden maxlength=250 size=35 >
          <P><input name=Mcreargrupo Id=Mcrearclaustro type=checkbox  >Crear el grupo alumnos con el nombre especificado</P></td>
      </tr>

      <tr>
        <td align=center   Class="MessageTR">
          Certificado:
        </td>
        <td align=center   Class="MessageTR">
          <input name=Mdiploma Id=Mdiploma type=text maxlength=250 size=35 >
        </td>
      </tr>


      
      <tr>
        <td align=center   Class="MessageTR">
          Objetivos Generales:
        </td>
        <td align=center   Class="MessageTR">
          <textarea name=Mobj Id=Mobj" rows="8" cols="58" maxlength=250></textarea>
        </td>
      </tr>

      <tr>
        <td align=center   Class="MessageTR">
          Programa:
        </td>
        <td align=center   Class="MessageTR">
          <textarea name=Mprograma Id=Mprograma" rows="8" cols="58" maxlength=250></textarea>
        </td>
      </tr>

      <tr>
        <td align=center   Class="MessageTR">
          Fechas y duraci�n:
        </td>
        <td align=center   Class="MessageTR">
          
          <textarea name=Mfechas Id=Mfechas" rows="8" cols="58" maxlength=250></textarea>
        </td>
      </tr>

      <tr>
        <td align=center   Class="MessageTR">
          Perfil del Egresado:
        </td>
        <td align=center   Class="MessageTR">
          <textarea name=Mfprevia Id=Mfprevia" rows="8" cols="58" maxlength=250></textarea>
        </td>
      </tr>

      <tr>
        <td align=center   Class="MessageTR">
          Sistema de Tutor�as:
        </td>
        <td align=center   Class="MessageTR">
          <textarea name=Mapp Id=Mapp" rows="8" cols="58" maxlength=250></textarea>
        </td>
      </tr>

      <tr>
        <td align=center   Class="MessageTR">
          Recursos:
        </td>
        <td align=center   Class="MessageTR">
          <textarea name=Mdoc Id=Mdoc" rows="8" cols="58" maxlength=250></textarea>
        </td>
      </tr>

      <tr>
        <td align=center   Class="MessageTR">
          Sistema de Evaluaci�n:
        </td>
        <td align=center   Class="MessageTR">
          <textarea name=Morg Id=Morg" rows="8" cols="58" maxlength=250></textarea>
        </td>
      </tr>

      <tr>
        <td align=center   Class="MessageTR">
          Coste:
        </td>
        <td align=center   Class="MessageTR">
          <textarea name=Mcoste Id=Mcoste" rows="8" cols="58" maxlength=250></textarea>
        </td>
      </tr>

      <tr>
        <td align=center   Class="MessageTR">
          Datos de la instituci�n:
        </td>
        <td align=center   Class="MessageTR">
          <textarea name=Minst Id=Minst" rows="8" cols="58" maxlength=250></textarea>
        </td>
      </tr>


      <tr>
        <td align=center   Class="MessageTR">
          Fundamentaci�n:
        </td>
        <td align=left   Class="MessageTR">
          <textarea name=MFundament Id=MFundament" rows="8" cols="58" maxlength=250></textarea>
        </td>
      </tr>

      <tr>
        <td align=center   Class="MessageTR">
          Sede:
        </td>
        <td align=left   Class="MessageTR">
          <textarea name=MSede Id=Msede" rows="8" cols="58" maxlength=250></textarea>
        </td>
      </tr>
      <tr>
        <td align=center   Class="MessageTR">
          Bibliograf�a General:
        </td>
        <td align=left   Class="MessageTR">
          <textarea name=MBibliog Id=MBibliog" rows="8" cols="58" maxlength=250></textarea>
        </td>
      </tr>

      
      <tr>
        <td colspan=2 align=center Class="Toolbar">
          <input type=submit value="Adicionar">
        </td>      
      </tr>
    </table>
  </form>
<%
        }
      else
        {
           var filePath = Application('dataPath');
           var oConn;
           var oRec;
 
           oConn = MakeConnection( oConn, filePath );
           Sql = "select * from Modulo where (Name = '" + Request.Form("Mname") + "')";

           oRec = Query( Sql, oRec, oConn  );	
          

          if (oRec.EOF == false)
            { 
              oConn.Close();   
              
              Response.Redirect("errorpage.asp?tipo=Error&short=" + MOD_NAME_EXIST_SHORT  + "&desc=" + MOD_NAME_EXIST_TEXT); 
            }  
          else
            {
              oConn.Close();   
              oConn.Errors.Clear();
              oConn.Open( filePath );
              oRec.Open("select * from Modulo",oConn,3,3);
              oRec.AddNew();
              if ((Request.Form("MName").Count != 0) && (Request.Form("Mname") + "" != ""))
                { 
                  oRec.Fields.Item("Name").Value = Request.Form("Mname") + "";
                }
              else Response.Redirect("errorpage.asp?tipo=Error&short=Nombre no v�lido&desc=El nombre de la modalidad introducido no es v�lido o se ha dejado en blanco.");

              oRec.Fields.Item("state").Value = Request.Form("Mstate") + "";

              if ((Request.Form("Mcordid").Count != 0) && (Request.Form("Mcordid") + "" != ""))
                { 
                  oRec.Fields.Item("cordinador").Value = Request.Form("Mcordid") + "";
                }
              else Response.Redirect("errorpage.asp?tipo=Error&short=Coordinador no v�lido&desc=El coordinador no ha sido seleccionado apropiadamente. En la p�gina de crear modalidad acad�mica para seleccionar el coordinador, presione el bot�n Seleccionar y esc�jalo de la lista.");

              if ((Request.Form("Mcreargrupo").Count > 0) && (Request.Form("Mcreargrupo") == 'on')) {
                if ((Request.Form("Mgrupo").Count > 0) && (Request.Form("Mgrupo") != '')) {
                  gid = AddGroup(Request.Form("Mgrupo") + "", "Grupo de alumnos de la modalidad acad�mica " + Request.Form("Mname"));
                  if (gid == -2) {
                     Response.Redirect("errorpage.asp?tipo=Error&short=" + GRUOP_NAME_EXIST_SHORT  + "&desc=" + GRUOP_NAME_EXIST_TEXT);
                  } 
                  else if (gid != -1) {
                    oRec.Fields.Item("grupo").Value = gid;
                  }     
                }
                else
                  Response.Redirect("../errorpage.asp?tipo=Error&short=Nombre no v�lido&desc=El nombre de grupo se ha dejado en blanco, o no es v�lido.");
                
              }
              else {
                if ((Request.Form("Mgrupoid").Count > 0) && (Request.Form("Mgrupoid") != '')) {
                  oRec.Fields.Item("grupo").Value = Request.Form("Mgrupoid") + "";
                }               	  
                else
                  Response.Redirect("../errorpage.asp?tipo=Error&short=Nombre no v�lido&desc=El nombre de grupo se ha dejado en blanco, o no es v�lido.");
                
              }
              
              if ((Request.Form("Mcrearclaustro").Count > 0) && (Request.Form("Mcrearclaustro") == 'on')) {
                if ((Request.Form("Mclaustro").Count > 0) && (Request.Form("Mclaustro") != '')) {
                  gid = AddGroup(Request.Form("Mclaustro") + "", "Grupo del claustro de la modalidad acad�mica " + Request.Form("Mname"));
                  if (gid == -2) {
                     Response.Redirect("errorpage.asp?tipo=Error&short=" + GRUOP_NAME_EXIST_SHORT  + "&desc=" + GRUOP_NAME_EXIST_TEXT);
                  } 
                  else if (gid != -1) {
                    oRec.Fields.Item("claustro").Value = gid;
                  }     
                }
                else
                  Response.Redirect("../errorpage.asp?tipo=Error&short=Nombre no v�lido&desc=El nombre de grupo se ha dejado en blanco, o no es v�lido.");
              }
              else {
                if ((Request.Form("Mclaustroid").Count > 0) && (Request.Form("Mclaustroid") != '')) {
                  oRec.Fields.Item("claustro").Value = Request.Form("Mclaustroid") + "";
                }               	  
                else
                  Response.Redirect("../errorpage.asp?tipo=Error&short=Nombre no v�lido&desc=El nombre de grupo se ha dejado en blanco, o no es v�lido.");
              }

              
              
              if ((Request.Form("Mfechas").Count != 0) && (Request.Form("Mfechas") + "" != ""))
                { 
                  oRec.Fields.Item("fechas").Value = Request.Form("Mfechas") + ""; 
                }
              else oRec.Fields.Item("fechas").Value = " ";   
              

              if ((Request.Form("Mdiploma").Count != 0) && (Request.Form("Mdiploma") + "" != ""))
                { 
                  oRec.Fields.Item("diploma").Value = Request.Form("Mdiploma") + ""; 
                }
              else oRec.Fields.Item("diploma").Value = " ";   
              
              if ((Request.Form("Mobj").Count != 0) && (Request.Form("Mobj") + "" != ""))
                { 
                  oRec.Fields.Item("objetivos").Value = Request.Form("Mobj") + ""; 
                }
              else oRec.Fields.Item("objetivos").Value = " ";   

              if ((Request.Form("Mprograma").Count != 0) && (Request.Form("Mprograma") + "" != ""))
                { 
                  oRec.Fields.Item("programa").Value = Request.Form("Mprograma") + ""; 
                }
              else oRec.Fields.Item("programa").Value = " ";   

              if ((Request.Form("Mfprevia").Count != 0) && (Request.Form("Mfprevia") + "" != ""))
                { 
                  oRec.Fields.Item("fprevia").Value = Request.Form("Mfprevia") + ""; 
                }
              else oRec.Fields.Item("fprevia").Value = " ";   

              if ((Request.Form("Mapp").Count != 0) && (Request.Form("Mapp") + "" != ""))
                { 
                  oRec.Fields.Item("app").Value = Request.Form("Mapp") + ""; 
                }
              else oRec.Fields.Item("app").Value = " ";   

              if ((Request.Form("Morg").Count != 0) && (Request.Form("Morg") + "" != ""))
                { 
                  oRec.Fields.Item("org").Value = Request.Form("Morg") + ""; 
                }
              else oRec.Fields.Item("org").Value = " ";   

              if ((Request.Form("Mcoste").Count != 0) && (Request.Form("Mcoste") + "" != ""))
                { 
                  oRec.Fields.Item("coste").Value = Request.Form("Mcoste") + ""; 
                }
              else oRec.Fields.Item("coste").Value = " ";   

              if ((Request.Form("Minst").Count != 0) && (Request.Form("Minst") + "" != ""))
                { 
                  oRec.Fields.Item("inst").Value = Request.Form("Minst") + ""; 
                }
              else oRec.Fields.Item("inst").Value = " ";   

              if ((Request.Form("Mdoc").Count != 0) && (Request.Form("Mdoc") + "" != ""))
                { 
                  oRec.Fields.Item("doc").Value = Request.Form("Mdoc") + ""; 
                }
              else oRec.Fields.Item("doc").Value = " ";   


              if ((Request.Form("MFundament").Count != 0) && (Request.Form("MFundament") + "" != ""))
                { 
                  oRec.Fields.Item("Fundament").Value = Request.Form("MFundament") + ""; 
                }
              else oRec.Fields.Item("Fundament").Value = " ";   

              if ((Request.Form("MSede").Count != 0) && (Request.Form("MSede") + "" != ""))
                { 
                  oRec.Fields.Item("Sede").Value = Request.Form("MSede") + ""; 
                }
              else oRec.Fields.Item("Sede").Value = " ";   
              
              if ((Request.Form("MBibliog").Count != 0) && (Request.Form("MBibliog") + "" != ""))
                { 
                  oRec.Fields.Item("Bibliog").Value = Request.Form("MBibliog") + ""; 
                }
              else oRec.Fields.Item("Bibliog").Value = " ";   

              oRec.Update();      
              oConn.Close();   
              
              Response.Redirect("agregaMod.asp?uid=" + uid);      
              
            }  
        }
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



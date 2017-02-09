<%@ Language=JavaScript %>

<%
 Response.Expires = -1;
%>
<!-- #include file="../js/Adolibrary.inc" -->
<!-- #include file="../js/user.inc" -->

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

%>

<%
var maxAttempt = 2;

var scoreLesson     = new Array();
var scorePuntuacion = new Array();
var scoreExer       = new Array();

var exerID          = new Array();
var exerLesson      = new Array();
var exerFile        = new Array();
var exerFileId      = new Array();
var exerType        = new Array();
var exerCant        = new Array();
var exerResponse    = new Array();
var exerDir         = new Array();

var Matrix          = new Array(Session("lessons").items.id[Session("lessonCant")] + 1);
var Ya              = new Array();
var YaCant          = -1;
var Pesos           = new Array();
var maxPeso         = 11;

var fullCourse = Session("fullcourse");

//Response.Write(fullCourse.ctdExAct - 1 + "<br>");
//Response.Write(fullCourse.ctdExRet - 1+ "<br>");
//Response.Write(fullCourse.Atmost - 1+ "<br>");


var actualCant      = parseInt(fullCourse.ctdExAct, 10);              // Cantidad de Ejercicios de la lección actual (2, 0 y 1)
var beforeCant      = parseInt(fullCourse.ctdExRet, 10);              // Cantidad de Ejercicios de las lecciones anteriores (2, 0 y 1)
var Cuales          = new Array();
var Cantidad        = -1;
var lastLesson      = Session("lastLesson");
var lastLessonUrl      = Session("lastLessonUrl");

//Response.Write(Session("lastLesson") + "<br>");
//Response.Write(actualCant + "<br>");
//Response.Write(beforeCant + "<br>");





function items1(lesson, puntuacion)

  { 
    this.lesson = lesson;
    this.puntuacion = puntuacion;
  }
    
function record1(lesson, puntuacion)

  { 
    this.items = new items1(lesson, puntuacion);
  } 
  
function items2(id, lesson, file, type, cant, response, fileid, dir)

  { 
    this.id = id;
    this.lesson = lesson;
    this.file = file;
    this.type = type;
    this.cant = cant;
    this.response = response;
    this.fileid = fileid;
    this.dir = dir;

  }
    
function record2(id, lesson, file, type, cant, response, fileid, dir)

  { 
    this.items = new items2(id, lesson, file, type, cant, response, fileid, dir);
  } 

var dataPath = Application('dataPath');
var oConn;
var oRec;
 
oConn = MakeConnection( oConn, dataPath );
Sql ="SELECT Lesson, Puntuation, Exercise FROM Evaluaciones WHERE ([User]=" + Session("userID") + ") and (Course=" + Session("course") + ")";

oRec = Query( Sql, oRec, oConn  );		
     
var voy1 = 0;   

while (oRec.EOF == false)
  { 
    scoreLesson[voy1] = oRec.Fields.Item("Lesson") + "";
    scorePuntuacion[voy1] = oRec.Fields.Item("Puntuation") + "";
    scoreExer[voy1] = oRec.Fields.Item("Exercise") + "";
    
    //Response.Write("scorepuntuation "+ scorePuntuacion[voy1] + " <br>scorelesson" + scoreLesson[voy1] + "<br>");
    oRec.Move(1);
    voy1++;
  }

Session("scoreCant") = voy1 - 1;  
Session("score") = new record1(scoreLesson, scorePuntuacion); 


function CanShowExer(exerid)  {
  var veces = 0;
  var five = false;

  if (voy1 > 0) {
   for (i = 0; i < voy1 - 1; i++)  {
     if (scoreExer[i] == exerid) {
       veces++;
       if (scorePuntuacion[i] == 5) {
         five = true;
       }   
     }
   }
  } 
  //Response.Write(five + "  " + veces + "<br>");
  if (!five && veces < maxAttempt ) 
    return true;
  else
    return false;  
}



Sql ="SELECT B.dir, A.ID, A.Lesson, C.Url, A.Type, A.Cant, A.Response, C.ID AS FileID FROM Ejercicios A, Lecciones B, Ficheros C WHERE (A.Lesson = B.ID) and (B.Course =" + Session("course") + ") and (A.[File] = C.Id)";

oRec = Query( Sql, oRec, oConn  );		

var voy2 = 0;
while (oRec.EOF == false)
  { 
    if (CanShowExer(oRec.Fields.Item("ID"))) {
      exerID[voy2] = oRec.Fields.Item("ID") + "";
      exerLesson[voy2] = oRec.Fields.Item("Lesson") + "";
      exerFile[voy2] = oRec.Fields.Item("Url") + "";
      exerType[voy2] = oRec.Fields.Item("Type") + "";
      exerCant[voy2] = oRec.Fields.Item("Cant") + "";
      exerResponse[voy2] = oRec.Fields.Item("Response") + "";
      exerFileId[voy2] = oRec.Fields.Item("FileID") + "";
      exerDir[voy2] = oRec.Fields.Item("dir") + "";
      //Response.Write("url " + oRec.Fields.Item("Url") + "<br>response" + oRec.Fields.Item("Response") + "<br>");
      voy2++;
	}	
    oRec.Move(1);
  }
Session("exerCant") = voy2 - 1; 
Session("exer") = new record2(exerID, exerLesson, exerFile, exerType, exerCant, exerResponse, exerFileId, exerDir);

DestroyAdoObjects( oConn, oRec );

function Make()

  {
    for (i = 0; i <= Session("lessons").items.id[Session("lessonCant")]; i++)
      for (k = 0; k <= Session("lessons").items.id[Session("lessonCant")]; k++)
        if (Matrix[i][k] == true)
          for (j = 0; j <= Session("lessons").items.id[Session("lessonCant")]; j++)
            Matrix[i][j] = Matrix[i][j] || Matrix[k][j];
  }

function Esta1(Lesson)

  {
    if (YaCant == -1)
      {
        return -1;
      }
      else
        {
          var i = 0;
          while ((Ya[i] != Lesson) && (i <= YaCant))
            i++;
          if (Ya[i] == Lesson)
            {
              return i;
            }
            else
              {
                return -1;
              }
        }
  }  
  
function Equivoco(Lesson)

  {
    var suma = 0;
    var cant = 0;
    for (i = Session("scoreCant"); i >= 0; i--)
      if ((cant < fullCourse.Atmost) && (Session("score").items.lesson[i] == Lesson))
        {
          cant++;
          suma += parseInt(Session("score").items.puntuacion[i], 10);
        }
    if (cant > 0)
      {
        promedio = suma / cant;
      }
      else
        {
          promedio = 2;
        }
    if (promedio < 3)
      {
        return 1;
      }
      else
        if (promedio < 4)
          {
            caca = Math.round(Math.random());
            if (caca == 1)
              {
                return 1;
              }
              else 
                {
                  return 0;
                }
          }
          else 
            {
              return 0;
            }
  }

function MyRandom1(maxvalue)

  {
    var st, value, num1 = 0, num2;     
    value = maxvalue.toString(10);
    while (num1 == 0) 
      {
        st = Math.random().toString(10);
        num2 = st.substr(3, value.length).valueOf(10); 
        if ((num2 > 0) && (num2 < maxvalue + 1)) 
          num1 = num2;
       }
    return(num1);
  }    

function MyRandom(maxvalue)

  {
   return parseInt(((maxvalue - 1 + 1) * Math.random() + 1),10);
 
  }    

function Primario(Lesson)

  {
    for (i = 0; i <= Session("lessonCant"); i++)
      if (Session("lessons").items.id[i] == Lesson)
        {
          if (Session("lessons").items.primary[i] == "true")
            {
              return true;
            }
            else
              {
                return false;
              }
        }
  }

function Cualquiera1(Lesson)

  {
    var tmp  = new Array();
    var caca = 0;
    for (i = 0; i < Session("exerCant"); i++)
      if (Session("exer").items.lesson[i] == Lesson)
        {
          caca++;
          tmp[caca] = i;
        }
    if (caca > 0)
      {
        return tmp[MyRandom(caca)];
      }
      else
        {
          return -1;
        }
  }

function record3(Cuales)

  { 
    this.items = Cuales;
  }

function Saca1()

  {
    for (i = 0; i <= Session("lessons").items.id[Session("lessonCant")] + 1; i++)
      Matrix[i] = new Array(Session("lessons").items.id[Session("lessonCant")] + 1);
    for (i = 0; i <= Session("lessons").items.id[Session("lessonCant")]; i++)
      for (j = 0; j <= Session("lessons").items.id[Session("lessonCant")]; j++)
        Matrix[i][j] = false;
    for (i = 0; i <= Session("grafoCant"); i++)
      Matrix[Session("grafo").items.id[i]][Session("grafo").items.lesson[i]] = true;
    Make();
    posy = Cualquiera1(lastLesson);
    if (posy > -1)
      {
        Cantidad++;
        Cuales[Cantidad] = posy;
      }
    for (i = 0; i <= Session("exerCant"); i++)
      {
        if ((Session("exer").items.lesson[i] < Session("lastLesson")) && (Esta1(Session("exer").items.lesson[i]) == -1))
          {
            YaCant++;
            Ya[YaCant] = Session("exer").items.lesson[i];
            if (Equivoco(Ya[YaCant]) == 1)
              {
                if ((Matrix[Ya[YaCant]][Session("lastLesson")] == false) && (Primario(Ya[YaCant]) == true))
                  {
                    posy = Cualquiera1(Ya[YaCant]);
                    if (posy > -1)
                      {
                        Cantidad++;
                        Cuales[Cantidad] = posy;
                      }
                  }
                  else
                    {
                      if ((Matrix[Ya[YaCant]][Session("lastLesson")] == true) && (Primario(Ya[YaCant]) == false))
                        {
                          posy = Cualquiera1(Ya[YaCant]);
                          if (posy > -1)
                            {
                              Cantidad++;
                              Cuales[Cantidad] = posy;
                            }
                        }
                        else
                          {
                            if ((Matrix[Ya[YaCant]][Session("lastLesson")] == true) && (Primario(Ya[YaCant]) == true))
                              {
                                posy = Cualquiera1(Ya[YaCant]);
                                if (posy > -1)
                                  {
                                    Cantidad++;
                                    Cuales[Cantidad] = posy;
                                  }
                              }
                          }
                    }
              }
          }
      }
    Session("Cantidad") = Cantidad;
    Session("Cuales") = new record3(Cuales);
    
  }

function ExerCant(Lesson)

  {
    var cant = 0;
    //Response.Write("Session('exerCant') " +Session("exerCant") + "<br>");
    for (i = 0; i < Session("exerCant"); i++) {
      //Response.Write("Session('exer').items.lesson[i] == Lesson " +Session("exer").items.lesson[i] + " == " +  Lesson + "<br>");
      //Response.Write(parseInt(Session("exer").items.lesson[i], 10) == parseInt(Lesson, 10));
      if (parseInt(Session("exer").items.lesson[i], 10) == parseInt(Lesson, 10))
        {
          cant++;
        }
    } 
    //Response.Write("cant " + cant + "<br>");   
    return cant;
  }

function Esta2(Exercise)

  {
    if (Cantidad == -1)
      {
        return -1;
      }
      else
        {
          var i = 0;
          while ((Cuales[i] != Exercise) && (i <= Cantidad))
            i++;
          if (Cuales[i] == Exercise)
            {
              return i;
            }
            else
              {
                return -1;
              }
        }
  }  
  
function Cualquiera2(Lesson)

  {
    var tmp  = new Array();
    var caca = 0;
    for (i = 0; i <= Session("exerCant"); i++)
      if ( (parseInt(Session("exer").items.lesson[i], 10) == parseInt(Lesson, 10)) 
            && (Esta2(i) == -1))
        {
          caca++;
          tmp[caca] = i;
        }
    if (caca > 0)
      {
        return tmp[MyRandom(caca)];
      }
      else
        {
          return -1;
        }
  }
  
function Peso(Lesson)

  {
    var suma = 0;
    var cant = 0;
    for (i = Session("scoreCant"); i >= 0; i--)
      if ((cant < parseInt(fullCourse.Atmost, 10)) && (Session("score").items.lesson[i] == Lesson))
        {
          cant++;
          suma += parseInt(Session("score").items.puntuacion[i], 10);
        }
    if (cant > 0)
      {
        promedio = suma / cant;
      }
      else
        {
          promedio = 2;
        }
    return Math.round(maxPeso - 2 * promedio);
  }

function Saca2()

  {
    //Obtener cantidad de ejercicios de la leccion actual...
    cantEjercicios = ExerCant(lastLesson);
    
//    Response.Write("lastlesson  "+ lastLesson + "<br>");
//    Response.Write("ExerCant(lastLesson)  "+ ExerCant(lastLesson) + "<br>");
    
    //Extraer ejercicos de la leccion actual
    Continue       = true;
    while (Continue == true)
      {
        //Obtener un ejercicio cualquiera de la leccion actual...
        posy = Cualquiera2(lastLesson);
        
        //Response.Write("posy  "+ posy + "<br>");
        
        //Si encontro alguno
        if (posy > -1)
          {
            Cantidad++;
            Cuales[Cantidad] = posy;
            
            //Si se acabaron los ejercicos o ya se cumplio la norma para la leccion actual...
            if ((Cantidad == cantEjercicios) || (Cantidad + 1  == actualCant))
              {
                Continue = false;
              }
          }
          else //No encontro...
            {
              Continue = false;
            }
      }
      //Response.Write(Session("ctdExRet"));
      
    //.......................................................
    
    //Extraer ejercicos de lecciones anteriores...      

    //Si hay retroalimentacion...
    if ((fullCourse.Atmost > 0) && (beforeCant > 0))
      {

    //Calcular las probabilidades... 
    for (var i = 0; i <= Session("lessonCant"); i++)
      {
        //Si la leccion es precendente de la actual...
        if (Session("lessons").items.id[i] < Session("lastLesson"))
            Pesos[i] = Peso(Session("lessons").items.id[i]);
      }

    
    //Response.Write("ca " + Cantidad + "<br>");
    
    //Si ni hubo ejercicios de la leccion actual
    if (Cantidad == -1) 
      { max = beforeCant; } 
    else 
      { max = beforeCant + actualCant;}
      
    //Response.Write("max " + max + "<br>");
    
    
    Continue = true;  
    while (Continue == true)
      {
        suma     = parseInt(MyRandom(maxPeso), 10);
        Encontre = false;
        for (var i = 0; i <= Session("lessonCant"); i++)
          {
            //Si la leccion es precendente de la actual...
            if (Session("lessons").items.id[i] < Session("lastLesson"))
              {
                suma -= Pesos[i];
                if (suma <= 0)
                  {
                    cantEjercicios = ExerCant(Session("lessons").items.id[i]) - 1;
                    Salgo          = false;
                    while (Salgo == false)
                      {
                        //Escoger cualquiera de la leccion...
                        posy = Cualquiera2(Session("lessons").items.id[i]);
                        
                        //Response.Write( "  posy " + posy + "<br>");
                        
                        //si encontro...
                        if (posy > -1) 
                          {
                            Cantidad++;
                            Cuales[Cantidad] = posy;
                            
                            
                            if (Cantidad + 1  == max) 
                              {
                                Salgo    = true; 
                                Continue = false; 
                                Encontre = true;
                              }
                          }
                          else //No encontro...
                            {
                              Salgo = true;
                            }
                      }
                    if (Continue == false)
                      {
                        break;
                      }
                  }
              }
          }
        if (Encontre == false)
          {
            Continue = false;
          }
      }
    }
    Session("Cantidad") = Cantidad;
    Session("Cuales") = new record3(Cuales);
  //  Response.Write("Session('Cantidad') "+ Session("Cantidad") + "");
  }
  
Saca2();

%>

<HTML>
<HEAD>
<link REL="stylesheet" TYPE="text/css" HREF="../css/<%=Session("skin")%>/Main.css" />
<link rel="stylesheet" href="../css/<%=Session("skin")%>/EvalResult.css" type="text/css">

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script language=jscript src="../js/ejercicios.js"></script>
<script type="text/javascript" language="javascript" src="../js/Windows.js"></script>
<script type="text/javascript">
<!--

function getFullPath(path) {
  return '../Courses/course<%=Session("course")%>/' + path;

}




function putImg(path) {

  var im = new Image;
  im.src = '../Courses/course<%=Session("course")%>/' + path; 
  //alert(getFullPath(path));
  document.write('<img src=' + getFullPath(path) + ' />');  

}

function putVideo(path) {
  //var im = new Image;
  //im.dynsrc = '../Courses/course<%=Session("course")%>/' + path; 
  //document.write('<br><center><img loop=infinity dynsrc="' + getFullPath(path) + '" alt="' + getFullPath(path) + '" /></center>');  
  document.write('<br><object id="NSPlay" classid="CLSID:22D6F312-B0F6-11D0-94AB-0080C74C7E95">' + 
                 '<param name="ControlType" value="0">' + 
  		 '<param name="AutoRewind" value="1">' + 
                 '<param name="AutoSize" value="1">' + 
                 '<param name="AutoStart" value="0">' + 
                 '<param name="BaseURL" value="">' + 
                 '<param name="Filename"  value= "' + getFullPath(path) + '" >' + 
                 '<param name="ShowControls" value="1">' + 
                 '<param name="ShowAudioControls" value="1">' + 
                 '<param name="ShowPositionControls" value="1">' + 
                 '<param name="ShowTracker" value="1">' + 
                 '<param name="ShowGotoBar" value="0">' + 			 		
                 '<param name="InvokeURLs" value="1">' + 
                 '<param name="Volume" value="10">' + 
                 '<param name="AllowChangeDisplaySize" value="0">' + 
                 '</object>');
  
}

function putSound(path) {
  //var im = new Image;
  //im.dynsrc = '../Courses/course<%=Session("course")%>/' + path; 
  //document.write('<br><center><img dynsrc="' + getFullPath(path) + '"  /></center>');  
  document.write('<br><object id="NSPlay" classid="CLSID:22D6F312-B0F6-11D0-94AB-0080C74C7E95">' + 
                 '<param name="ControlType" value="0">' + 
  		 '<param name="AutoRewind" value="1">' + 
                 '<param name="AutoSize" value="1">' + 
                 '<param name="AutoStart" value="0">' + 
                 '<param name="BaseURL" value="">' + 
                 '<param name="Filename"  value= "' + getFullPath(path) + '" >' + 
                 '<param name="ShowControls" value="1">' + 
                 '<param name="ShowAudioControls" value="1">' + 
                 '<param name="ShowPositionControls" value="1">' + 
                 '<param name="ShowTracker" value="1">' + 
                 '<param name="ShowGotoBar" value="0">' + 			 		
                 '<param name="InvokeURLs" value="1">' + 
                 '<param name="Volume" value="10">' + 
                 '<param name="AllowChangeDisplaySize" value="0">' + 
                 '</object>');
  
}

function putLink(path) {
  document.write(' <a target=_blank href="' + getFullPath(path) + '"  >Descargar</a>');  
}


function goFile(path) {
  abreVentana('', 600, 400, '../Courses/course<%=Session("course")%>/' + path, 'yes');
}
//-->
</script>

</HEAD>
<body bgcolor="#FFFFFF" text="#000000" style="overflow-y:scroll">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td><img src="../images/<%=Session("skin")%>/EvalResult.gif" width="80" height="54"></td>
          <td class="HeaderTable">Autoevaluaci&oacute;n</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="10" cellpadding="10">
  <tr> 
    <td>
<%

//Creacion de las preguntas

 
function setAsk()
  { 
    if (Session("Cantidad") >= 0)
     {
      for (i = 0; i <= Session("Cantidad"); i++)
       { 
        lessondir = Session("exer").items.dir[Session("Cuales").items[i]];
        if  ((lessondir + "" != "null") &&
             (lessondir + "" != ""))
          {
            lessondir = "/" + lessondir;      
          }
        else lessondir = "";
        
        
        //  *********OJO*******
        //Esto es hasta decidir si los ejercicios estaran en el dir de la leccion de 
        //lo contrario dejerlo asi.
        lessondir = "";
        
        
            
        Response.Write("<H6>" + (i + 1) + ".-</H6>");
        Response.Write('<SCRIPT LANGUAGE=JAVASCRIPT SRC="../Courses/Course' + Session("course") + lessondir + '/' + Session("exer").items.file[Session("Cuales").items[i]] + '"></Script>');
                                                                                                                    
       }
       Response.Write('<Center><INPUT id=button1 name=button1 type=button value=Aceptar onclick ="button1_onclick()"></Center>'); 
     }
    else 
     {
      Response.Write('<Table Class="MainTable"><TR Class="MainTR"><TD>Esta lección no tiene ejercicios asociados. Presione el botón <B>Aceptar<B/> para continuar.</TD></TR>');
      Response.Write('<TR Class="MessageTR"><TD><INPUT id=button2 name=button2 type=button value=Aceptar onclick = "button2_onclick()"></TD></TR></Table>');
     }   
  }

setAsk();

%>
<FORM  id=asks method=post action=checkreponse.asp?uid=<%=uid%>></FORM>

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function button1_onclick() 
{
     
   var coll = document.all;
   var listName = new Array();
   var listValue = new Array();
   var i;
   var k = 0;
   for (i=0;i< coll.length;i++)
     { 
      if ((coll[i].tagName == "INPUT") && (coll[i].type == "checkbox")) 
        {
         listName[k] = coll[i].name;   
         listValue[k] = coll[i].checked + "";
         k++;
        }
       else 
        if ((coll[i].tagName == "SELECT") ||(coll[i].tagName == "INPUT") || (coll[i].tagName == "TEXTAREA")) 
         {
          listName[k] = coll[i].name;  
          listValue[k] = coll[i].value; 
  
          k++;
         }
     }
 
  for (i=0;i< k;i++)
   { //alert(listName[i]);
     
//alert(listValue[i]);
    asks.insertAdjacentHTML("BeforeEnd","<input type = 'hidden'  id=" + listName[i] + " name=" + listName[i] + " value= '" + listValue[i] + "' >");
   //alert(listName[i] + "  " + listValue[i]); 
   }  
 
  asks.submit();  
}


function button2_onclick() 
{
 location.href = "docs.asp?uid=<%=uid%>&lesson=<%=Session("lastLesson")%>&lessonurl=<%=Session("lastLessonUrl")%>&lessonindex=<%=Session("lastLessonIndex")%>&lessonname=<%=Session("lastLessonName")%>";
}

//-->
</SCRIPT>


    </td>
  </tr>
</table>

</BODY>
</HTML>
<%
 }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
%>

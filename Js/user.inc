<%

//Esta biblioteca requiere que se use adoLibrary.js

var SHOW_CANT = Application("ListSize");
var MATRICULAS_PAGE = 'abreVentana("Matricular1", 500, 400,  "matriculas.asp", "yes", "no")';

//Errores del manejo de ficheros
INVALID_FILE_SHORT = 'Archivo no encontrado';
INVALID_FILE_TEXT  = 'El archivo especificado no existe, fue eliminado o tiene tama�o 0.'
INVALID_FILE_EXTENTION_SHORT = 'Tipo de archivo no v�lido';
INVALID_FILE_EXTENTION_TEXT  = 'El archivo especificado tiene una extensi�n permitida por el sistema.';
INVALID_FILE_EXTENTION_SHORT = 'Tipo de archivo no v�lido';
INVALID_FILE_EXTENTION_TEXT  = 'El archivo especificado no tiene un formato no soportado por el sistema.';
INVALID_FILE_SIZE_SHORT = 'Tama�o excedido';
INVALID_FILE_SIZE_TEXT = 'El archivo especificado excede el tama�o permitido para este m�dulo.';
DIRECTORY_FULL_SHORT = 'Directorio lleno';
DIRECTORY_FULL_TEXT = 'El directorio del m�dulo se encuentra lleno o no hay suficiente espacio para el archivo especificado.';
//Algunos errorer generales
SENDING_MAIL_SHORT = "Error enviando correo de notificaci�n."
var INVALID_DATE_SHORT = "Fecha no v�lida.";
var INVALID_DATE_TEXT = "Fecha no v�lida.";

var NOT_EXER_EXIST_SHORT = "Ejercicio no v�lido.";
var NOT_EXER_EXIST_TEXT = "El ejercicio pedido no existe o no est� disponible.";

var LOGIN_EXIST_SHORT = "Nombre de usuario no v�lido.";
var LOGIN_EXIST_TEXT = "El identificador especificado por usted ya se encuentra asignado a otro usuario. Presione regresar para introducir los datos nuevamente.";

var GRUOP_NAME_EXIST_SHORT = "Ya existe un grupo con este nombre.";
var GRUOP_NAME_EXIST_TEXT = "El nombre especificado por usted ya se encuentra asignado a otro grupo.";
var MOD_NAME_EXIST_SHORT = "Ya existe una modalidad acad�mica con este nombre.";
var MOD_NAME_EXIST_TEXT = "El nombre especificado por usted ya se encuentra asignado a otra modalidad acad�mica.";
var MOD_NAME_NOT_EXIST_SHORT = "No existe una modalidad acad�mica con este nombre.";
var MOD_NAME_NOT_EXIST_TEXT = "El nombre especificado por usted no se encuentra asignado a ninguna modalidad acad�mica.";
var CUR_NAME_EXIST_SHORT = "Ya existe un m�dulo este nombre.";
var CUR_NAME_EXIST_TEXT = "El nombre especificado por usted ya se encuentra asignado a otro m�dulo.";


var INVAILD_COURSE_SHORT = "M�dulo no existe.";
var INVAILD_COURSE_TEXT = "El m�dulo al que usted hace referencia no esta disponible o no existe.";

var SESSION_TIMEOUT_SHORT = "Se venci� su tiempo de espera en el sistema.";
var SESSION_TIMEOUT_TEXT = "Se venci� su tiempo de espera en el sistema o la p�gina fue accedida incorrectamente. La causa m�s frecuente de este error es permanecer m�s tiempo que el establecido (" + Session.Timeout + " minutos)  sin ejecutar ninguna acci�n. Si desea continuar vuelva a conectarse y autentificarse.";
var BAD_PARAMS_SHORT = "Par�metros no v�lidos.";
var BAD_PARAMS_TEXT = "Usted ha realizado una operaci�n ilegal o ha tratado de acceder una p�gina con par�metros no v�lidos.";
var DONT_HAS_PERMISON_SHORT = "Insuficientes permisos.";
var DONT_HAS_PERMISON_TEXT = "Usted no posee permisos para realizar esta acci�n.";
var INVALID_CHARACTERS_SHORT = "Car�cteres no v�lidos.";
var INVALID_CHARACTERS_TEXT = "Usted ha introducido car�cteres no permitidos en los campos del formulario que acaba de completar. Presione regresar para introducir los datos nuevamente.";
var COURSE_DONT_HAS_GROUP_SHORT = "NO hay grupo asociado.";
var COURSE_DONT_HAS_GROUP_TEXT = "Este m�dulo no tiene un grupo de usuarios asociados por lo que no se puede matricular alumnos en �l.";

var EXIST_QUESTION_SHORT = "Ya existe su pregunta";
var EXIST_QUESTION_TEXT = "La pregunta que intenta agregar ya est� entre las preguntas m�s frecuentes de este m�dulo. Si desea modificar su respuesta elija la opci�n Modificar Pregunta del men�.";

//Constantes para el trabajo con el usuario... 
var UNDEFINED_USER_SHORT     = "Usuario no definido.";
var INVALID_PASSWORD_SHORT   = "Contrase�a no v�lida.";
var ACTIVED_USER_SHORT       = "Usuario activo.";
var DIFFERENT_PASSWORD_SHORT = "No coinciden las contrase�as."
var VALID_USER_SHORT         = "Autentificaci�n satisfactoria.";



var UNDEFINED_USER_TEXT     = "El usuario especificado no forma parte de nuestro sistema, si usted esta interesado haga click <a href='javascript:" + MATRICULAS_PAGE  +  "'>aqui</a>";
var INVALID_PASSWORD_TEXT   = "Contrase�a no v�lida, asegurese que el Caps Lock no este activado. Si usted esta interesado en matricular haga su solicitud <a href='javascript:" + MATRICULAS_PAGE  +  "'>aqui</a>";
var ACTIVED_USER_TEXT       = "El usuario especificado ya se encuentra activo, o ha habido problemas a la hora de cerrar el SEPAD, si este es el caso, por favor espere un minuto.";
var DIFFERENT_PASSWORD_TEXT = "La nueva contrase�a entrada y su confirmaci�n no coinciden."
var VALID_USER_TEXT         = "Usted se ha identificado satisfactoriamente.";
var GUEST_USER_TEXT         = "Usuario invitado";
var GUEST_LOGIN             = "guest";		
var GUEST_PASSWORD          = "";		

var UNDEFINED_USER     = 100;
var INVALID_PASSWORD   = 120;
var ACTIVED_USER       = 130;
var DIFFERENT_PASSWORD = 140;
var VALID_USER         = 150;
var GUEST_USER         = 580;
var ADMIN_USER		   = 1830;

//Grupos predefinidos
var ADMIN_GROUP		   = 70;
var PUB_GROUP		   = 72;
var SEPAD_GROUP		   = 60;
var GUEST_GROUP		   = 4;

                                    

//Constantes de permisos
var USER_TEXT = "Usuario";
var GUEST = "Usuario Invitado";
var PUBLICATOR_TEXT = "Publicador";
var ADMINISTRATOR_TEXT = "Administrador";

var USER = 0;
var GUEST = -1;
var PUBLICATOR = 10;
var ADMINISTRATOR = 20;


//Estados de las modalidades acasdemicas...
var MOD_ACA_NOVISIBLE = 4;
var MOD_ACA_DISABLE   = 5;
var MOD_ACA_ININSCRIPTION = 0;
var MOD_ACA_INCOURSE = 1;
var MOD_ACA_INCOURINSC = 2;
var MOD_ACA_FINISHED = 3;
var MOD_ACA_FREE = 6;
var MOD_ACA_DEVELOP = 7;


var MOD_ACA_SELECT_STATE = "<select name=Mstate>" + 
                              "<option value=" + MOD_ACA_NOVISIBLE + ">No Visible</option>" + 
                              "<option value=" + MOD_ACA_DISABLE + ">Deshabilitado</option>" + 
                              "<option value=" + MOD_ACA_ININSCRIPTION + ">En inscripci�n</option>" + 
                              "<option value=" + MOD_ACA_INCOURSE + ">En curso</option>" + 
                              "<option value=" + MOD_ACA_INCOURINSC + ">En inscripci�n - En curso</option>" + 
                              "<option value=" + MOD_ACA_FINISHED + ">Finalizado</option>" + 
                              "<option value=" + MOD_ACA_FREE + ">Libre</option>" + 
                              "<option value=" + MOD_ACA_DEVELOP + ">En Desarrollo</option>" + 
                           "</select>";

var CUR_SELECT_STATE =     "<select name=Mstate>" + 
                              "<option value=" + MOD_ACA_NOVISIBLE + ">No Visible</option>" + 
                              "<option value=" + MOD_ACA_DISABLE + ">Deshabilitado</option>" + 
                              "<option value=" + MOD_ACA_INCOURSE + ">En curso</option>" + 
                              "<option value=" + MOD_ACA_FINISHED + ">Finalizado</option>" + 
                              "<option value=" + MOD_ACA_FREE + ">Libre</option>" + 
                           "</select>";



  function User(id, name, fullName, email) {
    this.id = id;
    this.name = name;
    this.fullName = fullName;
    this.email = email;
  }


function GetUserData(userid) {
   var dataPath = Application('dataPath');
   var oConn;
   var oRec;
   var res;
   oConn = MakeConnection( oConn, dataPath );

  
   oRec = Query( "select * from Usuarios where (id = " + userid + ")", oRec, oConn  );		

   if (oRec.EOF == false) {
     res = new User(oRec.Fields.Item("ID").Value, oRec.Fields.Item("Name").Value, oRec.Fields.Item("FullName").Value, oRec.Fields.Item("email").Value);  	
     return res; 
   } 
   else {
     return -1; //El usuario no existe

   }
   DestroyAdoObjects( oConn, oRec );
}


  function Group(id, name, desc) {
    this.id = id;
    this.name = name;
    this.desc = desc;
  }


function GetGroupData(groupid) {
   var dataPath = Application('dataPath');
   var oConn;
   var oRec;
   var res;
   oConn = MakeConnection( oConn, dataPath );

  
   oRec = Query( "select * from Grupos where (ID = " + groupid + ")", oRec, oConn  );		
   
   if (oRec.EOF == false) {
     
     res = new Group(oRec.Fields.Item("ID").Value, oRec.Fields.Item("Name").Value, oRec.Fields.Item("Description").Value);  	
     return res; 
   } 
   else {
     return -1; //El grupo no existe

   }
   DestroyAdoObjects( oConn, oRec );
}

//Crear un grupo.. asume que la coneccion esta creada 
function CreateGroup(name, desc) {
   var res = -1;

   oRec.Open("select * from Grupos where name = " + name,oConn,3,3);
   if (oRec.EOF == true) {
    oRec.AddNew();
    oRec.Fields.Item("Name").Value = name;
    oRec.Fields.Item("Description").Value = desc;
    oRec.Update();      
    oRec.MoveLast();
    res = oRec.Fields.Item("ID").Value + "";
   } 
   else {
     oRec.Close();
     return -2; //El grupo ya existe

   }
   oRec.Close();
   
   return res;
}


//Crear un grupo
function AddGroup(name, desc) {
   var dataPath = Application('dataPath');
   var oConn;
   var oRec;
 
   oConn = MakeConnection( oConn, dataPath );

   var res = -1;
  
   oRec = Query( "select * from Grupos where (name = '" + name + "')", oRec, oConn  );		

   if (oRec.EOF == true) {
    oRec.AddNew();
    oRec.Fields.Item("Name").Value = name;
    oRec.Fields.Item("Description").Value = desc;
    oRec.Update();      
    oRec.MoveLast();
    res = oRec.Fields.Item("ID").Value + "";
   } 
   else {
     oRec.Close();
     return -2; //El grupo ya existe

   }
   DestroyAdoObjects( oConn, oRec );
   return res;
}


//adicionar un usuario a un grupo... suponre que la conneccion esta creada
function AddUserGroup(userId, groupId) {
   oRec.Open("SELECT * From Grupos_de_Usuarios",oConn,3,3);
   oRec.AddNew();
   oRec.Fields.Item("User").Value = userId;
   oRec.Fields.Item("Group").Value = groupId;
   oRec.Update();      
   oRec.Close();
}	

//Adicionar un usuario a un grupo
function AddUserToGroup(userID, groupID) {
    if (!isUserInGroupByID(userID, groupID)) 
     {
     var dataPath = Application('dataPath');
      var oConn;
      var oRec;
 
      oConn = MakeConnection( oConn, dataPath );
      Sql =" SELECT * " + 
           " FROM Grupos_de_Usuarios";
      oRec = Query( Sql, oRec, oConn  );		
      
      oRec.AddNew();
      oRec.Fields.Item("User").Value = userID;
      oRec.Fields.Item("Group").Value = groupID;
      oRec.Update();
      DestroyAdoObjects( oConn, oRec );
    }
}

function es_usuario()
  {
  	return false;
  }

function es_tutor()
  {
  	return false;
  }




//chequea si un usuario pertenece a un grupo dado el nombre del grupo...
function isUserInGroup(userid, groupname) {
    var dataPath = Application('dataPath');
    
    var oConn;
    var oRec;
 
    oConn = MakeConnection( oConn, dataPath );
    Sql =" SELECT Grupos_de_Usuarios.ID " + 
         " FROM Grupos INNER JOIN Grupos_de_Usuarios ON Grupos.ID = Grupos_de_Usuarios.[Group] " +
         " WHERE (((Grupos_de_Usuarios.[User])=" + userid + ") AND ((Grupos.Name)='" + groupname + "'))";

    oRec = Query( Sql, oRec, oConn  );		
    var res =  !(oRec.EOF); 
//    Response.Write(Sql);
    DestroyAdoObjects( oConn, oRec );
    return(res);
}

//Chequea si un usuario pertenece a un grupo dado el id
function isUserInGroupByID(userid, groupid) {
    var dataPath = Application('dataPath');
    
    var oConn;
    var oRec;
 
    oConn = MakeConnection( oConn, dataPath );
    Sql =" SELECT Grupos_de_Usuarios.ID " + 
         " FROM Grupos INNER JOIN Grupos_de_Usuarios ON Grupos.ID = Grupos_de_Usuarios.[Group] " +
         " WHERE (((Grupos_de_Usuarios.[User])=" + userid + ") AND ((Grupos.ID)=" + groupid + "))";

    oRec = Query( Sql, oRec, oConn  );		
    var res =  !(oRec.EOF); 
//    Response.Write(Sql);
    DestroyAdoObjects( oConn, oRec );
    return(res);
}

function SubGrupo(ID, Name) {
   this.ID = ID;
   this.Name = Name;
}

//Chequea si un usuario es tutor de un curso
function isTutorCurso(userid, cursoid) {
    var dataPath = Application('dataPath');
    
    var oConn;
    var oRec;
 
    oConn = MakeConnection( oConn, dataPath );
    Sql =" SELECT * " + 
         " FROM SubGrupo " +
         " WHERE ((Tutor=" + userid + ") AND (Curso=" + cursoid + "))";

    oRec = Query( Sql, oRec, oConn  );		
    if (oRec.EOF == true) {
        var res = new SubGrupo(-1, ""); 
    }
    else  {
       var res = new SubGrupo(oRec.Fields.Item("ID").value, oRec.Fields.Item("Name").value); 
    }
//    Response.Write(Sql);
    DestroyAdoObjects( oConn, oRec );
    return(res);
}


 //Chequea si un usuario es valido para un modulo
 function IsValidInModule(userID, courseID) {
    var valid = false;
    var dataPath = Application('dataPath');
    
    var oConn;
    var oRec;
 
    oConn = MakeConnection( oConn, dataPath );
    Sql = "SELECT Cursos.Owner, Modulo.cordinador, Modulo.claustro, Modulo.grupo, Cursos.state as Cstate, Modulo.state as Mstate " + 
          "FROM Modulo INNER JOIN Cursos ON Modulo.ID = Cursos.modulo " + 
          "WHERE (((Cursos.ID)=" + courseID + "))";

    oRec = Query( Sql, oRec, oConn  );	
    
    if (!oRec.EOF) {
       var prof = isUserInGroupByID(userID, oRec.Fields.Item("claustro").value);	
       var estd = isUserInGroupByID(userID, oRec.Fields.Item("grupo").value);
       //Response.Write('Session("UserID") + "  " +  "oRec.Fields.Item("cordinador").value" + "oRec.Fields.Item("owner").value" + "estd" + "prof"' +   "<BR>");
       //Response.Write(Session("UserID") + "  " +  oRec.Fields.Item("cordinador").value + "  " +  oRec.Fields.Item("owner").value  + "  " +   estd  + "  " +   prof  +     "<BR><BR>");
       var isowner = ((Session("UserID") == oRec.Fields.Item("cordinador").value) || (Session("UserID") == oRec.Fields.Item("owner").value) || (Session("permissionType") == ADMINISTRATOR));
       var valid = ( isowner || (prof) || (estd));
       valid = valid && ((oRec.Fields.Item("Cstate").value == MOD_ACA_INCOURSE) && ((oRec.Fields.Item("Mstate").value == MOD_ACA_INCOURSE) || (oRec.Fields.Item("Mstate").value == MOD_ACA_INCOURINSC)))
    }
    	
   DestroyAdoObjects( oConn, oRec );
   return valid;	  
 }

 //Chequea si un usario con su clave son validos      
 function checkuser(username, userpassword) 
  {

    
    /* Coneccion con la tabla de usuarios */
    //Response.Write(Application('dataPath'));
    var dataPath = Application('dataPath');
    
    var oConn;
    var oRec;
 
    oConn = MakeConnection( oConn, dataPath );
    Sql ="select A.ID,A.Name, A.skin, A.Password,A.PermissionType,A.lastLogin, A.FullName,A.Email,A.Logins from Usuarios A where (A.Name='" + username + "')";

    oRec = Query( Sql, oRec, oConn  );		
    /* Fin Coneccion con la tabla usuario */
   
    /*Chequeo de validez del usuario*/
    var password = "";
    if (oRec.EOF) //no hay ningun usuario con ese login y password...
       {
        //..........login de un usuario desconocido
        return UNDEFINED_USER; 
       }
    else 
       {
        // usuario si existe.......                    
        password              = oRec.Fields.Item("Password").value;                             

        //Response.Write("  " + password + "     "  );

        if ((password == userpassword ) || (oRec.Fields.Item("ID").value == GUEST_USER)) //Si la clave esta correcta
          { 
           Session("userID")         = oRec.Fields.Item("ID").value;
           Session("permissionType") = oRec.Fields.Item("PermissionType").value;
           Session("logins")         = oRec.Fields.Item("Logins").value + 1; 
           Session("fullName")       = oRec.Fields.Item("FullName").Value;
           Session("email")          = oRec.Fields.Item("Email").value;
           Session("name")	     = oRec.Fields.Item("Name").value;  
           Session("userPassword")   = oRec.Fields.Item("Password").value;  
           Session("skin")           = oRec.Fields.Item("skin").value;  
           if ((Session("skin") == null) || (Session("skin") == "")) Session("skin") = "aqua"; 

           
           //Actualizar la cantidad de logeos del usuario
           oRec.Fields.Item("Logins").Value = oRec.Fields.Item("Logins").Value + 1;
           var hoy =  new Date();
           oRec.Fields.Item("lastLogin").Value = hoy.getVarDate();
           oRec.Update();

           permissionType = Session("permissionType");                           
           
           if (permissionType == USER) {Session("textPermission") = USER_TEXT;}
           else if (permissionType == PUBLICATOR) {Session("textPermission") = PUBLICATOR_TEXT;}
                else if (permissionType == ADMINISTRATOR){Session("textPermission") = ADMINISTRATOR_TEXT;}

           return VALID_USER;
          }
        else
          {
           //password no v�lido......  
           return INVALID_PASSWORD; 
          }   
        }                 
      DestroyAdoObjects( oConn, oRec );
  }
  
 function DoubleQuotationToSimpleQuotation(str) {
 		var r;
 		re = /"/g;
 		r = str.replace(re, "'");
 		return(r); 
 }

  %>

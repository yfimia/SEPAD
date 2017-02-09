
var MATRICULAS_PAGE = 'abreVentana("Matricular1", 375, 365,  "matriculas.asp", "no", "no")';


//Algunos errorer generales
var INVALID_DATE_SHORT = "Fecha no v�lida.";
var INVALID_DATE_TEXT = "Fecha no v�lida.";

var NOT_EXER_EXIST_SHORT = "Ejercicio no v�lido.";
var NOT_EXER_EXIST_TEXT = "El ejercicio pedido no existe o no est� disponible.";

var SESSION_TIMEOUT_SHORT = "Se venci� su tiempo de espera en el sistema.";
var SESSION_TIMEOUT_TEXT = "Se venci� su tiempo de espera en el sistema o la p�gina fue accedida incorrectamente. La causa m�s frecuente de este error es estar mucho tiempo sin ejecutar ninguna acci�n en el sistema y este tiene un tiempo m�ximo de espera establecido en el servidor. Si desea continuar vuelva a conectarse al sistema.";
var BAD_PARAMS_SHORT = "Par�metros no v�lidos.";
var BAD_PARAMS_TEXT = "Usted ha realizado una operaci�n ilegal o ha tratado de acceder una p�gina con par�metros no v�lidos.";
var DONT_HAS_PERMISON_SHORT = "Insuficientes permisos.";
var DONT_HAS_PERMISON_TEXT = "Usted no posee permisos para realizar esta acci�n.";
var INVALID_CHARACTERS_SHORT = "Car�cteres no v�lidos.";
var INVALID_CHARACTERS_TEXT = "Usted ha introducido car�cteres no permitidos en los campos del formulario que acaba de completar.";
var COURSE_DONT_HAS_GROUP_SHORT = "NO hay grupo asociado.";
var COURSE_DONT_HAS_GROUP_TEXT = "Este curso no tiene un grupo de usuarios asociados por lo que no se puede matricular alumnos en �l.";




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

                                    

//Constantes de permisos
var USER_TEXT = "Usuario";
var GUEST = "Usuario Invitado";
var PUBLICATOR_TEXT = "Publicador";
var ADMINISTRATOR_TEXT = "Administrador";

var USER = 0;
var GUEST = -1;
var PUBLICATOR = 10;
var ADMINISTRATOR = 20;

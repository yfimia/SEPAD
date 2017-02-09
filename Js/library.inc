 <%
 /*
   Este metodo crea una tabla a partir de una cadena de conexion y una query
     dataPath : cadena de conexion
     Title: titulo con que se mostrara la tabla.
     Image: Imagen para la tabla, si es "" se pone la que trae por defecto.
     Sql : consulta a mostrar
     TotalField : posicion del campo para el que se desea calcular el total
                  comienza en 1 y si es -1 no se calcula ninguno.
     Ordered    : campo por el que esta oredenado la tabla y por el de refenecia para mostrar por lotes              
 */

 var last = "-1";

 function putTable(dataPath, Title, Image,  Sql, TotalField, Ordered) {
 var oConn;
 oConn = MakeConnection( oConn, dataPath ); 
 var oRec;
 oRec = Query( Sql, oRec, oConn  );		
 if (Image == "") Image = "../images/ReportesIMG.gif";

  %>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="ToolBar">
      <table border="0" cellspacing="0" cellpadding="0" class="ToolBar">
        <tr> 
          <td><img src="<%=Image%>" width="80" height="54"></td>
          <td class="HeaderTable"><%=Title%></td>
        </tr>
      </table>
    </td>
  </tr>
</table>  
<table width="100%" border="0" cellspacing="1" cellpadding="1">
  <tr> 
    <%
    for (a=0;a<oRec.Fields.Count;a++) {
    %>
    <td width="30%" class="ToolBar" align="center"><b><%=oRec.Fields.Item(a).name%></b></td>
    <%
    }
    %>
  </tr>    

  <%
  var a = 0;
  var Tclase = 'MessageTR';
  var total = 0;


  last = "-1";


  while (!(oRec.EOF)) {
            last = oRec.Fields.Item(Ordered).value;          

%>
          <tr>
<%
          if (Tclase == 'MessageTR') {Tclase = 'MessageTR1';} else {Tclase = 'MessageTR';} 
          for (b=0;b<oRec.Fields.Count;b++) {
            if ((TotalField  != -1) && (TotalField - 1 == b)) total += oRec.Fields.Item(b);

%>
          <td width="30%" class="<%=Tclase%>" align="center"><%=oRec.Fields.Item(b)%></td>
          
<%
            }
%>
        </tr>

<%
    a++;
    oRec.Move(1);
    }
  if (TotalField  != -1) {
%>
       <tr>
         <td  class="<%=Tclase%>" align="center"><b>Total:</b></td>       
         <td  align=right colspan=<%=oRec.Fields.Count - 1 %> class="<%=Tclase%>" align="center"><b><%=total%></b></td>       
       </tr>
<%  
  }
    
  %></table><%  
      DestroyAdoObjects( oConn, oRec );        
   
  }

 //Encia un correo interno
 function SendMsg( From, To, Subject, Body)
  {
    var filePath = Application('dataPath');
    var oConn;
    var oRec;
 
    oConn = MakeConnection( oConn, filePath );
    Sql = "SELECT * FROM Mensajeria"; 
    oRec = Query( Sql, oRec, oConn  );		
    
    oRec.AddNew();
    oRec.Fields.Item("From").Value = From;
    oRec.Fields.Item("To").Value = To;
    oRec.Fields.Item("Subject").Value = Subject;
    oRec.Fields.Item("Content").Value = Body;
    oRec.Fields.Item("Date").Value = getAtualTime();
    oRec.Update(); 

    DestroyAdoObjects( oConn, oRec );
  }

 function MarkReaded( ID )
  {
    var filePath = Application('dataPath');
    var oConn;
    var oRec;
 
    oConn = MakeConnection( oConn, filePath );
    Sql = "SELECT * FROM Mensajeria WHERE ID = " + ID; 
    oRec = Query( Sql, oRec, oConn  );		
    
    oRec.Fields.Item("Readed").Value = true;
    oRec.Update(); 

    DestroyAdoObjects( oConn, oRec );
  }
 
 //Obtiene la cantidad de visitas al sistema
 function getCantVisist()
  {
    var filePath = Application('dataPath');
    var oConn;
    var oRec;
 
    oConn = MakeConnection( oConn, filePath );
    Sql = "select SUM(Logins) as numvisit from Usuarios"; 

    oRec = Query( Sql, oRec, oConn  );		
    
    
    Application.Lock;
     Application("usersCount") = oRec.Fields.Item(0).value;
    Application.UnLock;  

    Session("userVisit") = oRec.Fields.Item(0).value;

    DestroyAdoObjects( oConn, oRec );
   
   return Session("userVisit");
  }


//Chequea si no estan tratanso de ejecutar un script...
  function notIsScript(cad)

      {  //alert(cad);
       if (cad != '')
         {
          if ((cad.indexOf('>') != -1) || (cad.indexOf('&#62;') != -1) || (cad.indexOf('&gt;') != -1))
             {
              return (false); 
             }
          else
           return true;  
         } 
       else
         return true; 

      }
%>
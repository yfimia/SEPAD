<%@language = "JScript"%>

<!-- #include file='../js/adolibrary.inc' -->
<!-- #include file='../js/user.inc' -->


<html>
<head>
<title>
Lista de Anotaciones
</title>
</head>

<body>
<%
var  filePath = Application("dataPath");   
var  oConn    = Server.CreateObject("ADODB.Connection");
var  oRec     = Server.CreateObject("ADODB.Recordset");
oConn.Open( filePath );

oRec.Open("select * from Anotaciones where (alumn = " + Session("uid") + ")",oConn,3,3);

//Mostrar todas las anotaciones de este usuario

while (!oRec.eof)
{
Response.write(oRec.Fields.Item.("texto").value); 

oRec.Move(1);
}

oRec.Close();
oConn.Close();
%>
</body>
</html>

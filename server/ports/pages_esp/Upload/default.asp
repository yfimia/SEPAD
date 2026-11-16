<%@ Language=JScript %>
<%
// Port a JScript de pages_esp/Upload/default.asp: archivos de descarga del módulo actual.
var course = parseInt(Session("course"), 10);
var UPLOAD_VPATH = "../../Courses/Course" + (isNaN(course) ? 0 : course) + "/descargas/";
var UPLOAD_LINK = "../../courses/course" + (isNaN(course) ? 0 : course) + "/descargas/";
var UPLOAD_ROOT = "../../";
// Solo usuarios autenticados (no el invitado) con un módulo seleccionado
var UPLOAD_ALLOWED = !isNaN(course) && Session("userID") && Session("userID") != 580;
%>
<!-- #include file="../../inc/uploadpage.inc" -->

<%@ Language=JScript %>
<%
// Port a JScript de pages_esp/Administracion/Upload/default.asp: descargas generales del sitio.
var UPLOAD_VPATH = "../../../downloads/";
var UPLOAD_LINK = "../../../downloads/";
var UPLOAD_ROOT = "../../../";
// Solo administradores (PermissionType 20)
var UPLOAD_ALLOWED = Session("permissionType") == 20;
%>
<!-- #include file="../../../inc/uploadpage.inc" -->

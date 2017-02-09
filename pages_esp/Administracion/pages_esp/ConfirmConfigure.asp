<%@ Language=JavaScript %>
<%
  Response.Expires = -1;
  Response.Buffer = true;  
%>

<!-- #include file='../js/user.inc' -->

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       
if (Session("PermissionType") == ADMINISTRATOR)
    {
      Application.Lock();
      Application("SQL")       = Request("SQL") + "";
      Application("dataPath")       = Request("DSN") + "";
      Application("filePath")       = Request("DSN") + "";

      Application("MailUser")       = Request("MailUser") + "";
      Application("MailServer")     = Request("MailServer") + "";
      Application("MailAddress")    = Request("MailAddress") + "";
      Application("MailPassword")   = Request("MailPassword") + "";
      Application("Port")   = Request("Port") + "";
      Application("Auth")   = Request("Auth") + "";
      Application("Subject")   = Request("Subject") + "";
      Application("Body")   = Request("Body") + "";
      Application("LogoImg")       = Request("LogoImg") + "";
      Application("HomePage")     = Request("HomePage") + "";
      Application("Downloads")     = Request("Downloads") + "";
      Application("ListSize")    = Request("ListSize") + "";
      Application("NewNews")   = Request("NewNews") + "";      
      Application.Unlock();
      
    
    var sXml = Application("Config");
    var oXmlDoc = Server.CreateObject("MICROSOFT.XMLDOM");
    oXmlDoc.async = false;
    oXmlDoc.load(Server.MapPath(sXml));

  
    //DBConn
    oXmlDoc.documentElement.selectSingleNode("/SysConfig/DBConnection/SQL").text = Application("SQL");
    oXmlDoc.documentElement.selectSingleNode("/SysConfig/DBConnection/DSN").text = Application("dataPath");

    //MailReply
    oXmlDoc.documentElement.selectSingleNode("/SysConfig/MailReply/User").text =   Application("MailUser");
    oXmlDoc.documentElement.selectSingleNode("/SysConfig/MailReply/Server").text =   Application("MailServer");
    oXmlDoc.documentElement.selectSingleNode("/SysConfig/MailReply/Address").text =   Application("MailAddress");
    oXmlDoc.documentElement.selectSingleNode("/SysConfig/MailReply/Password").text =   Application("MailPassword");

    oXmlDoc.documentElement.selectSingleNode("/SysConfig/MailReply/Port").text =   Application("Port");
    oXmlDoc.documentElement.selectSingleNode("/SysConfig/MailReply/Auth").text =   Application("Auth");
    oXmlDoc.documentElement.selectSingleNode("/SysConfig/MailReply/Subject").text =   Application("Subject");
    oXmlDoc.documentElement.selectSingleNode("/SysConfig/MailReply/Body").text =   Application("Body");
     
    //General
    oXmlDoc.documentElement.selectSingleNode("/SysConfig/General/Logo").text =   Application("LogoImg");
    oXmlDoc.documentElement.selectSingleNode("/SysConfig/General/Homepage").text =   Application("HomePage");
    oXmlDoc.documentElement.selectSingleNode("/SysConfig/General/Downloads").text =   Application("Downloads");
    oXmlDoc.documentElement.selectSingleNode("/SysConfig/General/MaxList").text =   Application("ListSize") + "";
    oXmlDoc.documentElement.selectSingleNode("/SysConfig/General/NewNews").text =   Application("NewNews") + "";

    oXmlDoc.save(Server.MapPath(sXml));

      Response.Redirect('configure.asp?uid=' + uid);     

%>

<HTML>
  <HEAD>
    <LINK href="../css/SepadCss1.css" rel="stylesheet" type="text/css" />
  </HEAD>
  <BODY>
<%

    }
    else
      Response.Redirect("errorpage.asp?tipo=Error&short=" + DONT_HAS_PERMISON_SHORT  + "&desc=" + DONT_HAS_PERMISON_TEXT);     
%>
    
    <CENTER><INPUT style="BACKGROUND-COLOR: darkkhaki; BORDER-BOTTOM-COLOR: darkkhaki; FONT-FAMILY: sans-serif; FONT-SIZE: small; FONT-STYLE: normal; HEIGHT: 27px; WIDTH: 90px"  type=button value=Regresar onclick ="history.back(-1)"></CENTER>
  </BODY>
</HTML>
<%
  }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
  //Response.Write(Request.QueryString.Item("uid"));
%>
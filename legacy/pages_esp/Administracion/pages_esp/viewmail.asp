<%@ Language=JavaScript %>

<%Response.Expires = -1%>
<!-- #include file='../js/user.inc' -->
<!-- #include file='inc/viewmail.inc' -->

<%
    if (Request.QueryString.Item("uid") + "" == Session("uid"))
    {    
      var uid = Request.QueryString.Item("uid") + "";       

%>


<%
   var Ind;
   
   if (Request.QueryString.Count  > 0 ) 
     {
      Ind           = Request.QueryString.Item("ID");
      Tmp           = Session("MailContent");                       
       
      HTML = "<TABLE Class='MessageTable' cellspacing='0' cellpadding='0' Height='60%'>" +
              "<TR>" + 
                "<TD  height='5%' Class='MainTD' Width='10%'>" + TEXT1 + ":</TD><TD height='5%' Class='DatTD' Width='90%'>" + Tmp[Ind].From + "</TD>" +
              "</TR>"  + 
              "<TR>" + 
                "<TD Class='MainTD'' height='5%' Width='10%'>" + TEXT2 + ":</TD><TD height='5%' Class='DatTD'' Width='90%'>" + Tmp[Ind].Subject + "</TD>" +
              "</TR>" +              
              "<TR>" + 
                "<TD colspan='2' height='50%'>" +
                "<TABLE cellspacing='0' cellpadding='0' Class='MainTable' height='100%'>" +
                  "<TR>" + 
                    "<TD Width='100%' height='100%'>" + 
                   "<TEXTAREA Class='MArea'>" + Tmp[Ind].Content +"</TEXTAREA>" + 
                "</TD>" +                 
                "</TR>" + 
                "</TABLE>" +
                "</TD>" + 
              "</TR>" + 
              "<TR Height='5%'>" + 
                "<TD colspan='2'>" +
                "<TABLE width='100%' Height='100%' cellspacing='0' cellpadding='0'>" +
                  "<TR>" + 
                    "<TD Width='100%' Height='100%' >" + 
                      "<INPUT Type='Button' Value='Volver' onclick='return Ionclick()'>" + 
                "</TD>" +                 
                "</TR>" + 
                "</TABLE>" +
                "</TD>" + 
              "</TR>" + 
             "</TABLE>";
    
     if (Session("MailContent") != null) {Session("MailContent") = null};
     if (Session("MessageID")!= null) {Session("MessageID") = null};
    } 
      
%>

<HTML>
<HEAD>
<link REL="stylesheet" TYPE="text/css" HREF="../css/<%=Session("skin")%>/SepadCss.css" />

<Script languaje="jscript">
   function Ionclick()
    {
     <% if (Request.QueryString("Flag") ==1){ %>
        location.href = "../beta/mail.asp";
     <%} else{ %> 
        location.href = "mailmanager.asp?uid=<%=uid%>";       
     <%}%>
    }
</Script>
</HEAD>

<BODY>
<%=HTML%>
<BODY>
</HTML>
<%
  }
 else
   Response.Redirect("errorpage.asp?tipo=Error&short=" + SESSION_TIMEOUT_SHORT  + "&desc=" + SESSION_TIMEOUT_TEXT);
  //Response.Write(Request.QueryString.Item("uid"));
%>

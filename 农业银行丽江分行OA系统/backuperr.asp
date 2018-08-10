<%@ Language=VBScript %>
<%
dim errno,errstr
errno=Request.QueryString("err")
errstr=""
select case cstr(errno)
case "18456"
errstr="administrators or password error!"
case "20482"
errstr="server name error or server cannot connect!"
case "911"
errstr="database not found!"
case "15026"
errstr="server path not found!"
case "3201"
errstr="server path not found!"
case "3254"
errstr="restore from file lawlessness!"
case else
errstr="unknown error! retry later please!"
end select
%>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<META NAME="GENERATOR" Content="Microsoft FrontPage 4.0">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub btnret_onclick
history.back
End Sub

-->
</SCRIPT>
<title></title>
</HEAD>
<body class="bg_frame_up" BACKGROUND="images/main_bg.gif">
<p align=center><font color=#006666><%=errstr%></font></p><p align=center><input id=btnret name=btnret type=button value=Return style="font-family: Arial; font-size: 9pt"></p>
</BODY>
</HTML> 



<%@ Language=VBScript %>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<META NAME="GENERATOR" Content="Microsoft FrontPage 4.0">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub btnbak_onclick
if frmbak.txtsvr.value="" then
window.alert("'Server Name' is empty!")
frmbak.txtsvr.focus
exit sub
end if
if frmbak.txtuid.value="" then
window.alert("'Administrators' is empty!")
frmbak.txtuid.focus
exit sub
end if
if frmbak.txtdb.value="" then
window.alert("'Database' is empty!")
frmbak.txtdb.focus
exit sub
end if
if frmbak.txtto.value="" then
window.alert("'Backup To' is empty!")
frmbak.txtto.focus
exit sub
end if
frmbak.submit
End Sub

-->
</SCRIPT>
<link rel="stylesheet" href="../../sheets/B2BStyle.css">
</HEAD>
<BODY BACKGROUND="images/main_bg.gif">
<form action="backupdbsave.asp" method="post" id=frmbak name=frmbak> <body class="bg_frame_up"> 
<p class=heading align=center><STRONG><FONT size=5>数 据 备 份</FONT></STRONG> </p><P align=center> 
<div align="center"> <center> <table width="60%" cellpadding=1 cellspacing=1 border=0 align=center> 
<tr> <td class=TD_Mand_FN align="middle" height="35" width="40%">服务器名:</td><td class=TD_Mand_F height="35" width="59%"> 
<INPUT id=txtsvr name=txtsvr style="FONT-SIZE: 9pt; FONT-FAMILY: Arial" ></td></tr> 
<tr> <td class=TD_Mand_FN align="middle" height="35" width="40%">管 理 员:</td><td class=TD_Mand_F height="35" width="59%"> 
<INPUT id=txtuid name=txtuid style="FONT-SIZE: 9pt; FONT-FAMILY: Arial" ></td></tr> 
<tr> <td class=TD_Mand_FN align="middle" height="35" width="40%">密&nbsp;&nbsp;&nbsp;&nbsp;码:</td><td class=TD_Mand_F height="35" width="59%"> 
<INPUT id=txtpwd name=txtpwd type=password style="FONT-SIZE: 9pt; FONT-FAMILY: Arial"></td></tr> 
<tr> <td class=TD_Mand_FN align="middle" height="35" width="40%">数据库名:</td><td class=TD_Mand_F height="35" width="59%"> 
<p align="left"> <INPUT id=txtdb name=txtdb style="FONT-SIZE: 9pt; FONT-FAMILY: Arial" ></p></td></tr> 
<TR> <td class=TD_Mand_FN align="middle" height="35" width="40%">备 份 至:<br> <u>(服务器路径)</u></td><td class=TD_Mand_F height="35" width="59%"> 
<input id=txtto name=txtto style="FONT-SIZE: 9pt; FONT-FAMILY: Arial"></td></TR> 
</table></center></div><P align=center><input id=btnbak name=btnbak type=button value="Start Backup" style="FONT-SIZE: 9pt; WIDTH: 105px; FONT-FAMILY: Arial; HEIGHT: 22px"></P></body> 
</form> 
</HTML> 

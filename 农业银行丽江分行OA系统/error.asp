<HTML>
<HEAD>
<TITLE> New Document </TITLE>
</HEAD>
<style>
<!-- 
td{font-size:10.5pt}

-->
</style>

<BODY bgColor=#ffffff leftMargin=0 
style="BACKGROUND-ATTACHMENT: scroll; BACKGROUND-IMAGE: url(images/main_bg.gif); BACKGROUND-POSITION: left bottom; BACKGROUND-REPEAT: no-repeat" 
topMargin=0>
<TABLE width=100% height=100%>
<TR>

    <TD valign=middle align=center><table width="400" border="0" height="300" align=center> 
<tr align="center"> 
<td> 
<form method="post" action=""> 
<table border="1" bordercolorlight="000000" bordercolordark="FFFFFF" cellspacing="0" bgcolor="E0E0E0"> 
<tr> 
<td> 
<table border="0" bgcolor="#0066CC" cellspacing="0" cellpadding="2" width="350"> 
<tr> 
<td width="342"><font color="FFFFFF">¤出错提示</font></td> 
<td width="18">&nbsp; 
</td> 
</tr> 
</table> 
<table border="0" width="350" cellpadding="4"> 
<tr> 
<td width="59" align="center" valign="top"><font face="Wingdings" color="#FF0000" style="font-size:32pt">L</font></td> 
<td width="269"> 
<%
 if request("id")="admin" then
%> 
<p>　　您没有足够权限进行此操作，如有任何问题请和<a href="mailto:lj_lw@ynmail.com">管理员联系</A>！</p>
<%end if%> 
</td> 
</tr> 
<tr> 
<td colspan="2" align="center" valign="top"> 
<input type="button" name="ok" value="　确 定　" onclick="location.href='tongzhi.asp'"> 
</td> 
</tr> 
</table> 
</td> 
</tr> 
</table> 
</form> 
</td> 
</tr> 
</table> </TD>
</TR>
</TABLE>

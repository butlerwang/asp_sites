<!--#INCLUDE FILE="data.asp" -->
<!--#INCLUDE FILE="check.asp" -->
<%
if Session("Urule")<>"a" then
response.redirect("error.asp?id=admin")
response.end
end if
Set rs= Server.CreateObject("ADODB.Recordset") 
strSql="select * from jhtdata where id="&request("id")
rs.open strSql,Conn,1,1 
%>

<link rel="stylesheet" href="oa.css">
<body leftmargin="0" topmargin="20" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellspacing="1" cellpadding="2"> <tr > <td class="heading"> 
<div align="center"> <center> <table width="81%" border="0" cellspacing="0" cellpadding="0" bgcolor="#000000" bordercolorlight="#000000"> 
<tr> <td width="2%" align="right"></td><td align="left" height="25"> <p align="center"><font color="#FFFFFF"><b>收 
取 文 件</b></font></p></td><td width="3%"></td></tr> </table></center></div></td></tr> 
</table><div align="center"> <table width="80%" border="1" cellspacing="0" cellpadding="0" bordercolordark="#FFFFFF" bordercolor="#FFFFFF" bordercolorlight="#000000"> 
<tr> <td width="17%" valign="top"> <p align="right">单位名称:</p></td><td width="83%"> 
<%=rs("部门")%> </td></tr> <tr> <td width="17%" valign="top" height="6"> <p align="right">文件报送人:</p></td><td width="83%" height="6"> 
<%=rs("真实姓名")%> </td></tr> <tr> <td width="17%"  valign="top" height="16"> <p align="right">报送时间:</p></td><td width="83%" height="16"> 
<%=rs("时间")%> </td></tr> <tr> <td width="17%"  valign="top" height="16"> <p align="right">IP地址:</p></td><td width="83%" height="16"> 
<%=rs("IP")%> </td></tr> <tr> <td width="17%"  valign="top" height="16"> <p align="right"><BR>附 
加 &nbsp;<BR><BR>说 明 &nbsp<BR><BR></p></td><td width="83%" height="16" valign=top> 
<%=rs("标题")%> </td></tr> <tr> <td width="17%"  valign="top"> <p align="right">下载该文件: 
</td><td width="83%"> <A HREF="<%=rs("链接")%>">点击这里下载</A> </td></tr> </table></div><div align="center"> 
<a  href="javaScript:window.close()"><img   border="0" src="images/close_1.gif"></a> 
</div>     
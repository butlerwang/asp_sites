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
<tr> <td width="2%" align="right"></td><td align="left" height="25"> <p align="center"><font color="#FFFFFF"><b>�� 
ȡ �� ��</b></font></p></td><td width="3%"></td></tr> </table></center></div></td></tr> 
</table><div align="center"> <table width="80%" border="1" cellspacing="0" cellpadding="0" bordercolordark="#FFFFFF" bordercolor="#FFFFFF" bordercolorlight="#000000"> 
<tr> <td width="17%" valign="top"> <p align="right">��λ����:</p></td><td width="83%"> 
<%=rs("����")%> </td></tr> <tr> <td width="17%" valign="top" height="6"> <p align="right">�ļ�������:</p></td><td width="83%" height="6"> 
<%=rs("��ʵ����")%> </td></tr> <tr> <td width="17%"  valign="top" height="16"> <p align="right">����ʱ��:</p></td><td width="83%" height="16"> 
<%=rs("ʱ��")%> </td></tr> <tr> <td width="17%"  valign="top" height="16"> <p align="right">IP��ַ:</p></td><td width="83%" height="16"> 
<%=rs("IP")%> </td></tr> <tr> <td width="17%"  valign="top" height="16"> <p align="right"><BR>�� 
�� &nbsp;<BR><BR>˵ �� &nbsp<BR><BR></p></td><td width="83%" height="16" valign=top> 
<%=rs("����")%> </td></tr> <tr> <td width="17%"  valign="top"> <p align="right">���ظ��ļ�: 
</td><td width="83%"> <A HREF="<%=rs("����")%>">�����������</A> </td></tr> </table></div><div align="center"> 
<a  href="javaScript:window.close()"><img   border="0" src="images/close_1.gif"></a> 
</div>     
<!--#INCLUDE FILE="data.asp" -->
<!--#INCLUDE FILE="check.asp" -->
<%
Set rs= Server.CreateObject("ADODB.Recordset") 
strSql="select * from userinfo where userid="&session("Uid")
rs.open strSql,Conn,1,1 
if rs.eof then
response.write "no record"
end if

%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="oa.css">
</head>
<title>���˵���</title>  
<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0" style="background-image: url('images/main_bg.gif'); background-attachment: scroll; background-repeat: no-repeat; background-position: left bottom">
<BR><table border="0" cellpadding="0" cellspacing="0" width="95%" bordercolorlight=#000000 bordercolordark=#ffffff align=right> 
<tr align="center"> <td><b>¼��ʱ�䣺</b><%=session("time")%></td><td><b>�޸�ʱ�䣺</b><%=rs("Ltime")%></td></tr> 
</table><BR> <table border="1" cellpadding="0" cellspacing="0" width="95%" bordercolorlight=#000000 bordercolordark=#ffffff align=right> 
<tr> <td align="center" width="15%"><b>��&nbsp;&nbsp;&nbsp;&nbsp;��</b></td><td width="30%">&nbsp;<%=session("Rname")%></td><td align="center" width="15%"><b>��&nbsp;��&nbsp;��</b></td><td width="25%">&nbsp;<%if check("0")="no" and session("Uid")<>rs("userid") then response.write "����" else response.write rs("Uname")%></td><td width="80" height="100" rowspan="5" align="center" valign=center><%if rs("havephoto")=false then%>��<BR>��<BR>Ƭ<%else%> 
<img src="showpic.asp?id=<%=rs("id")%>" width="80" height="100" border="0"><%end if%> 
</td></tr> <tr> <td align="center"><b>��&nbsp;&nbsp;&nbsp;&nbsp;��</b></td><td>&nbsp;<%=rs("sex")%></td><td align="center"><b>��&nbsp;&nbsp;&nbsp;&nbsp;��</b></td><td>&nbsp;<%=rs("nation")%></td></tr> 
<tr> <td align="center"><b>��������</b></td><td>&nbsp;<%=Session("Upart")%></td><td align="center"><b>ְ&nbsp;&nbsp;&nbsp;&nbsp;��</b></td><td>&nbsp;<%=rs("duty")%></td></tr> 
<tr> <td align="center"><b>ְ&nbsp;&nbsp;&nbsp;&nbsp;��</b></td><td>&nbsp;<%=rs("grade")%></td><td align="center"><b>��������</b></td><td>&nbsp;<%=rs("birthday")%></td></tr> 
<tr> <td align="center"><b>������ò</b></td><td>&nbsp;<%=rs("polity")%></td><td align="center"><b>����״��</b></td><td>&nbsp;<%=rs("health")%></td></tr> 
<tr> <td align="center" height="20"><b>��&nbsp;&nbsp;&nbsp;&nbsp;��</b></td><td>&nbsp;<%=rs("Nplace")%></td><td align="center"><b>��&nbsp;&nbsp;&nbsp;&nbsp;��</b></td><td colspan="2">&nbsp;<%=rs("weight")%></td></tr> 
<tr> <td height="20" align="center"><b>���֤��</b></td><td>&nbsp;<%=rs("idcard")%></td><td align="center"><b>��&nbsp;&nbsp;&nbsp;&nbsp;��</b></td><td colspan="2">&nbsp;<%=rs("height")%></td></tr> 
<tr> <td align="center" height="20"><b>����״��</b></td><td>&nbsp;<%=rs("marriage")%></td><td align="center"><b>��ҵԺУ</b></td><td colspan="2">&nbsp;<%=rs("Fschool")%></td></tr> 
<tr> <td align="center" height="20"><b>���˳ɷ�</b></td><td>&nbsp;<%=rs("member")%></td><td align="center"><b>ר&nbsp;&nbsp;&nbsp;&nbsp;ҵ</b></td><td colspan="2">&nbsp;<%=rs("speciality")%></td></tr> 
<tr> <td height="20" align="center"><b>��&nbsp;&nbsp;&nbsp;&nbsp;��</b></td><td>&nbsp;<%=rs("length")%></td><td align="center"><b>ѧ&nbsp;&nbsp;&nbsp;&nbsp;��</b></td><td colspan="2">&nbsp;<%=rs("study")%></td></tr> 
<tr> <td height="20" align="center"><b>��������</b></td><td>&nbsp;<%=rs("foreign")%></td><td align="center"><b>����ˮƽ</b></td><td colspan="2">&nbsp;<%=rs("Elevel")%></td></tr> 
<tr> <td height="20" align="center"><b>���������</b></td><td>&nbsp;<%=rs("Clevel")%></td><td align="center"><b>�������ڵ�</b></td><td colspan="2">&nbsp;<%=rs("Hplace")%></td></tr> 
<tr> <td height="20" align="center"><b>QQ����</b></td><td>&nbsp;<%=rs("QQ")%></td><td align="center"><b>EMAIL</b></td><td colspan="2">&nbsp;<%=Session("email")%></td></tr> 
<tr> <td height="20" align="center"><b>���õ绰</b></td><td>&nbsp;<%=Session("tel")%></td><td align="center"><b>�ֻ�����</b></td><td colspan="2">&nbsp;<%=session("mobile")%></td></tr> 
<tr> <td height="20" align="center"> <b>��������</b> </td><td>&nbsp;<%=rs("call")%></td><td align="center">&nbsp;</td><td colspan="2">&nbsp;</td></tr> 
<tr> <td height="20" align="center"><b>��&nbsp;ס&nbsp;ַ</b></td><td colspan="4">&nbsp;<%=rs("place")%></td></tr> 
<tr> <td height="20" align="center"><b>����ר��<br> �Լ�����</b></td><td colspan="4">&nbsp;<%=rs("love")%></td></tr> 
<tr> <td height="20" align="center"><b>��������<br> �����ֽ�<br> ���ʹ���</b></td><td colspan="4">&nbsp;<%=rs("award")%></td></tr> 
<tr> <td height="20" align="center"><b>��������</b></td><td colspan="4">&nbsp;<%=rs("experience")%></td></tr> 
<tr> <td height="20" align="center"><b>��ͥ���</b></td><td colspan="4">&nbsp;<%=rs("family")%></td></tr> 
<tr> <td height="20" align="center"><b>��&nbsp;&nbsp;&nbsp;&nbsp;��<br> ��ϵ��ʽ</b></td><td colspan="4">&nbsp;<%=rs("contact")%></td></tr> 
<tr> <td height="20" align="center"><b>��&nbsp;&nbsp;&nbsp;&nbsp;ע</b></td><td colspan="4">&nbsp;<%=rs("remark")%></td></tr> 
<form method="post" action="savelogo.asp" name="reg" enctype="multipart/form-data"> 
<td height="20" align="center" colspan="5"> <input type="file" name="file" size="12"><INPUT TYPE="submit" value="�ϴ�/�޸���Ƭ">&nbsp; 
<input type="button" value="�༭���˵���" name="submit2" onclick="location.href='archives_edit.asp'"> 
</td></FORM></table>

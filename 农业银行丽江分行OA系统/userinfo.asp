<!--#INCLUDE FILE="data.asp" -->
<!--#INCLUDE FILE="check.asp" -->
<%
Set rs1= Server.CreateObject("ADODB.Recordset") 
strSql1="select ����,����,mobile,����,�绰,��½IP,times,Utime from user where id="&request("id")
rs1.open strSql1,Conn,1,1 
if rs1.eof then
response.write "no record1"
end if

Set rs= Server.CreateObject("ADODB.Recordset") 
strSql="select * from userinfo where userid="&request("id")
rs.open strSql,Conn,1,1 
if rs.eof then
response.write "no record"
end if


check=split(rs("check"), ",", -1, 1)
%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="oa.css">
</head>
<title>���˵���</title>  
<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0" style="background-image: url('images/main_bg.gif'); background-attachment: scroll; background-repeat: no-repeat; background-position: left bottom">

<table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight=#000000 bordercolordark=#ffffff align=right>
    <tr> 
    <td colspan=5 align=center><FONT COLOR="blue">
���û�����½ϵͳ<FONT COLOR="red"><%=rs1("times")%></FONT>�Σ����һ�ε�½ʱ����:<FONT COLOR="red"><%=left(rs1("Utime"),4)%>/<%=mid(rs1("Utime"),5,2)%>/<%=mid(rs1("Utime"),7,2)%>&nbsp;<%=mid(rs1("Utime"),9,2)%>:<%=right(rs1("Utime"),2)%></FONT>����½IP:<FONT COLOR="red"><%=rs1("��½IP")%></FONT></FONT>
	</td>
  </tr>

<tr> 
    <td align="center" width="15%"><b>��&nbsp;&nbsp;&nbsp;&nbsp;��</b></td>
    <td width="30%">&nbsp;<%=rs1("����")%></td>
    <td align="center" width="15%"><b>��&nbsp;��&nbsp;��</b></td>
    <td width="25%">&nbsp;<%if check("0")="no" and session("Urule")<>"a" then response.write "����" else response.write rs("Uname")%></td>
    <td width="80" height="100" rowspan="5" align="center" valign=center><%if rs("havephoto")=false then%>��<BR>��<BR>Ƭ<%else%> <img src="showpic.asp?id=<%=rs("id")%>" width="80" height="100" border="0"><%end if%> 
    </td>
  </tr>
  <tr> 
    <td align="center"><b>��&nbsp;&nbsp;&nbsp;&nbsp;��</b></td>
    <td>&nbsp;<%if check("1")="no" and session("Urule")<>"a" then Response.Write "����" else Response.Write rs("sex")%></td>
    </td>
    <td align="center"><b>��&nbsp;&nbsp;&nbsp;&nbsp;��</b></td>
    <td>&nbsp;<%if check("2")="no" and session("Urule")<>"a" then response.write "����" else response.write rs("nation")%></td>
  </tr>
  <tr> 
    <td align="center"><b>��������</b></td>
    <td>&nbsp;<%=Session("Upart")%></td>
    <td align="center"><b>ְ&nbsp;&nbsp;&nbsp;&nbsp;��</b></td>
    <td>&nbsp;<%if check("3")="no" and session("Urule")<>"a" then response.write "����" else response.write rs("duty")%></td>
  </tr>
  <tr> 
    <td align="center"><b>ְ&nbsp;&nbsp;&nbsp;&nbsp;��</b></td>
    <td>&nbsp;<%if check("4")="no" and session("Urule")<>"a" then response.write "����" else response.write rs("grade")%></td>
    <td align="center"><b>��������</b></td>
    <td>&nbsp;<%if check("5")="no" and session("Urule")<>"a" then response.write "����" else response.write rs("birthday")%></td>
  </tr>
  <tr> 
    <td align="center"><b>������ò</b></td>
    <td>&nbsp;<%if check("6")="no" and session("Urule")<>"a" then response.write "����" else response.write rs("polity")%></td>
    <td align="center"><b>����״��</b></td>
    <td>&nbsp;<%if check("7")="no" and session("Urule")<>"a" then response.write "����" else response.write rs("health")%></td>
  </tr>
  <tr> 
    <td align="center" height="20"><b>��&nbsp;&nbsp;&nbsp;&nbsp;��</b></td>
    <td>&nbsp;<%if check("8")="no" and session("Urule")<>"a" then response.write "����" else response.write rs("Nplace")%></td>
    <td align="center"><b>��&nbsp;&nbsp;&nbsp;&nbsp;��</b></td>
    <td colspan="2">&nbsp;<%if check("9")="no" and session("Urule")<>"a" then response.write "����" else response.write rs("weight")%></td>
  </tr>
  <tr> 
    <td height="20" align="center"><b>���֤��</b></td>
    <td>&nbsp;<%if check("10")="no" and session("Urule")<>"a" then response.write "����" else response.write rs("idcard")%></td>
    <td align="center"><b>��&nbsp;&nbsp;&nbsp;&nbsp;��</b></td>
    <td colspan="2">&nbsp;<%if check("11")="no" and session("Urule")<>"a" then response.write "����" else response.write rs("height")%></td>
  </tr>
  <tr> 
    <td align="center" height="20"><b>����״��</b></td>
    <td>&nbsp;<%if check("12")="no" and session("Urule")<>"a" then response.write "����" else response.write rs("marriage")%></td>
    <td align="center"><b>��ҵԺУ</b></td>
    <td colspan="2">&nbsp;<%if check("13")="no" and session("Urule")<>"a" then response.write "����" else response.write rs("Fschool")%></td>
  </tr>
  <tr> 
    <td align="center" height="20"><b>���˳ɷ�</b></td>
    <td>&nbsp;<%if check("14")="no" and session("Urule")<>"a" then response.write "����" else response.write rs("member")%></td>
    <td align="center"><b>ר&nbsp;&nbsp;&nbsp;&nbsp;ҵ</b></td>
    <td colspan="2">&nbsp;<%if check("15")="no" and session("Urule")<>"a" then response.write "����" else response.write rs("speciality")%></td>
  </tr>
  <tr> 
    <td height="20" align="center"><b>��&nbsp;&nbsp;&nbsp;&nbsp;��</b></td>
    <td>&nbsp;<%if check("16")="no" and session("Urule")<>"a" then response.write "����" else response.write rs("length")%></td>
    <td align="center"><b>ѧ&nbsp;&nbsp;&nbsp;&nbsp;��</b></td>
    <td colspan="2">&nbsp;<%if check("17")="no" and session("Urule")<>"a" then response.write "����" else response.write rs("study")%></td>
  </tr>
  <tr> 
    <td height="20" align="center"><b>��������</b></td>
    <td>&nbsp;<%if check("18")="no" and session("Urule")<>"a" then response.write "����" else response.write rs("foreign")%></td>
    <td align="center"><b>����ˮƽ</b></td>
    <td colspan="2">&nbsp;<%if check("19")="no" and session("Urule")<>"a" then response.write "����" else response.write rs("Elevel")%></td>
  </tr>
  <tr> 
    <td height="20" align="center"><b>���������</b></td>
    <td>&nbsp;<%if check("20")="no" and session("Urule")<>"a" then response.write "����" else response.write rs("Clevel")%></td>
    <td align="center"><b>�������ڵ�</b></td>
    <td colspan="2">&nbsp;<%if check("21")="no" and session("Urule")<>"a" then response.write "����" else response.write rs("Hplace")%></td>
  </tr>
  <tr>
    <td height="20" align="center"><b>QQ����</b></td>
    <td>&nbsp;<%if check("22")="no" and session("Urule")<>"a" then response.write "����" else response.write rs("QQ")%></td>
    <td align="center"><b>EMAIL</b></td>
    <td colspan="2">&nbsp;<%=rs1("����")%></td>
  </tr>
  <tr>
    <td height="20" align="center"><b>���õ绰</b></td>
    <td>&nbsp;<%=rs1("�绰")%></td>
    <td align="center"><b>�ֻ�����</b></td>
    <td colspan="2">&nbsp;<%=rs1("mobile")%></td>
  </tr>
  <tr>
    <td height="20" align="center"> <b>��������</b> </td>
    <td>&nbsp;<%if check("23")="no" and session("Urule")<>"a" then response.write "����" else response.write rs("call")%></td>
    <td align="center">&nbsp;</td>
    <td colspan="2">&nbsp;</td>
  </tr>
  <tr> 
    <td height="20" align="center"><b>��&nbsp;ס&nbsp;ַ</b></td>
    <td colspan="4">&nbsp;<%if check("24")="no" and session("Urule")<>"a" then response.write "����" else response.write rs("place")%></td>
  </tr>
  <tr> 
    <td height="20" align="center"><b>����ר��<br>
      �Լ�����</b></td>
    <td colspan="4">&nbsp;<%if check("25")="no" and session("Urule")<>"a" then response.write "����" else response.write rs("love")%></td>
  </tr>
  <tr> 
    <td height="20" align="center"><b>��������<br>
      �����ֽ�<br>
      ���ʹ���</b></td>
    <td colspan="4">&nbsp;<%if check("26")="no" and session("Urule")<>"a" then response.write "����" else response.write rs("award")%></td>
  </tr>
  <tr> 
    <td height="20" align="center"><b>��������</b></td>
    <td colspan="4">&nbsp;<%if check("27")="no" and session("Urule")<>"a" then response.write "����" else response.write rs("experience")%></td>
  </tr>
  <tr> 
    <td height="20" align="center"><b>��ͥ���</b></td>
    <td colspan="4">&nbsp;<%if check("28")="no" and session("Urule")<>"a" then response.write "����" else response.write rs("family")%></td>
  </tr>
  <tr> 
    <td height="20" align="center"><b>��&nbsp;&nbsp;&nbsp;&nbsp;��<br>
      ��ϵ��ʽ</b></td>
    <td colspan="4">&nbsp;<%if check("29")="no" and session("Urule")<>"a" then response.write "����" else response.write rs("contact")%></td>
  </tr>
  <tr> 
    <td height="20" align="center"><b>��&nbsp;&nbsp;&nbsp;&nbsp;ע</b></td>
    <td colspan="4">&nbsp;<%if check("30")="no" and session("Urule")<>"a" then response.write "����" else response.write rs("remark")%></td>
  </tr>
</table>

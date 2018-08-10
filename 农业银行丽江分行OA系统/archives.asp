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
<title>个人档案</title>  
<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0" style="background-image: url('images/main_bg.gif'); background-attachment: scroll; background-repeat: no-repeat; background-position: left bottom">
<BR><table border="0" cellpadding="0" cellspacing="0" width="95%" bordercolorlight=#000000 bordercolordark=#ffffff align=right> 
<tr align="center"> <td><b>录入时间：</b><%=session("time")%></td><td><b>修改时间：</b><%=rs("Ltime")%></td></tr> 
</table><BR> <table border="1" cellpadding="0" cellspacing="0" width="95%" bordercolorlight=#000000 bordercolordark=#ffffff align=right> 
<tr> <td align="center" width="15%"><b>姓&nbsp;&nbsp;&nbsp;&nbsp;名</b></td><td width="30%">&nbsp;<%=session("Rname")%></td><td align="center" width="15%"><b>曾&nbsp;用&nbsp;名</b></td><td width="25%">&nbsp;<%if check("0")="no" and session("Uid")<>rs("userid") then response.write "保密" else response.write rs("Uname")%></td><td width="80" height="100" rowspan="5" align="center" valign=center><%if rs("havephoto")=false then%>无<BR>照<BR>片<%else%> 
<img src="showpic.asp?id=<%=rs("id")%>" width="80" height="100" border="0"><%end if%> 
</td></tr> <tr> <td align="center"><b>性&nbsp;&nbsp;&nbsp;&nbsp;别</b></td><td>&nbsp;<%=rs("sex")%></td><td align="center"><b>民&nbsp;&nbsp;&nbsp;&nbsp;族</b></td><td>&nbsp;<%=rs("nation")%></td></tr> 
<tr> <td align="center"><b>所属部门</b></td><td>&nbsp;<%=Session("Upart")%></td><td align="center"><b>职&nbsp;&nbsp;&nbsp;&nbsp;务</b></td><td>&nbsp;<%=rs("duty")%></td></tr> 
<tr> <td align="center"><b>职&nbsp;&nbsp;&nbsp;&nbsp;称</b></td><td>&nbsp;<%=rs("grade")%></td><td align="center"><b>出生日期</b></td><td>&nbsp;<%=rs("birthday")%></td></tr> 
<tr> <td align="center"><b>政治面貌</b></td><td>&nbsp;<%=rs("polity")%></td><td align="center"><b>健康状况</b></td><td>&nbsp;<%=rs("health")%></td></tr> 
<tr> <td align="center" height="20"><b>籍&nbsp;&nbsp;&nbsp;&nbsp;贯</b></td><td>&nbsp;<%=rs("Nplace")%></td><td align="center"><b>体&nbsp;&nbsp;&nbsp;&nbsp;重</b></td><td colspan="2">&nbsp;<%=rs("weight")%></td></tr> 
<tr> <td height="20" align="center"><b>身份证号</b></td><td>&nbsp;<%=rs("idcard")%></td><td align="center"><b>身&nbsp;&nbsp;&nbsp;&nbsp;高</b></td><td colspan="2">&nbsp;<%=rs("height")%></td></tr> 
<tr> <td align="center" height="20"><b>婚姻状况</b></td><td>&nbsp;<%=rs("marriage")%></td><td align="center"><b>毕业院校</b></td><td colspan="2">&nbsp;<%=rs("Fschool")%></td></tr> 
<tr> <td align="center" height="20"><b>本人成分</b></td><td>&nbsp;<%=rs("member")%></td><td align="center"><b>专&nbsp;&nbsp;&nbsp;&nbsp;业</b></td><td colspan="2">&nbsp;<%=rs("speciality")%></td></tr> 
<tr> <td height="20" align="center"><b>工&nbsp;&nbsp;&nbsp;&nbsp;龄</b></td><td>&nbsp;<%=rs("length")%></td><td align="center"><b>学&nbsp;&nbsp;&nbsp;&nbsp;历</b></td><td colspan="2">&nbsp;<%=rs("study")%></td></tr> 
<tr> <td height="20" align="center"><b>外语语种</b></td><td>&nbsp;<%=rs("foreign")%></td><td align="center"><b>外语水平</b></td><td colspan="2">&nbsp;<%=rs("Elevel")%></td></tr> 
<tr> <td height="20" align="center"><b>计算机能力</b></td><td>&nbsp;<%=rs("Clevel")%></td><td align="center"><b>户口所在地</b></td><td colspan="2">&nbsp;<%=rs("Hplace")%></td></tr> 
<tr> <td height="20" align="center"><b>QQ号码</b></td><td>&nbsp;<%=rs("QQ")%></td><td align="center"><b>EMAIL</b></td><td colspan="2">&nbsp;<%=Session("email")%></td></tr> 
<tr> <td height="20" align="center"><b>常用电话</b></td><td>&nbsp;<%=Session("tel")%></td><td align="center"><b>手机号码</b></td><td colspan="2">&nbsp;<%=session("mobile")%></td></tr> 
<tr> <td height="20" align="center"> <b>传呼号码</b> </td><td>&nbsp;<%=rs("call")%></td><td align="center">&nbsp;</td><td colspan="2">&nbsp;</td></tr> 
<tr> <td height="20" align="center"><b>现&nbsp;住&nbsp;址</b></td><td colspan="4">&nbsp;<%=rs("place")%></td></tr> 
<tr> <td height="20" align="center"><b>个人专长<br> 以及爱好</b></td><td colspan="4">&nbsp;<%=rs("love")%></td></tr> 
<tr> <td height="20" align="center"><b>本人曾受<br> 过何种奖<br> 励和处分</b></td><td colspan="4">&nbsp;<%=rs("award")%></td></tr> 
<tr> <td height="20" align="center"><b>工作经历</b></td><td colspan="4">&nbsp;<%=rs("experience")%></td></tr> 
<tr> <td height="20" align="center"><b>家庭情况</b></td><td colspan="4">&nbsp;<%=rs("family")%></td></tr> 
<tr> <td height="20" align="center"><b>本&nbsp;&nbsp;&nbsp;&nbsp;人<br> 联系方式</b></td><td colspan="4">&nbsp;<%=rs("contact")%></td></tr> 
<tr> <td height="20" align="center"><b>备&nbsp;&nbsp;&nbsp;&nbsp;注</b></td><td colspan="4">&nbsp;<%=rs("remark")%></td></tr> 
<form method="post" action="savelogo.asp" name="reg" enctype="multipart/form-data"> 
<td height="20" align="center" colspan="5"> <input type="file" name="file" size="12"><INPUT TYPE="submit" value="上传/修改照片">&nbsp; 
<input type="button" value="编辑个人档案" name="submit2" onclick="location.href='archives_edit.asp'"> 
</td></FORM></table>

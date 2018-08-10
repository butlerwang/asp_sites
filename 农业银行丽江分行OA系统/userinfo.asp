<!--#INCLUDE FILE="data.asp" -->
<!--#INCLUDE FILE="check.asp" -->
<%
Set rs1= Server.CreateObject("ADODB.Recordset") 
strSql1="select 姓名,部门,mobile,信箱,电话,登陆IP,times,Utime from user where id="&request("id")
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
<title>个人档案</title>  
<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0" style="background-image: url('images/main_bg.gif'); background-attachment: scroll; background-repeat: no-repeat; background-position: left bottom">

<table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight=#000000 bordercolordark=#ffffff align=right>
    <tr> 
    <td colspan=5 align=center><FONT COLOR="blue">
该用户共登陆系统<FONT COLOR="red"><%=rs1("times")%></FONT>次，最后一次登陆时间是:<FONT COLOR="red"><%=left(rs1("Utime"),4)%>/<%=mid(rs1("Utime"),5,2)%>/<%=mid(rs1("Utime"),7,2)%>&nbsp;<%=mid(rs1("Utime"),9,2)%>:<%=right(rs1("Utime"),2)%></FONT>、登陆IP:<FONT COLOR="red"><%=rs1("登陆IP")%></FONT></FONT>
	</td>
  </tr>

<tr> 
    <td align="center" width="15%"><b>姓&nbsp;&nbsp;&nbsp;&nbsp;名</b></td>
    <td width="30%">&nbsp;<%=rs1("姓名")%></td>
    <td align="center" width="15%"><b>曾&nbsp;用&nbsp;名</b></td>
    <td width="25%">&nbsp;<%if check("0")="no" and session("Urule")<>"a" then response.write "保密" else response.write rs("Uname")%></td>
    <td width="80" height="100" rowspan="5" align="center" valign=center><%if rs("havephoto")=false then%>无<BR>照<BR>片<%else%> <img src="showpic.asp?id=<%=rs("id")%>" width="80" height="100" border="0"><%end if%> 
    </td>
  </tr>
  <tr> 
    <td align="center"><b>性&nbsp;&nbsp;&nbsp;&nbsp;别</b></td>
    <td>&nbsp;<%if check("1")="no" and session("Urule")<>"a" then Response.Write "保密" else Response.Write rs("sex")%></td>
    </td>
    <td align="center"><b>民&nbsp;&nbsp;&nbsp;&nbsp;族</b></td>
    <td>&nbsp;<%if check("2")="no" and session("Urule")<>"a" then response.write "保密" else response.write rs("nation")%></td>
  </tr>
  <tr> 
    <td align="center"><b>所属部门</b></td>
    <td>&nbsp;<%=Session("Upart")%></td>
    <td align="center"><b>职&nbsp;&nbsp;&nbsp;&nbsp;务</b></td>
    <td>&nbsp;<%if check("3")="no" and session("Urule")<>"a" then response.write "保密" else response.write rs("duty")%></td>
  </tr>
  <tr> 
    <td align="center"><b>职&nbsp;&nbsp;&nbsp;&nbsp;称</b></td>
    <td>&nbsp;<%if check("4")="no" and session("Urule")<>"a" then response.write "保密" else response.write rs("grade")%></td>
    <td align="center"><b>出生日期</b></td>
    <td>&nbsp;<%if check("5")="no" and session("Urule")<>"a" then response.write "保密" else response.write rs("birthday")%></td>
  </tr>
  <tr> 
    <td align="center"><b>政治面貌</b></td>
    <td>&nbsp;<%if check("6")="no" and session("Urule")<>"a" then response.write "保密" else response.write rs("polity")%></td>
    <td align="center"><b>健康状况</b></td>
    <td>&nbsp;<%if check("7")="no" and session("Urule")<>"a" then response.write "保密" else response.write rs("health")%></td>
  </tr>
  <tr> 
    <td align="center" height="20"><b>籍&nbsp;&nbsp;&nbsp;&nbsp;贯</b></td>
    <td>&nbsp;<%if check("8")="no" and session("Urule")<>"a" then response.write "保密" else response.write rs("Nplace")%></td>
    <td align="center"><b>体&nbsp;&nbsp;&nbsp;&nbsp;重</b></td>
    <td colspan="2">&nbsp;<%if check("9")="no" and session("Urule")<>"a" then response.write "保密" else response.write rs("weight")%></td>
  </tr>
  <tr> 
    <td height="20" align="center"><b>身份证号</b></td>
    <td>&nbsp;<%if check("10")="no" and session("Urule")<>"a" then response.write "保密" else response.write rs("idcard")%></td>
    <td align="center"><b>身&nbsp;&nbsp;&nbsp;&nbsp;高</b></td>
    <td colspan="2">&nbsp;<%if check("11")="no" and session("Urule")<>"a" then response.write "保密" else response.write rs("height")%></td>
  </tr>
  <tr> 
    <td align="center" height="20"><b>婚姻状况</b></td>
    <td>&nbsp;<%if check("12")="no" and session("Urule")<>"a" then response.write "保密" else response.write rs("marriage")%></td>
    <td align="center"><b>毕业院校</b></td>
    <td colspan="2">&nbsp;<%if check("13")="no" and session("Urule")<>"a" then response.write "保密" else response.write rs("Fschool")%></td>
  </tr>
  <tr> 
    <td align="center" height="20"><b>本人成分</b></td>
    <td>&nbsp;<%if check("14")="no" and session("Urule")<>"a" then response.write "保密" else response.write rs("member")%></td>
    <td align="center"><b>专&nbsp;&nbsp;&nbsp;&nbsp;业</b></td>
    <td colspan="2">&nbsp;<%if check("15")="no" and session("Urule")<>"a" then response.write "保密" else response.write rs("speciality")%></td>
  </tr>
  <tr> 
    <td height="20" align="center"><b>工&nbsp;&nbsp;&nbsp;&nbsp;龄</b></td>
    <td>&nbsp;<%if check("16")="no" and session("Urule")<>"a" then response.write "保密" else response.write rs("length")%></td>
    <td align="center"><b>学&nbsp;&nbsp;&nbsp;&nbsp;历</b></td>
    <td colspan="2">&nbsp;<%if check("17")="no" and session("Urule")<>"a" then response.write "保密" else response.write rs("study")%></td>
  </tr>
  <tr> 
    <td height="20" align="center"><b>外语语种</b></td>
    <td>&nbsp;<%if check("18")="no" and session("Urule")<>"a" then response.write "保密" else response.write rs("foreign")%></td>
    <td align="center"><b>外语水平</b></td>
    <td colspan="2">&nbsp;<%if check("19")="no" and session("Urule")<>"a" then response.write "保密" else response.write rs("Elevel")%></td>
  </tr>
  <tr> 
    <td height="20" align="center"><b>记算机能力</b></td>
    <td>&nbsp;<%if check("20")="no" and session("Urule")<>"a" then response.write "保密" else response.write rs("Clevel")%></td>
    <td align="center"><b>户口所在地</b></td>
    <td colspan="2">&nbsp;<%if check("21")="no" and session("Urule")<>"a" then response.write "保密" else response.write rs("Hplace")%></td>
  </tr>
  <tr>
    <td height="20" align="center"><b>QQ号码</b></td>
    <td>&nbsp;<%if check("22")="no" and session("Urule")<>"a" then response.write "保密" else response.write rs("QQ")%></td>
    <td align="center"><b>EMAIL</b></td>
    <td colspan="2">&nbsp;<%=rs1("信箱")%></td>
  </tr>
  <tr>
    <td height="20" align="center"><b>常用电话</b></td>
    <td>&nbsp;<%=rs1("电话")%></td>
    <td align="center"><b>手机号码</b></td>
    <td colspan="2">&nbsp;<%=rs1("mobile")%></td>
  </tr>
  <tr>
    <td height="20" align="center"> <b>传呼号码</b> </td>
    <td>&nbsp;<%if check("23")="no" and session("Urule")<>"a" then response.write "保密" else response.write rs("call")%></td>
    <td align="center">&nbsp;</td>
    <td colspan="2">&nbsp;</td>
  </tr>
  <tr> 
    <td height="20" align="center"><b>现&nbsp;住&nbsp;址</b></td>
    <td colspan="4">&nbsp;<%if check("24")="no" and session("Urule")<>"a" then response.write "保密" else response.write rs("place")%></td>
  </tr>
  <tr> 
    <td height="20" align="center"><b>个人专长<br>
      以及爱好</b></td>
    <td colspan="4">&nbsp;<%if check("25")="no" and session("Urule")<>"a" then response.write "保密" else response.write rs("love")%></td>
  </tr>
  <tr> 
    <td height="20" align="center"><b>本人曾受<br>
      过何种奖<br>
      励和处分</b></td>
    <td colspan="4">&nbsp;<%if check("26")="no" and session("Urule")<>"a" then response.write "保密" else response.write rs("award")%></td>
  </tr>
  <tr> 
    <td height="20" align="center"><b>工作经历</b></td>
    <td colspan="4">&nbsp;<%if check("27")="no" and session("Urule")<>"a" then response.write "保密" else response.write rs("experience")%></td>
  </tr>
  <tr> 
    <td height="20" align="center"><b>家庭情况</b></td>
    <td colspan="4">&nbsp;<%if check("28")="no" and session("Urule")<>"a" then response.write "保密" else response.write rs("family")%></td>
  </tr>
  <tr> 
    <td height="20" align="center"><b>本&nbsp;&nbsp;&nbsp;&nbsp;人<br>
      联系方式</b></td>
    <td colspan="4">&nbsp;<%if check("29")="no" and session("Urule")<>"a" then response.write "保密" else response.write rs("contact")%></td>
  </tr>
  <tr> 
    <td height="20" align="center"><b>备&nbsp;&nbsp;&nbsp;&nbsp;注</b></td>
    <td colspan="4">&nbsp;<%if check("30")="no" and session("Urule")<>"a" then response.write "保密" else response.write rs("remark")%></td>
  </tr>
</table>

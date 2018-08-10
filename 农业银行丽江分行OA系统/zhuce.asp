<!--#include file="data.asp"-->
<!--#include file="html.asp"-->
<!--#INCLUDE FILE="mouse.js" -->
<%
Set myrs= Server.CreateObject("ADODB.Recordset") 
strSql="select * from bumen"
myrs.open strSql,Conn,1,1 
 name=htmlencode2(request("name"))
 password=request("password")
 userid=htmlencode2(request("userid"))
 question=htmlencode2(request("question"))
 answer=htmlencode2(request("answer"))
 email=request("email")
 mobile=request("mobile")
 tel=request("tel")
 ilevel=request("ilevel")
 department=htmlencode2(request("company"))
 ip= Request.ServerVariables("REMOTE_ADDR")
 nowtime=now()
sj=cstr(year(nowtime))+"-"+cstr(month(nowtime))+"-"+cstr(day(nowtime))+" "+cstr(hour(nowtime))+":"+right("0"+cstr(minute(nowtime)),2)+":"+right("0"+cstr(second(nowtime)),2)

set rs=server.createobject("ADODB.recordset")
rs.open "select * from user where 用户名='"& userid &"'order by id",conn,3,3
if rs.eof or rs.bof then
 else if userid=rs("用户名") then
  userid=""
  password=""
  %>
<link rel="stylesheet" href="oa.css">
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellspacing="1" cellpadding="2"> <tr > <td class="heading"> 
<div align="center"> <center> <table width="81%" border="0" cellspacing="0" cellpadding="0" bgcolor="#000000" bordercolorlight="#000000"> 
<tr> <td width="2%" align="right"></td><td align="left" height="25"> <p align="center"><font color="#FFFFFF"><b>您已经成功申请帐号</b></font></p></td><td width="3%"></td></tr> 
</table></center></div></td></tr> </table><div align="center"> <form method="post" action="saveedit1.asp?id=<%=id%>" name="myform" onsubmit="return  validate()"> 
<table width="80%" border="1" cellspacing="0" cellpadding="0" bordercolordark="#FFFFFF" bordercolor="#FFFFFF" bordercolorlight="#000000"> 
<tr> <td> <font COLOR="red">该帐号已经存在</font> </td></tr> </table></div><div align="center"><a  href="javascript:history.back(1)"><img border="0" src="images/previous.gif"></a>&nbsp;&nbsp; 
<a  href="javaScript:window.close()"><img   border="0" src="images/close_1.gif"></a> 
</div></form> <div align="center"> <center> <table border="1" width="80%" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#FFFFFF" bordercolor="#C0C0C0"> 
<tr> <td width="100%"> 系统提醒您: <ul> <li>你的姓名（<font color="#FF0000">必填</font>）</li><li>登录帐号（<font color="#FF0000">必填</font>）</li><li>登录密码（<font COLOR="#FF0000">必填</font>）</li><li>邮箱级别（<font COLOR="#FF0000">必填</font>）</li><li>公司名称（<font color="#FF0000">必填</font>）</li><li>电子邮件（提示：如果填写可以自动填写发电子邮件地址发电子邮件）</li></ul></td></tr> 
</table></center></div><%
  response.end
end if
end if
rs.close


set rs=server.createobject("ADODB.recordset") 
rs.Open "SELECT * FROM user Where ID is null",conn,1,3 
rs.addnew

rs("用户名")=userid
rs("密码")=password
rs("信箱")=email
rs("部门")=department
rs("问题")=question
rs("答案")=answer
rs("权限")=request("admin")
rs("审核")=false
rs("时间")=sj
rs("IP")=ip
rs("电话")=tel
rs("姓名")=name
rs("mobile")=mobile
rs("ilevel")=ilevel
if rs("ilevel")="" then rs("ilevel")="1"
if rs("权限")="" then rs("权限")="c"
rs.update 
id=rs("id")
%> <title>您已经成功申请帐号</title> <meta http-equiv="Content-Type" content="text/html; charset=gb2312"> 
<script Language="javaScript">
    function  validate()
    {
        if  (document.myform.name.value=="")
        {
            alert("姓名不能为空");
            document.myform.name.focus();
            return false ;
        }
        if  (document.myform.Userid.value=="")
        {
            alert("登录帐号不能为空");
            document.myform.Userid.focus();
            return false ;
        }
		if  (document.myform.company.value=="")
        {
            alert("单位名称不能为空");
            document.myform.company.focus();
            return false ;
        }
		if  (document.myform.tel.value=="")
        {
            alert("电话号码不能为空");
            document.myform.tel.focus();
            return false ;
        }
		if  (document.myform.email.value=="")
        {
            alert("电子邮件不能为空");
            document.myform.email.focus();
            return false ;
        }
        if  (document.myform.password.value=="")
        {
            alert("密码不能为空");
            document.myform.password.focus();
            return false ;
       }
        if  (document.myform.ilevel.value=="")
        {
            alert("邮箱级别不能为空");
            document.myform.password.focus();
            return false ;
        }
        return  true;
    }
</script>
<link rel="stylesheet" href="oa.css"> 
<table width="100%" border="0" cellspacing="1" cellpadding="2"> 
<tr > <td class="heading"> <div align="center"> <center> <table width="81%" border="0" cellspacing="0" cellpadding="0" bgcolor="#000000" bordercolorlight="#000000"> 
<tr> <td width="2%" align="right"></td><td align="left" height="25"> <p align="center"><font color="#FFFFFF"><b>您已经成功申请帐号</b></font></p></td><td width="3%"></td></tr> 
</table></center></div></td></tr> </table><div align="center">
 <form method="post" action="saveedit1.asp?id=<%=id%>" name="myform" onsubmit="return  validate()" > 
<table width="80%" border="1" cellspacing="0" cellpadding="0" bordercolordark="#FFFFFF" bordercolor="#FFFFFF" bordercolorlight="#000000"> 
<tr> <td width="17%" valign="top"> <p align="right">你的姓名:</p></td><td width="83%"> 
<input type="text" name="name" class="form" value="<%=rs("姓名")%>" size="24"> </td></tr> 
<tr> <td width="17%" valign="top" height="6"> <p align="right"><font size="2">登录帐号:</font></p></td><td width="83%" height="6"> 
	    <input type="hidden" name="Userid"  value="<%=rs("用户名")%>"  >
        <input type="text" name="Userid2" class="form" value="<%=rs("用户名")%>" size="24" disabled>


</td></tr> <tr> <td width="17%"  valign="top" height="16"> <p align="right"><font size="2">登录密码:</font></p></td><td width="83%" height="16"> 
<input type="password" name="password" class="form" size="24" value="<%=rs("密码")%>"> 
</td></tr> <tr> <td width="17%"  valign="top" height="16"> <p align="right"><font size="2">密码问题:</font></p></td><td width="83%" height="16"> 
<input type="text" name="question" class="form" size="24" value="<%=rs("问题")%>"> 
</td></tr> <tr> <td width="17%"  valign="top" height="16"> <p align="right"><font size="2">密码答案:</font></p></td><td width="83%" height="16"> 
<input type="text" name="answer" class="form" size="24" value="<%=rs("答案")%>"> 
</td></tr> <tr> <td width="17%"  valign="top"> <p align="right"><font size="2">单位名称:</font> 
</td><td width="83%"> <select NAME="company">

 <%if myrs.eof and myrs.bof then
response.write "<font color='red'>还没有任何东东</font>"
else

do while not (myrs.eof or myrs.bof)
if myrs("type")=rs("部门") then
sel="selected"
else 
sel=""
end if
%> <option value="<%=myrs("type")%>" <%=sel%>><%=myrs("type")%></option> <%myrs.movenext 
loop 
end if%> </select> </td></tr> <tr> <td width="17%"  valign="top"> <p align="right"><font size="2">手机号码:</font></p></td><td width="83%"> 
<input type="text" name="tel" class="form" value="<%=rs("mobile")%>" size="24"> 
</td></tr> <tr> <td width="17%"  valign="top"> <p align="right"><font size="2">电话号码:</font></p></td><td width="83%"> 
<input type="text" name="tel" class="form" value="<%=rs("电话")%>" size="24"> </td></tr> 
<tr> <td width="17%"  valign="top"> <p align="right"><font size="2">电子邮件:</font></p></td><td width="83%"> 
<input type="text" name="email" class="form" value="<%=rs("信箱")%>" size="24"> 
</td></tr>

<%if session("id")<>"" then %>
 <tr> 
<td WIDTH="17%"  VALIGN="top"> <p ALIGN="right">邮箱级别:</p></td><td WIDTH="83%"> 
<p><input TYPE="text" NAME="ilevel" CLASS="form" SIZE="1" value="<%=rs("ilevel")%>"></p></td></tr>
<tr> <td width="17%"  valign="top"> <p align="right">管理权限:</p></td><td width="83%"> 
<select NAME="admin"> <option value="a" <%if rs("权限")="a" then%>selected<%end if%>>超级用户</option> 
<option value="b" <%if rs("权限")="b" then%>selected<%end if%>>管理员</option> <option value="c" <%if rs("权限")="c" then%>selected<%end if%>>普通用户</option> 
</select> </td></tr> 
<%end if%> 

</table></div><div align="center"><input type=image  src="images/modify_off.gif">&nbsp;&nbsp; 
<a  href="javaScript:window.close()"><img   border="0" src="images/close_1.gif"></a> 
</div></form> <div align="center"> <center> <table border="1" width="80%" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#FFFFFF" bordercolor="#C0C0C0"> 
<tr> <td width="100%"> 系统提醒您: <ul> <li>你的姓名（<font color="#FF0000">必填</font>）</li><li>登录帐号（<font color="#FF0000">必填</font>）</li><li>登录密码（<font color="#FF0000">必填</font>）</li><li>邮箱级别（<font COLOR="#FF0000">必填</font>）</li><li>公司名称（<font color="#FF0000">必填</font>）</li><li>电子邮件（提示：如果填写可以自动填写发电子邮件地址发电子邮件）</li></ul></td></tr> 
</table></center></div>       


</body>
</html>
<%rs.close
set rs=nothing
%>
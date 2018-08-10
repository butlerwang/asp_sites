<!--#INCLUDE FILE="data.asp" -->
<!--#INCLUDE FILE="mouse.js" -->
<%
Set rs= Server.CreateObject("ADODB.Recordset") 
strSql="select * from bumen"
rs.open strSql,Conn,1,1 
%>
<html><head><title>丽江分行网络办公系统----申请帐号</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
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
		if  (document.myform.question.value=="")
        {
            alert("密码问题不能为空");
            document.myform.question.focus();
            return false ;
        }
		if  (document.myform.answer.value=="")
        {
            alert("密码答案不能为空");
            document.myform.answer.focus();
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
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form method="post" action="zhuce.asp" name="myform" onsubmit="return  validate()"> 
<table width="100%" border="0" cellspacing="1" cellpadding="2"> <tr > <td class="heading"> 
<div align="center"> <center> <table width="81%" border="0" cellspacing="0" cellpadding="0" bgcolor="#000000" bordercolorlight="#000000"> 
<tr> <td width="2%" align="right"></td><td align="left" height="25"> <p align="center"><font color="#FFFFFF"><b>申 
请 帐 号</b></font></p></td><td width="3%"></td></tr> </table></center></div></td></tr> 
</table><div align="center"> <table WIDTH="80%" BORDER="1" CELLSPACING="0" CELLPADDING="0" BORDERCOLORDARK="#FFFFFF" BORDERCOLOR="#FFFFFF" BORDERCOLORLIGHT="#000000"> 
<tr> <td WIDTH="17%" VALIGN="top"> <p ALIGN="right">你的姓名:</p></td><td WIDTH="83%"> 
<input TYPE="text" NAME="name" CLASS="form" SIZE="24">[请用你的真名] </td></tr> <tr> <td WIDTH="17%" VALIGN="top" HEIGHT="6"> 
<p ALIGN="right">登录姓名:</p></td><td WIDTH="83%" HEIGHT="6"> <input TYPE="text" NAME="Userid" CLASS="form" SIZE="24">[请用你的真名]<br></td></tr> 
<tr> <td WIDTH="17%"  VALIGN="top" HEIGHT="16"> <p ALIGN="right">登录密码:</p></td><td WIDTH="83%" HEIGHT="16"> 
<input TYPE="password" NAME="password" CLASS="form" SIZE="24"> [请牢记你的密码]</td></tr> <tr> 
<td WIDTH="17%"  VALIGN="top" HEIGHT="16"> <p ALIGN="right">密码问题:</p></td><td WIDTH="83%" HEIGHT="16"> 
<input TYPE="text" NAME="question" CLASS="form" SIZE="24" value=不用管>  [可以不用管]</td></tr> <tr> <td WIDTH="17%"  VALIGN="top" HEIGHT="16"> 
<p ALIGN="right">密码答案:</p></td><td WIDTH="83%" HEIGHT="16"> <input TYPE="text" NAME="answer" CLASS="form" SIZE="24" value=不用管> [可以不用管]
</td></tr> <tr> <td WIDTH="17%"  VALIGN="top"> <p ALIGN="right">单位名称: </td><td WIDTH="83%"> 
<select NAME="company"> <option selected> --==单位名称==--</option> <%if rs.eof and rs.bof then
response.write "<font color='red'>还没有任何东东</font>"
else

do while not (rs.eof or rs.bof)
%> <option VALUE="<%=rs("type")%>"><%=rs("type")%></option> <%rs.movenext 
loop 
end if%> </select> </td></tr> <tr> <td WIDTH="17%"  VALIGN="top"> <p ALIGN="right">手机号码:</p></td><td WIDTH="83%"> 
<input TYPE="text" NAME="mobile" CLASS="form" SIZE="24">为了联系方便请填写 </td></tr> <tr> <td WIDTH="17%"  VALIGN="top" HEIGHT="27"> 
<p ALIGN="right">电话号码:</p></td><td WIDTH="83%" HEIGHT="27"> <input TYPE="text" NAME="tel" CLASS="form" SIZE="24"> 
为了联系方便请填写</td></tr><tr> <td WIDTH="17%"  VALIGN="top"> <p ALIGN="right">电子邮件:</p></td><td WIDTH="83%"> 
<p><input TYPE="text" NAME="email" CLASS="form" SIZE="24" value="XX@LJ.YN.ABC"></p></td></tr> <%if session("id")<>"" then %> 
<tr> <td WIDTH="17%"  VALIGN="top"> <p ALIGN="right">邮箱级别:</p></td><td WIDTH="83%"> 
<p><input TYPE="text" NAME="ilevel" CLASS="form" SIZE="1"></p></td></tr> <tr> 
<td width="17%"  valign="top"> <p align="right">管理权限:</p></td><td width="83%"> 
<select NAME="admin"> <option value="a" <%if rs("权限")="a" then%>selected<%end if%>>超级用户</option> 
<option value="b" <%if rs("权限")="b" then%>selected<%end if%>>管理员</option> <option value="c" <%if rs("权限")="c" then%>selected<%end if%>>普通用户</option> 
</select> </td></tr> <%end if%> </table></div><div align="center"><input type=image  src="images/add_off.gif">&nbsp;&nbsp; 
<a  href="javaScript:window.close()"><img   border="0" src="images/close_1.gif"></a> 
</div></form><div align="center"> <center> <table border="1" width="80%" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#FFFFFF" bordercolor="#C0C0C0" HEIGHT="102"> 
<tr> <td width="100%" HEIGHT="101"> 系统提醒您:<FONT COLOR="#FF0000">请用您的真实姓名进行注册，否则不予以审核（凡未经审核通过的，将不能进入系统及使用任何功能）</FONT> <ul><li>你的姓名（<font color="#FF0000">必填</font>）；&nbsp;</li><li>登录姓名（<font color="#FF0000">必填</font>）；&nbsp;</li><li>登录密码：（<font color="#FF0000">必填</font>）</li><li>单位名称（<font color="#FF0000">必填</font>）</li><li>手机号码</li><li>电子邮件</li></ul></td></tr> 
</table></center></div>
       
</body>       
</html>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<!-- #include file="inc/functions.asp" -->
<%
errno=Request("errno")
If errno<>"" Then
	If CInt(errno)=2 Then
		errmsg="用户名或密码不能为空！"
	End If

	If CInt(errno)=1 Then
		errmsg="用户名或密码错误！"
	End If
	
	If CInt(errno)=0 Then
		errmsg="验证码错误！"
	End If

End If
%>

<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=gdb("select web_name from web_settings ")%>-企业网站管理系统</title>
<link href="inc/logincss.css" rel="stylesheet" type="text/css" />
<script language="JavaScript">
function CheckLogin()
{
	if(document.MyForm.username.value=="")
	{
		alert("请输入用户名！");
		document.MyForm.username.focus();
		return false 
	}
	
	if(document.MyForm.password.value=="")
	{
		alert("请输入密码！");
		document.MyForm.password.focus()
		return false 
	}
	
	if(document.MyForm.verifycode.value=="")
	{
		alert("请输入验证码！");
		document.MyForm.verifycode.focus()
		return false 
	}
}
</script>
</head>
<body>
<div class="mains">
<div class="inners">
<div class="lefts"> </div>

<div class="login">
<form action="admin_login.asp" method="post" name="MyForm" id="MyForm"><INPUT type="hidden" value="chklogin" name="reaction">
<div class="center">
<div class="inner">
 <table   cellpadding="0" cellspacing="0" id="innnertalbe">
        <tr>
          <td  height="50">用户名</td>
          <td ><input name="username" type="text" class="login_textfield" id="username" size="16" maxlength="100" /></td>
         </tr>
		
		<tr>
		  <td height="50" >密　码</td>
          <td ><input name="password" type="password" class="login_textfield" id="password" size="16" maxlength="100" /></td>
        </tr>
		
		<tr>
		  <td height="50" >验证码</td>
          <td><input name="verifycode"  maxLength=5 size=10 >&nbsp;<img src="/inc/getcode.asp" width="55"  onclick="this.src=this.src+'?'" alt="图片看不清？点击重新得到验证码" style="cursor:hand;"></td>
        </tr>
		
		<tr>
		  <td height="50"></td>
          <td><input name="image" type="submit" class="LoginSub" onclick="return CheckLogin()" value=" 登 录 " />　　<span style="color:#ffffff;"><%=errmsg%></span></td>
        </tr>
      </table>
</div>
<div class="clearfix"></div>
</div>
</form>
</div>


</div>
</div>
<div class="CopyR">2012-2112 &copy; <a href="http://www.juhaoyong.net" target="_blank">JuhaoyongCMS</a> www.juhaoyong.net 版权所有 All rights reserved  </div>

</body>
</html>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<!-- #include file="inc/functions.asp" -->
<%
errno=Request("errno")
If errno<>"" Then
	If CInt(errno)=2 Then
		errmsg="�û��������벻��Ϊ��!"
	End If

	If CInt(errno)=1 Then
		errmsg="������û���������!"
	End If
	If CInt(errno)=0 Then
		errmsg="����ĸ�����!"
	End If
End If
%>

<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=gdb("select web_name from web_settings ")%>-��ҵ��վ����ϵͳ</title>
<link href="inc/logincss.css" rel="stylesheet" type="text/css" />
<script language="JavaScript">
function CheckLogin()
{
	if(document.MyForm.username.value=="")
	{
		alert("�������û������ύ��");
		document.MyForm.username.focus();
		return false 
	}
	if(document.MyForm.password.value=="")
	{
		alert("�������������ύ��");
		document.MyForm.password.focus()
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
          <td  height="60">�� ��</td>
          <td ><input name="username" type="text" class="login_textfield" id="username" size="16" maxlength="100" /></td>
         </tr><tr>
		  <td height="60" >�� ��</td>
          <td ><input name="password" type="password" class="login_textfield" id="password" size="16" maxlength="100" /></td>
        </tr><tr >
		  <td height="49"></td>
          <td><input name="image" type="submit" class="LoginSub" onclick="return CheckLogin()" value=" �� ¼ " />
          <br><span style="color:#FF0000;"><%=errmsg%></span></td>
        </tr>
      </table>
</div>
<div class="clearfix"></div>
</div>
</form>
</div>


</div>
</div>
<div class="CopyR">2012 &copy; <a href="http://www.Junww.com" target="_blank">CNQI</a> www.west263.org��Ȩ���� All rights reserved  </div>

</body>
</html>

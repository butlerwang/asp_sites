<!--#include file="check.asp"-->
<!--#include file="../inc/Check_Sql.asp"-->

<%
If Session("name") = "" then
response.Redirect("login.asp")
'response.Write "<script LANGUAGE=javascript>alert('���糬ʱ�������㻹û�е�½! ');window.location.href='login.asp';</'script>"
response.End
end if%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��̨����ϵͳ</title>
<link href="include/style.css" rel="stylesheet" type="text/css">
</head>

<frameset rows="93,*" framespacing="0" frameborder="1" border="false" scrolling="yes">
  <frame name="top" scrolling="no" src="top.asp">
  <frame name="main" scrolling="auto" src="ad_main.asp">
</frameset>
<noframes>
  <body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
  <p>���������汾���ͣ�IE5�����ϰ汾����ʹ�ã�</p>
  </body>
</noframes>
</html>

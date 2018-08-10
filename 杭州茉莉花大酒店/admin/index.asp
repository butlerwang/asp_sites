<!--#include file="check.asp"-->
<!--#include file="../inc/Check_Sql.asp"-->

<%
If Session("name") = "" then
response.Redirect("login.asp")
'response.Write "<script LANGUAGE=javascript>alert('网络超时，或者你还没有登陆! ');window.location.href='login.asp';</'script>"
response.End
end if%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>后台管理系统</title>
<link href="include/style.css" rel="stylesheet" type="text/css">
</head>

<frameset rows="93,*" framespacing="0" frameborder="1" border="false" scrolling="yes">
  <frame name="top" scrolling="no" src="top.asp">
  <frame name="main" scrolling="auto" src="ad_main.asp">
</frameset>
<noframes>
  <body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
  <p>你的浏览器版本过低！IE5及以上版本才能使用！</p>
  </body>
</noframes>
</html>

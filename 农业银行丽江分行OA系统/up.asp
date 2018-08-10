<html>
<head>
<title>上传图标</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
table {  font-size: 9pt}
select {  font-size: 9pt}
input {  font-size: 9pt; background-color: #CCCCFF; font-weight: bold; color: #FF6633; border-style: groove}
.smallbox {  font-size: 1pt}
a:link {  font-size: 9pt; text-decoration: none; }
a:hover {  font-size: 9pt;}
body {  font-size: 9pt}
-->
</style>
</head>

<body bgcolor="#FFFFFF">
<%
 if session("admin")="" then
if request("file")<>"" and request("id")<>"" then
if request("file")="admin" and request("id")="admin" then
  session("admin")="admin"
  response.redirect "up.asp"
  else
response.write "用户名或密码错误"
response.end
end if
end if
%>
<form method="post" action="">
        用户名： 
        <input type="text" name="file" size="32">
        <br>
        密&nbsp;&nbsp;码：
        <input type="password" name="id" size="32">
        <br>
        <input type="submit" name="Submit" value=" 提 交 ">
        <input type="reset" name="reset" value=" 重 写 ">
        <br>
</form>
<%
 else
%>

<table border="0" align="center" cellpadding="0">
  <tr valign="middle"> 
    <td>
      <form method="post" action="savelogo.asp" name="reg" enctype="multipart/form-data">
        路径： 
        <input type="file" name="file" size="32">
        <br>
        说明：
        <input type="text" name="id" value="test">
        （如果已经存在，则覆盖）<br>
        <input type="submit" name="Submit" value="开始上传">
        <input type="reset" name="reset" value="重新选择">
        <br>
      </form>
    </td>
  </tr>
</table>
<script language=vbscript>
function reg_Onsubmit()
if document.reg.id.value="" then
	msgbox "说明必须填写！"
	reg_onsubmit=false
	exit function
else
	document.reg.action="savelogo.asp?id="&document.reg.id.value
end if
end function
</script>
</body>
</html>
<%
end if
%>
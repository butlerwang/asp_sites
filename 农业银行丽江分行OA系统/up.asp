<html>
<head>
<title>�ϴ�ͼ��</title>
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
response.write "�û������������"
response.end
end if
end if
%>
<form method="post" action="">
        �û����� 
        <input type="text" name="file" size="32">
        <br>
        ��&nbsp;&nbsp;�룺
        <input type="password" name="id" size="32">
        <br>
        <input type="submit" name="Submit" value=" �� �� ">
        <input type="reset" name="reset" value=" �� д ">
        <br>
</form>
<%
 else
%>

<table border="0" align="center" cellpadding="0">
  <tr valign="middle"> 
    <td>
      <form method="post" action="savelogo.asp" name="reg" enctype="multipart/form-data">
        ·���� 
        <input type="file" name="file" size="32">
        <br>
        ˵����
        <input type="text" name="id" value="test">
        ������Ѿ����ڣ��򸲸ǣ�<br>
        <input type="submit" name="Submit" value="��ʼ�ϴ�">
        <input type="reset" name="reset" value="����ѡ��">
        <br>
      </form>
    </td>
  </tr>
</table>
<script language=vbscript>
function reg_Onsubmit()
if document.reg.id.value="" then
	msgbox "˵��������д��"
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
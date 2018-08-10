<!--#include file="../check.asp"-->
<!--#include file="../include/conn.asp"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>后台管理系统</title>
<link href="../include/style.css" rel="stylesheet" type="text/css">
</head>
<body>
<br>
<%Call OpenData()
IF instr(webConfig,", 7")>=1 Then'网站功能配置
	    IF instr(session("manconfig"),", 7")>=1 Then'网站管理权限设置
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="6%" align="center">&nbsp;</td>
    <td width="94%" height="20"><img src="../images/ducument.gif" width="25" height="13"> <strong>在线留言</strong></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="8%">&nbsp;</td>
        <td width="92%" height="20"><img src="../images/next(1).gif" width="9" height="10"> &nbsp;<a href="list.asp" target="right">在线留言</a></td>
      </tr>
    </table></td>
  </tr>
</table>
<%
End IF
End IF
IF instr(webConfig,", 9")>=1 Then'网站功能配置
	    IF instr(session("manconfig"),", 9")>=1 Then'网站管理权限设置
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="6%" align="center">&nbsp;</td>
    <td width="94%" height="20"><img src="../images/ducument.gif" width="25" height="13"> <strong>友情链接</strong></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="8%">&nbsp;</td>
        <td width="92%" height="20"><img src="../images/next(1).gif" width="9" height="10"> &nbsp;<a href="../weblink/" target="right">友情链接管理</a></td>
      </tr>
    </table></td>
  </tr>
</table>
<%
End IF
End IF
Call CloseDataBase()  
%>
</body>
</html>
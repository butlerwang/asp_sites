<!--#include file="../check.asp"-->
<!--#include file="../include/conn.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>后台管理系统</title>
<link href="../include/style.css" rel="stylesheet" type="text/css">
<base target="right">
</head>
<body>
<br>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td align="center">&nbsp;</td>
    <td height="20"><img src="../images/ducument.gif" width="25" height="13"> <span class="font_bold">企业信息</span></td>
  </tr>
  <tr>
    <td height="45">&nbsp;</td>
    <td>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="8%">&nbsp;</td>
          <td width="92%" height="20"><img src="../images/next(1).gif" width="9" height="10"> 
            &nbsp;<a href="company.asp">增加信息</a></td>
        </tr>
      </table>	  
    <%OpenData()
	  set rs=server.CreateObject("adodb.recordset")
	  sql="select * From Sbe_WebConfig "
	  Rs.open sql,conn,1,1
	    if rs("ShowqiyeClass")=true then	  
	  %>
	  <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="8%">&nbsp;</td>
          <td width="92%" height="20"><img src="../images/next(1).gif" width="9" height="10"> 
            &nbsp;<a href="../class/index.asp?classtitle=Sbe_Company" target="right">信息分类</a></td>
        </tr>
      </table>
	  <%end if
	  rs.close
	  set rs=nothing
	  Call CloseDataBase()%>
	  </td>
  </tr>
  <tr> 
    <td width="6%" height="45">&nbsp;</td>
    <td width="94%">&nbsp;</td>
  </tr>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
</body>
</html>
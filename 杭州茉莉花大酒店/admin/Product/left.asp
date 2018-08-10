<!--#include file="../check.asp"-->
<!--#include file="../include/conn.asp"-->
<html>                                                                               
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>后台管理系统</title>
<link href="../include/style.css" rel="stylesheet" type="text/css">
<style  type="text/css">  

</style>  

</head>
<body>
<br>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="6%" align="center">&nbsp;</td>
    <td width="94%" height="20"><img src="../images/ducument.gif" width="25" height="13"> 
      <strong>客房发布中心</strong></td>
  </tr>
  <tr> 
    <td height="45">&nbsp;</td>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="8%">&nbsp;</td>
          <td width="92%" height="20"><img src="../images/next(1).gif" width="9" height="10"> 
            &nbsp;<a href="list.asp" target="right">客房列表</a></td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="8%">&nbsp;</td>
          <td width="92%" height="20"><img src="../images/next(1).gif" width="9" height="10"> 
            &nbsp;<a href="add.asp" target="right">增加客房</a></td>
        </tr>
      </table>
      <%
	  OpenData()
	  set rs=server.CreateObject("adodb.recordset")
	  sql="select * From Sbe_WebConfig "
	  Rs.open sql,conn,1,1
	    if rs("showproclass")=true then	  
	  %>
	  <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="8%">&nbsp;</td>
          <td width="92%" height="20"><img src="../images/next(1).gif" width="9" height="10"> 
            &nbsp;<a href="../class/index.asp?classtitle=sbe_product" target="right">客房标志分类管理</a></td>
        </tr>
      </table>
	  <%end if
	  if rs("Pro_order")=true then%>
     <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="8%">&nbsp;</td>
          <td width="92%" height="20"><img src="../images/next(1).gif" width="9" height="10"> 
            &nbsp;<a href="dingdan.asp" target="right">产品订单管理</a></td>
        </tr>
      </table>
	  <%end if
	  rs.close
	  set rs=nothing
	  Call CloseDataBase()
	  %>
<!--<table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="8%">&nbsp;</td>
          <td width="92%" height="20"><img src="../images/next(1).gif" width="9" height="10"> 
            &nbsp;<a href="news.asp?classtitle=sbe_news" target="right">公告管理</a></td>
        </tr>
      </table>-->
	  </td>
  </tr>
</table>
<p>&nbsp;</p>
</body>
</html>
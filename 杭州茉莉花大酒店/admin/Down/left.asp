<!--#include file="../check.asp"-->
<!--#include file="../include/conn.asp"-->
<html>                                                                               
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��̨����ϵͳ</title>
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
      <strong>���������������</strong></td>
  </tr>
  <tr> 
    <td height="45">&nbsp;</td>
    <td> 
	  <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="8%">&nbsp;</td>
          <td width="92%" height="20"><img src="../images/next(1).gif" width="9" height="10"> 
            &nbsp;<a href="list.asp" target="right">��Ϣ�б�</a></td>
        </tr>
      </table>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="8%">&nbsp;</td>
          <td width="92%" height="20"><img src="../images/next(1).gif" width="9" height="10"> 
            &nbsp;<a href="add.asp" target="right">������Ϣ</a></td>
        </tr>
      </table>
      <%
	  OpenData()
	  set rs=server.CreateObject("adodb.recordset")
	  sql="select ShowProClass From Sbe_WebConfig "
	  Rs.open sql,conn,1,1
	    if rs("showproclass")=true then
		if session("flag")=99 then
	  %>
	 <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="8%">&nbsp;</td>
          <td width="92%" height="20"><img src="../images/next(1).gif" width="9" height="10"> 
            &nbsp;<a href="../class/index.asp?classtitle=Sbe_Down" target="right">��Ϣ�������</a></td>
        </tr>
      </table>
<!--<table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="8%">&nbsp;</td>
          <td width="92%" height="20"><img src="../images/next(1).gif" width="9" height="10"> 
            &nbsp;<a href="news.asp?classtitle=sbe_news" target="right">�������</a></td>
        </tr>
      </table>-->
	  <%
	  end if
	  end if
	  rs.close
	  set rs=nothing
	  Call CloseDataBase()
	  %>
	  
	  </td>
  </tr>
</table>
<p>&nbsp;</p>
</body>
</html>
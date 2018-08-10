<!--#include file="check.asp"-->
<!--#include file="conn.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>后台管理系统</title>
<link href="include/style.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<br><br><br><br>
<%Call OpenData()%>
<table width="600" border="0" align="center" cellpadding="0" cellspacing="0" id="sbe_table">
  <tr>
    <td height="30" align="center" class="sbe_table_title" id="title"><b>
	<%
	Set oRs=Conn.Execute("select WebName,Company from Sbe_WebConfig",1,1)
	Response.Write ("技术支持：<a href='"&oRs(0)&"' target='_blank'>"&oRs(1)&"</a>")
	oRs.Close:set oRs=Nothing 
	
	%>	
	</b></td>
  </tr>
  <tr>
    <td height="22" align="left">　管 理员 :&nbsp;<%=ucase(session("name"))%>&nbsp;
    <%call Geet()%></td>
  </tr>
  <tr>
    <td height="22" align="left">　登陆时间:&nbsp;<%=DATE()%></td>
  </tr>  
  <tr>
    <td height="22" align="left">　登 陆IP :&nbsp;<%=request.serverVariables("remote_host")%></td>
  </tr>
  <tr align="center">
    <td height="30" class="sbe_table_title" id="title"><strong>管理员权限说明</strong></td>
  </tr>
  <tr>
    <td>
　<%Call check_name_str(session("manconfig"))%>
   </td>
  </tr>  
</table>
<%Call CloseDataBase()%>
</body>
</html>
<%
sub Geet()
TD=hour(now)
if TD<12 then 
str="早上好!"
elseif TD<18 then
str="下午好!"
else
str="晚上好!"
end if 
response.write(str)
end sub
Private Sub check_name_str(strID)
   arry=split(strID,",")
   for i=0 to ubound(arry)
     Call check_name(arry(i))
   next
End Sub   

%>
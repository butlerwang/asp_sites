<!--#include file="check.asp"-->
<!--#include file="conn.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>后台管理系统</title>
<link href="include/style.css" rel="stylesheet" type="text/css">
</head>
<base target="main">
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
Call OpenData()
%>
<table width="100%" height="93" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="100%" height="93" background="images/top.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="34%" height="41">&nbsp;&nbsp;<img src="images/logo.gif" width="180" height="49"></td>
          <td width="40%">&nbsp;</td>
          <td width="26%">&nbsp;</td>
        </tr>
        <tr valign="middle"> 
          <td height="37" colspan="3"><table border="1" cellpadding="0" id="button_img">
            <tr>
			    <td><a href="index.asp" target="_top">首页</a></td>
			  <%
				IF instr(webConfig," 5")>=1 Then'网站功能配置
				  IF instr(session("manconfig")," 5")>=1 Then'网站管理权限设置	
				%>
                <td><a href="member/">权限管理</a></td>
				<%
				  End IF
				End IF
				%>
				<%
				IF instr(webConfig," 1")>=1 Then'网站功能配置
				  IF instr(session("manconfig")," 1")>=1 Then'网站管理权限设置	
				%>
                <td><a href="Company/">企业信息</a></td>
				<%
				  End If
				End IF
				%>
				<%
				IF instr(webConfig," 2")>=1 Then'网站功能配置
				  IF instr(session("manconfig")," 2")>=1 Then'网站管理权限设置	
				%>
                <td><a href="product/">客房中心</a></td>
				<%
				  End IF
				End IF
				%>
				<%
				IF instr(webConfig," 3")>=1 Then'网站功能配置
				  IF instr(session("manconfig")," 3")>=1 Then'网站管理权限设置	
				%>
                <td><a href="news/">资讯中心</a></td>
				<%
				  End If
				End IF
				%>
				<%
				IF instr(webConfig," 6")>=1 Then'网站功能配置
				  IF instr(session("manconfig")," 6")>=1 Then'网站管理权限设置	
				%>
                <td><a href="job/">人事招聘</a></td>
				<%
				  End If
				End IF
				%>
				<%
				IF instr(webConfig," 7")>=1 Then'网站功能配置
				  IF instr(session("manconfig")," 7")>=1 Then'网站管理权限设置	
				%>
                <td><a href="ServeCenter/">在线留言</a></td>
				<%
				  End If
				End IF
				%>
				<%
				IF instr(webConfig," 4")>=1 Then'网站功能配置
				  IF instr(session("manconfig")," 4")>=1 Then'网站管理权限设置	
				%>
                <td><a href="Down/">店铺形象</a></td>
				<%
				  End If
				End IF
				%>
                <%
				IF instr(webConfig," 8")>=1 Then'网站功能配置
				  IF instr(session("manconfig")," 8")>=1 Then'网站管理权限设置	
				%>
                <td><a href="orders/">在线预定</a></td>             
				<%
				  End If
				End IF
				%>
                <%
				IF instr(webConfig," 9")>=1 Then'网站功能配置
				  IF instr(session("manconfig")," 9")>=1 Then'网站管理权限设置	
				%>
                <td><a href="weblink/">楼盘标志</a></td>             
				<%
				  End If
				End IF
				%>
				<%if Session("flag")=99 then%>
				<td><a href="manage/">网站设置</a></td>
				<%end if%>
			<td><%if session("name") ="" then%><a href="quit.asp"><span style="font-weight:bold;color:#F00; ">登录</span></a><%else%><a href="quit.asp"><span style="font-weight:bold;color:#F00; ">退出</span></a><%end if%></td>
			</tr>
          </table></td>
        </tr>
      </table></td>
  </tr>
</table>
<%Call CloseDataBase()%>
</body>
</html>
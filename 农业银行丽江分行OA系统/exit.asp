<!--#INCLUDE FILE="data.asp" -->

<%
Set my_rs= Server.CreateObject("ADODB.Recordset") 
	strSql="select * from user where 用户名='"&Session("Uname")&"'"
	my_rs.open strSql,Conn,3,3 
	
		my_rs("状态")=false
		my_rs.update
        my_rs.close
		set my_rs=nothing
		conn.close
		set conn=nothing
		Session("Uid")=""
		Session("Uname")=""
		Session("Upass")=""
		Session("Upart")=""
		Session("Urule")=""
		Session("Ulogin")="no"
		
'-----------------------------邮箱系统清除环境变量----------------------------------------
For each k in Session.Contents
Session.Contents(k)=""
next

'-----------------------------邮箱系统清除环境变量----------------------------------------
%>
<script>
top.location.href="close.htm"
</script>
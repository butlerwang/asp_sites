<!--#INCLUDE FILE="data.asp" -->

<%
Set my_rs= Server.CreateObject("ADODB.Recordset") 
	strSql="select * from user where �û���='"&Session("Uname")&"'"
	my_rs.open strSql,Conn,3,3 
	
		my_rs("״̬")=false
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
		
'-----------------------------����ϵͳ�����������----------------------------------------
For each k in Session.Contents
Session.Contents(k)=""
next

'-----------------------------����ϵͳ�����������----------------------------------------
%>
<script>
top.location.href="close.htm"
</script>
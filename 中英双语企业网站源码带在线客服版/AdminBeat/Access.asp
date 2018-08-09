

<%
'chk session
If Session("log_name")="" Then 
	Session.abandon
	Response.redirect "login.asp"
End If 
%>
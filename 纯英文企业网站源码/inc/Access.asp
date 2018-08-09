

<%
'chk session
If Session("log_name")="" Then 
response.redirect "login.asp"	
%>
<%
End If 
%>
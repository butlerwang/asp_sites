<!--#include file="data.asp"--> 
<%
if session("Urule")<>"a" then
response.redirect "error.asp?id=admin"
end if
%>
<%

	set rs=server.createobject("adodb.recordset")
	sql="DELETE * from soft where id="&request("id")
	rs.open sql,conn,1,3
	rs.update
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing

	Response.Redirect "file.asp" 

%>

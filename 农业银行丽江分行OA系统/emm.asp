<!--#include file="data.asp"-->
<%
if session("Urule")<>"a" then
response.redirect "error.asp?id=admin"
end if
	set rs=server.createobject("adodb.recordset")
	if request("del")<>"" then
		call delsoft()
	elseif request("edit")<>"" then
		call editsoft()
	else
		call newsoft()
	end if
sub newsoft()
	sql="select * from bumen where (id is null)" 
	rs.open sql,conn,1,3
	rs.addnew
	rs("type")=request("type")
	rs.update
end sub
sub editsoft()
	sql="select * from bumen where id="&request("id")
	rs.open sql,conn,1,3
	rs("type")=request("type")
	rs.update
end sub
sub delsoft()
	sql="DELETE * from bumen where id="&request("id")
	rs.open sql,conn,1,3
	rs.update
end sub

	rs.close
	set rs=nothing
	conn.close
	set conn=nothing

	Response.Redirect("mm.asp")

%>

<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file="data.asp"-->
<!--#include file="check.asp"-->
<%
if Session("Urule")<>"a" then
response.write "您没有权限：P"
response.end
end if
   dim sql 
   dim rs
   
   set rs=server.createobject("adodb.recordset")
   sql="delete from bbs where id="&request("id")
   rs.open sql,conn,3,3
   response.redirect "bbs.asp"
   conn.close
   set conn=nothing
   rs.close
   set rs=nothing  
%>

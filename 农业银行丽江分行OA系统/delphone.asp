<!--#INCLUDE FILE="data.asp" -->
<!--#INCLUDE FILE="check.asp" -->
<%

if session("Urule")<>"a" then
response.redirect("tel.asp")
end if
set rs=server.createobject("ADODB.recordset") 
rs.Open "DELETE * FROM tel Where ID="&request("id"),conn,1,3 
rs.update 
rs.close
set rs=nothing
conn.Close
set conn = nothing
Response.Redirect ("tel.asp")

%>


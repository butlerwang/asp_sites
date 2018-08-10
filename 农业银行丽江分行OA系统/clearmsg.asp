<!--#INCLUDE FILE="data.asp" -->
<%
set rs=server.createobject("ADODB.recordset") 
rs.open "select * from chat where send='"&session("Uid")&"' and receive='"&request("id")&"'order by id",conn,1,3
if not rs.eof then
do while not (rs.eof or rs.bof)
rs.Delete        
rs.movenext 
loop 
end if
rs.close
set rs=nothing
conn.Close
set conn = nothing
Response.Redirect ("show.asp?receiveuser="&session("receiveuser")&"&id="&session("receive"))

%>


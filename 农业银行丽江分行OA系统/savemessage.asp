<!--#INCLUDE FILE="data.asp" -->

<%
message=request("message")
receive=request("id")
send=session("Uid")
user=session("receiveuser")
nowtime=now()
sj=cstr(year(nowtime))+"-"+cstr(month(nowtime))+"-"+cstr(day(nowtime))+" "+cstr(hour(nowtime))+":"+right("0"+cstr(minute(nowtime)),2)

set rs=server.createobject("ADODB.recordset") 
rs.Open "SELECT * FROM chat Where id is null",conn,1,3 
rs.addnew

rs("message")="("&sj&") "&session("Rname")&" ¶Ô "&session("receiveuser")&chr(13)&message&chr(13)
rs("receive")=receive
rs("send")=send
rs("from")=Session("Rname")
rs("time")=sj
rs.update 

if user<>session("Rname") then
set rs=server.createobject("ADODB.recordset") 
rs.Open "SELECT * FROM chat Where id is null",conn,1,3 
rs.addnew

		rs("message")="("&sj&") "&session("Rname")&" ¶Ô "&session("receiveuser")&chr(13)&message&chr(13)&rs("message")
rs("receive")=send
rs("send")=receive
rs("from")=Session("Rname")
rs("time")=sj
rs("zt")=true
rs.update 
rs.close
set rs=nothing

end if%>
<script>
window.close()
</script>

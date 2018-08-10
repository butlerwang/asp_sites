<!--#include file="data.asp"-->
<!--#include file="html.asp"-->
<!--#include file="char.asp"-->
<!--#include file="check.asp"-->

<%
 title=request("title")
 picture=request("pictureid")
 content=request("content")
 name=Session("Uname")
 ip= Request.ServerVariables("REMOTE_ADDR")
 nowtime=now()
sj=cstr(year(nowtime))+"-"+cstr(month(nowtime))+"-"+cstr(day(nowtime))+" "+cstr(hour(nowtime))+":"+right("0"+cstr(minute(nowtime)),2)+":"+right("0"+cstr(second(nowtime)),2)

set rs=server.createobject("ADODB.recordset") 
rs.Open "SELECT * FROM bbs Where ID is null",conn,1,3 
rs.addnew

rs("name")=name
rs("subject")=htmlencode2(title)
rs("Time")=sj
rs("Content")=ubbcode(Content)
rs("Pic")=picture
rs("Knock")="0"
if request("subjectid")<>"" then
rs("SubjectId")=request("subjectid")
else
rs("subjectid")=0
end if
rs("IP")=ip
rs.update 
rs.close
set rs=nothing
%>
<script language=javascript>
opener.location=opener.location;window.close();
</script>

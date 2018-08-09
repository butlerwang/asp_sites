<%@ LANGUAGE=VBScript CodePage=65001%>
<% response.charset="utf-8" %>
<% session.codepage=65001 %>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" /><!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<style type="text/css">
*{
	width:500px;
	margin:0px;
	padding:0px;}
.PicsList{
	padding:5px;
	margin:0px;}
.PicsList ul{
	margin:0px;
	padding:0px;}
.PicsList ul li{
	float:left;
	font-size:12px;
	width:60px;
	padding:5px 10px;
	margin:0}
.PicsList ul li p{
	text-align:center;
	padding:0;
	margin:0;}
.PicsList ul li img{
	border:none;}
</style>
<div class='PicsList'>
<ul>
<%
a_id=cint(request.querystring("id"))
FileName=request.querystring("Name")
FileOrder=cint(request.querystring("order"))

set rs=server.createobject("adodb.recordset")
sql="select Pics,[id] from [Article] where id="&a_id&" and Pics like '%"&FileName&"%'"
rs.open(sql),cn,1,3
if not rs.eof then
PicsContent=split(rs("Pics"),",")
PicsCount=ubound(PicsContent)

if PicsCount>0 then

if FileOrder=PicsCount then
rs("Pics")=replace(rs("Pics"),","&FileName,"")
else
rs("Pics")=replace(rs("Pics"),FileName&",","")
end if

else
rs("Pics")=replace(rs("Pics"),FileName,"")
end if
rs.update
'先判断文件是否存在，否则删除
Set fso=Server.CreateObject("Scripting.FileSystemObject")
If fso.FileExists(Server.MapPath("/images/up_images/"&FileName)) then
FilePath="/images/up_images/"&FileName
call DelFile(FilePath)
end if


response.Write "<script language='javascript'>alert('删除成功！');location.href='Pics_list.asp?id="&a_id&"';</script>"

end if
rs.close
set rs=nothing
%>



</ul>   
</div>

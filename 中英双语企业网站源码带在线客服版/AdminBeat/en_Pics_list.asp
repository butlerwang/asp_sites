<%@ LANGUAGE=VBScript CodePage=65001%>
<% response.charset="utf-8" %>
<% session.codepage=65001 %>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" /><!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<style type="text/css">
*{
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
set rs=server.createobject("adodb.recordset")
sql="select Pics,[id] from [en_Article] where id="&a_id&" and Pics<>'' "
rs.open(sql),cn,1,1
if not rs.eof then
PicsContent=split(rs("Pics"),",")
PicsCount=ubound(PicsContent)
for i=0 to PicsCount
%>
<li><a href="/images/up_images/<%=PicsContent(i)%>" target="_blank"><img src="/images/up_images/<%=PicsContent(i)%>" width="60" height="60"></a><p><a href='en_Pics_Del.asp?order=<%=i%>&id=<%=rs("id")%>&Name=<%=PicsContent(i)%>'>删除</a></p></li>

<%
next
else
response.write "<li style='color:#FF0000;'>0 图片</li>"
end if
rs.close
set rs=nothing
%>



</ul>   
</div>

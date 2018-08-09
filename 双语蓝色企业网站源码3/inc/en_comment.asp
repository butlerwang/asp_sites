<%@ LANGUAGE=VBScript CodePage=65001%>
<% response.charset="utf-8" %>
<% session.codepage=65001 %>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!-- #include file="AntiAttack.asp" -->
<!-- #include file="conn.asp" -->
<!-- #include file="md5.asp" -->
<!-- #include file="html_clear.asp" -->
<!-- #include file="Create.asp" -->
<!-- #include file="en_x_to_html/Post_index_to_html.asp" -->

<%'判断
if request("act")="add" then

article_id=request("id")
name1=trim(request.form("name"))
email1=trim(request.form("email"))
qq1=trim(request.form("qq"))
comment=trim(request.form("content"))
input_code=trim(request.form("verycode"))
url1=trim(request.form("homepage"))
image1=trim(request.form("img"))

if comment="" then
response.Write "<script language='javascript'>alert('Please input your content！');history.go(-1)</script>"
else

    if request("verycode")="" then
    response.write "<script language=javascript>alert('Your code is wrong^_^');history.go(-1);</script>"
  	Response.End 
	elseif session("getcode")="9999" then
    session("getcode")=""
	elseif session("getcode")="" then
    response.write "<script language=javascript>alert('Your code is wrong^_^');history.go(-1);</script>"
 	Response.End 
	elseif cstr(session("getcode"))<>cstr(trim(request("verycode"))) then
    response.write "<script language=javascript>alert('Your code is wrong^_^');history.go(-1);</script>"
	Response.End 
	end if

' 发布评论
set rs=server.createobject("adodb.recordset")
sql="select * from en_web_article_comment where [content]='"&nohtml(comment)&"'"
rs.open(sql),cn,1,3
if not rs.eof then  
response.Write "<script language='javascript'>alert('Bad request！');history.go(-1)</script>"
else
rs.addnew
if article_id<>"" then
rs("article_id")=article_id
end if

rs("name")=nohtml(name1)
rs("email")=nohtml(email1)
rs("qq")=nohtml(qq1)
rs("url")=nohtml(url1)
'rs("image")=image1
rs("content")=nohtml(comment)
rs("ip")=Request.serverVariables("REMOTE_ADDR")
rs("time")=now()
rs("view_yes")=0
rs.update
rs.close
set rs=nothing



'call Post_index_to_html()
'问吧模板文件夹获取
set rs_1=server.createobject("adodb.recordset")
sql="select FolderName from web_Models_type where [id]=41"
rs_1.open(sql),cn,1,1
if not rs_1.eof then
if rs_1("FolderName")<>"" then
Post_FolderName="/"&rs_1("FolderName")
end if
end if
rs_1.close
set rs_1=nothing
			response.write"<SCRIPT language=JavaScript>alert('Your message has been sent^_^');"
  response.write"location.href='/English"&Post_FolderName&"/';</SCRIPT>"
end if


end if
end if
%>
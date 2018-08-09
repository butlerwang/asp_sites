<!-- #include file="AntiAttack.asp" -->
<!-- #include file="conn.asp" -->
<!-- #include file="md5.asp" -->
<!-- #include file="html_clear.asp" -->
<!-- #include file="Create.asp" -->

<%'判断
if request("act")="add" then

article_id=request("id")
name1=trim(request.form("name"))
ordercount1=trim(request.form("ordercount"))
address1=trim(request.form("address"))
tel1=trim(request.form("tel"))
email1=trim(request.form("email"))
qq1=trim(request.form("qq"))
comment=trim(request.form("content"))
input_code=trim(request.form("verycode"))
url1=trim(request.form("homepage"))
image1=trim(request.form("img"))

if article_id="" then
response.Write "<script language='javascript'>alert('非法提交！');history.go(-1)</script>"
 	Response.End 
else

    if request("verycode")="" then
    response.write "<script language=javascript>alert('您输入的验证码有误^_^');history.go(-1);</script>"
  	Response.End 
	elseif session("getcode")="9999" then
    session("getcode")=""
	elseif session("getcode")="" then
    response.write "<script language=javascript>alert('您输入的验证码有误^_^');history.go(-1);</script>"
 	Response.End 
	elseif cstr(session("getcode"))<>cstr(trim(request("verycode"))) then
    response.write "<script language=javascript>alert('您输入的验证码有误^_^');history.go(-1);</script>"
	Response.End 
	end if

' 发布评论
set rs=server.createobject("adodb.recordset")
sql="select * from web_order where [content]='"&nohtml(comment)&"'"
rs.open(sql),cn,1,3
if not rs.eof then  
response.Write "<script language='javascript'>alert('请不要重复发布！');history.go(-1)</script>"
else
rs.addnew
if article_id<>"" then
rs("article_id")=article_id
end if

rs("name")=nohtml(name1)
rs("ordercount")=nohtml(ordercount1)
rs("address")=nohtml(address1)
rs("tel")=nohtml(tel1)
rs("email")=nohtml(email1)
rs("qq")=nohtml(qq1)
'rs("url")=nohtml(url1)
'rs("image")=image1
rs("content")=nohtml(comment)
rs("ip")=Request.serverVariables("REMOTE_ADDR")
rs("time")=now()
rs("view_yes")=0
rs.update
rs.close
set rs=nothing

			response.write"<SCRIPT language=JavaScript>alert('您的订单已提交成功，将马上安排发货^_^');"
  response.write"window.close();</SCRIPT>"
end if


end if
end if
%>
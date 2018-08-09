<%@ LANGUAGE=VBScript CodePage=65001%>
<% response.charset="utf-8" %>
<% session.codepage=65001 %>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!-- #include file="AntiAttack.asp" -->
<!-- #include file="conn.asp" -->
<!-- #include file="md5.asp" -->
<!-- #include file="html_clear.asp" -->
<!-- #include file="web_config.asp" -->
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
response.Write "<script language='javascript'>alert('bad request！');history.go(-1)</script>"
 	Response.End 
else

    if request("verycode")="" then
    response.write "<script language=javascript>alert('your code is wrong^_^');history.go(-1);</script>"
  	Response.End 
	elseif session("getcode")="9999" then
    session("getcode")=""
	elseif session("getcode")="" then
    response.write "<script language=javascript>alert('your code is wrong^_^');history.go(-1);</script>"
 	Response.End 
	elseif cstr(session("getcode"))<>cstr(trim(request("verycode"))) then
    response.write "<script language=javascript>alert('your code is wrong^_^');history.go(-1);</script>"
	Response.End 
	end if

' 发布评论
set rs=server.createobject("adodb.recordset")
sql="select * from web_order where [content]='"&nohtml(comment)&"'"
rs.open(sql),cn,1,3
if not rs.eof then  
response.Write "<script language='javascript'>alert('bad request！');history.go(-1)</script>"
else
rs.addnew
if article_id<>"" then
rs("article_id")=article_id
end if

rs("name")=nohtml(name1)
'rs("ordercount")=nohtml(ordercount1)
rs("address")=nohtml(address1)
rs("tel")=nohtml(tel1)
rs("email")=nohtml(email1)
rs("qq")=nohtml(qq1)
'rs("url")=nohtml(url1)
'rs("image")=image1
rs("content")=nohtml(comment)
ip1=Request.serverVariables("REMOTE_ADDR")
rs("ip")=ip1
rs("time")=now()
rs("view_yes")=0
rs.update
rs.close
set rs=nothing

''=========利用Jmail在线发送邮件函数 start=============
Function sendjmail(t1,t2,t3)
't1:接收邮件地址 t2:接收邮件用户名 t3:邮件正文
dim jmail
set jmail=server.createobject("Jmail.message")
jmail.silent=true
jmail.charset="gb2312"
jmail.ContentType = "text/html"

'发件人邮箱
jmail.from="movotek_web@163.com"
'发件人名称
jmail.fromname=web_name
'收件人邮箱,姓名  
jmail.AddRecipient t1,t2
'邮件的紧急程度，1最快，5最慢
jmail.Priority=1
'发送邮件标题
jmail.subject=name1&"发来询价信息"
'指定别的回信地址   
JMail.ReplyTo="taihu123_noreply@163.com"

JMail.HTMLBody = "<html><head><meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312""><style type=text/css>A:link { FONT-SIZE: 10pt; TEXT-DECORATION: none; color: #000000}A:visited {FONT-SIZE: 10pt; TEXT-DECORATION: none; color: #666666}A:hover {COLOR: #ff6600; FONT-SIZE:14pt; TEXT-DECORATION: underline}BODY {FONT-SIZE: 9pt} --></style></head><body><br>"&t3&"</body></html>"

jmail.mailserverusername="movotek_web"       '邮件发送服务器登录名称
jmail.mailserverpassword="movotek.20121214"       '邮件发送服务器登录密码
sendok=jmail.send("smtp.163.com")         'smtp服务器名称

if sendok then
'response.write "恭喜您，"&t1&"邮件发送成功！"&NOW()
else
'response.write "对不起，邮件发送失败，可能由于服务器登录设置不当或信息有误！"&NOW()
end if

jmail.Close
set jmail=nothing
End Function
''=========利用Jmail在线发送邮件函数 end=============


set rst=server.createobject("adodb.recordset")
			sql="select [title],file_path from [article] where id="&article_id&""
			rst.open(sql),cn,1,1
			if not rst.eof and not rst.bof then
			RProduct="<a href='"&web_url&"/"&Article_FolderName&"/"&rst("file_path")&"' target='_blank'>"&rst("title")&"</a>"
			end if
			rst.close
			set rst=nothing


''调用上面定义的函数发送邮件的方法
MailDetails=MailDetails&"<table cellpadding='0' cellspacing='0' width='700' align='center' style='font-family:Microsoft Yahei,Verdana,Arial;'>"
MailDetails=MailDetails&"<tr>"
MailDetails=MailDetails&"<td style='background:#4169E1;line-height:45px;font-size:16px;font-weight:bold;color:#FFFFFF;font-family:'黑体';'>&nbsp;&nbsp;"&web_name&" -- Inquire Online</td>"
MailDetails=MailDetails&"</tr>"
MailDetails=MailDetails&"<tr>"
MailDetails=MailDetails&"<td style='border:#CCCCCC 1px solid;padding:20px 20px 20px 20px;line-height:180%;font-size:13px;'>"
MailDetails=MailDetails&"<strong>您好,网站有人正在询价：</strong><br>"
MailDetails=MailDetails&"<br>"
MailDetails=MailDetails&"<strong>产品：</strong>"&RProduct&"<br>"
MailDetails=MailDetails&"<strong>姓名：</strong>"&name1&"<br>"
MailDetails=MailDetails&"<strong>IP：</strong>"&ip1&"<br>"
MailDetails=MailDetails&"<strong>联系地址：</strong>"&address1&"<br>"
MailDetails=MailDetails&"<strong>联系电话：</strong>"&tel1&"<br>"
MailDetails=MailDetails&"<strong>电子邮件：</strong>"&email1&"<br>"
MailDetails=MailDetails&"<strong>详细内容：</strong>"&comment&"<br>"
MailDetails=MailDetails&"请注意查看相关信息，及时与其取得联系。<br>"
MailDetails=MailDetails&"</td>"
MailDetails=MailDetails&"</tr>"
MailDetails=MailDetails&"<tr>"
MailDetails=MailDetails&"<td style='background:#333333;padding:10px;line-height:180%;font-size:12px;color:#FFFFFF;'>请注意：此邮件系 <a href='"&web_url&"' target='_blank' style='color:#FFFFFF;'>"&web_name&" "&web_url&"</a> 自动发送，请勿直接回复。<br>如果此邮件不是您请求的，请忽略并删除！</td>"
MailDetails=MailDetails&"</tr>"
MailDetails=MailDetails&"</table>"

response.write sendjmail("business@movotek.com",name1,MailDetails)

			response.write"<SCRIPT language=JavaScript>alert('Your Inquire information has been sent,please wait for our reply^_^');"
  response.write"location.href='/';</SCRIPT>"
end if


end if
end if
%>
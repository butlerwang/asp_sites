<!-- #include file="AntiAttack.asp" -->
<!-- #include file="conn.asp" -->
<!-- #include file="md5.asp" -->
<!-- #include file="html_clear.asp" -->
<!-- #include file="Create.asp" -->
<!-- #include file="x_to_html/Post_index_to_html.asp" -->

<%'�ж�
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
response.Write "<script language='javascript'>alert('���������ݣ�');history.go(-1)</script>"
else

    if request("verycode")="" then
    response.write "<script language=javascript>alert('���������֤������^_^');history.go(-1);</script>"
  	Response.End 
	elseif session("getcode")="9999" then
    session("getcode")=""
	elseif session("getcode")="" then
    response.write "<script language=javascript>alert('���������֤������^_^');history.go(-1);</script>"
 	Response.End 
	elseif cstr(session("getcode"))<>cstr(trim(request("verycode"))) then
    response.write "<script language=javascript>alert('���������֤������^_^');history.go(-1);</script>"
	Response.End 
	end if

' ��������
set rs=server.createobject("adodb.recordset")
sql="select * from web_article_comment where [content]='"&nohtml(comment)&"'"
rs.open(sql),cn,1,3
if not rs.eof then  
response.Write "<script language='javascript'>alert('�벻Ҫ�ظ�������');history.go(-1)</script>"
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
'�ʰ�ģ���ļ��л�ȡ
set rs_1=server.createobject("adodb.recordset")
sql="select FolderName from web_Models_type where [id]=8"
rs_1.open(sql),cn,1,1
if not rs_1.eof then
if rs_1("FolderName")<>"" then
Post_FolderName="/"&rs_1("FolderName")
end if
end if
rs_1.close
set rs_1=nothing
			response.write"<SCRIPT language=JavaScript>alert('���������Ѿ�����ɹ�,лл^_^');"
  response.write"location.href='"&Post_FolderName&"/';</SCRIPT>"
end if


end if
end if
%>
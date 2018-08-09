<% '文件夹获取
'文章文件夹获取
set rs_1=server.createobject("adodb.recordset")
sql="select FolderName from web_Models_type where [id]=39"
rs_1.open(sql),cn,1,1
if not rs_1.eof then
Article_FolderName=rs_1("FolderName")
end if
rs_1.close
set rs_1=nothing

'产品文件夹获取
set rs_1=server.createobject("adodb.recordset")
sql="select FolderName from web_Models_type where [id]=40"
rs_1.open(sql),cn,1,1
if not rs_1.eof then
Product_FolderName=rs_1("FolderName")
end if
rs_1.close
set rs_1=nothing


'案例文件夹获取
set rs_1=server.createobject("adodb.recordset")
sql="select FolderName from web_Models_type where [id]=52"
rs_1.open(sql),cn,1,1
if not rs_1.eof then
Case_FolderName=rs_1("FolderName")
end if
rs_1.close
set rs_1=nothing
%>

<%
'首页基本信息内容读取替换
set rs=server.createobject("adodb.recordset")
sql="select web_name,web_url,web_image,web_title,web_keywords,web_contact,web_TopHTML,web_BottomHTML,web_description,web_copyright,web_theme from en_web_settings"
rs.open(sql),cn,1,1
if not rs.eof and not rs.bof then
web_name=rs("web_name")
web_url=rs("web_url")
web_image=rs("web_image")
web_title=rs("web_title")
web_keywords=rs("web_keywords")
web_description=rs("web_description")
web_copyright=rs("web_copyright")
web_theme=rs("web_theme")
web_consult=rs("web_contact")
web_TopHTML=rs("web_TopHTML")
web_BottomHTML=rs("web_BottomHTML")
end if
rs.close
set rs=nothing
%>
<!-- #include file="../access.asp" -->

<%'容错处理
function ADs_to_js(l_id)
On Error Resume Next
%>
<%
'模板内容读取替换
set rs=server.createobject("adodb.recordset")
sql="select * from web_ads where id="&l_id
rs.open(sql),cn,1,1
if not rs.eof  then

select case rs("ADtype")
'文字链
case 1
ADs_Content=ADs_Content&"<a href='"&rs("url")&"' target='_blank' title='"&rs("name")&"'>"&rs("name")&"</a>"
'图片广告
case 2
ADs_Content=ADs_Content&"<a href='"&rs("url")&"' target='_blank' title='"&rs("name")&"'><img src='/images/up_images/"&rs("image")&"'  width='"&rs("ADWidth")&"' height='"&rs("ADHeight")&"' ></a>"
'Flash广告
case 3
ADs_Content=ADs_Content&"<embed src='"&rs("FlashUrl")&"' width='"&rs("ADWidth")&"' height='"&rs("ADHeight")&"'></embed>"
'case 4 广告联盟代码广告，一般是JS代码，不需要生成JS文件。
end select

end if
rs.close
set rs=nothing
%>
<% '读取模板内容
Set fso=Server.CreateObject("Scripting.FileSystemObject") 
Set htmlwrite=fso.OpenTextFile(Server.MapPath("/templates/common/template.js")) 
replace_code=htmlwrite.ReadAll() 
htmlwrite.close 
%>
<%
replace_code=replace(replace_code,"$ADs_Content$",ADs_Content)

%>
<% '判断ADs文件夹是否存在，否则创建
Set Fso=Server.CreateObject("Scripting.FileSystemObject") 
If Fso.FolderExists(Server.MapPath("/ADs"))=false Then
NewFolderDir="/ADs"
call CreateFolderB(NewFolderDir)
end if
%>

<%'声明HTML文件名,指定文件路径
filepath="/ADs/"&l_id&".js"
%>

<% '生成静态文件
Set fout = fso.CreateTextFile(Server.MapPath(filepath))
fout.WriteLine replace_code
fout.close
fso.close
set fso=nothing
end function
%>
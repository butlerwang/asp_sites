<!-- #include file="../access.asp" -->

<%'容错处理
function ADs_to_js(l_id)
On Error Resume Next
%>
<%
'模板内容读取替换
set rs=server.createobject("adodb.recordset")
sql="select * from en_web_ads where id="&l_id
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
TemplatePath="/templates/common/template.js"
replace_code=ReadFromUTF(TemplatePath,"utf-8") 
%>
<%
replace_code=replace(replace_code,"$ADs_Content$",ADs_Content)

%>
<% '判断ADs文件夹是否存在，否则创建
Set Fso=Server.CreateObject("Scripting.FileSystemObject") 
If Fso.FolderExists(Server.MapPath("/English/ADs"))=false Then
NewFolderDir="/English/ADs"
call CreateFolderB(NewFolderDir)
end if
%>

<%'声明HTML文件名,指定文件路径
filepath="/English/ADs/"&l_id&".js"
%>

<% '读取模板
'******************************************
'功能：生成UTF-8文件
'******************************************
mappath =filepath
Set objStream = Server.CreateObject("ADODB.Stream")
With objStream
.Open
.Charset = "utf-8"
.Position = objStream.Size
.WriteText=replace_code
.SaveToFile server.mappath(mappath),2
.Close
End With
Set objStream = Nothing

%>
<%
end function
%>
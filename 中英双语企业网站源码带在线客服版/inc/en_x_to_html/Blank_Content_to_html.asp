<!-- #include file="../access.asp" -->
<!-- #include file="../html_clear.asp" -->

<%'容错处理
function Blank_Content_to_html(ClassID)
On Error Resume Next
%>
<%
'首页基本信息内容读取替换
set rs=server.createobject("adodb.recordset")
sql="select web_name,web_url,web_image,web_title,web_keywords,web_contact,web_tel,web_TopHTML,web_BottomHTML,web_description,web_copyright,web_theme from en_web_settings"
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
web_tel=rs("web_tel")
end if
rs.close
set rs=nothing
%>
<% '文件夹获取
'搜索文件夹获取
set rs_1=server.createobject("adodb.recordset")
sql="select FileName,FolderName from web_Models_type where [id]=45"
rs_1.open(sql),cn,1,1
if not rs_1.eof and rs_1("FolderName")<>"" then
Search_FolderName="/English/"&rs_1("FolderName")
end if
rs_1.close
set rs_1=nothing

'文章内容文件夹获取
set rs_1=server.createobject("adodb.recordset")
sql="select FileName,FolderName from web_Models_type where [id]=39"
rs_1.open(sql),cn,1,1
if not rs_1.eof and rs_1("FolderName")<>"" then
ArticleContent_FolderName="/English/"&rs_1("FolderName")
end if
rs_1.close
set rs_1=nothing

%>

<% '读取模板内容
'模板类型获取
set rs_1=server.createobject("adodb.recordset")
sql="select FileName,FolderName from web_Models_type where [id]=43"
rs_1.open(sql),cn,1,1
if not rs_1.eof then
Model_FileName=rs_1("FileName")
if rs_1("FolderName")<>"" then
Model_FolderName="/"&rs_1("FolderName")
end if
end if
rs_1.close
set rs_1=nothing

TemplatePath="/templates/"&web_theme&"/"&Model_FileName
replace_code=ReadFromUTF(TemplatePath,"utf-8") 
%>
<%
replace_code=replace(replace_code,"$web_name$",web_name)
replace_code=replace(replace_code,"$web_url$",web_url)
replace_code=replace(replace_code,"$web_image$",web_image)
replace_code=replace(replace_code,"$web_title$",web_title)
replace_code=replace(replace_code,"$web_copyright$",web_copyright)
replace_code=replace(replace_code,"$web_theme$",web_theme)
replace_code=replace(replace_code,"$web_consult$",web_consult)
replace_code=replace(replace_code,"$web_TopHTML$",web_TopHTML)
replace_code=replace(replace_code,"$web_BottomHTML$",web_BottomHTML)
replace_code=replace(replace_code,"$web_tel$",web_tel)
replace_code=replace(replace_code,"$search_FolderName$",search_FolderName)


'顶部导航
web_TopMenu=""
set rs=server.createobject("adodb.recordset")
sql="select * from en_web_menu_type where TopNav=1 order by [order]"
rs.open(sql),cn,1,1
if not rs.eof then
web_TopMenu=web_TopMenu&"<ul id='sddm'>"
for i=1 to rs.recordcount
if i=1 then
web_TopMenu=web_TopMenu&"<li class='CurrentLi'><a href='"&rs("url")&"'>"&rs("name")&"</a></li> "
else

set rss=server.createobject("adodb.recordset")
sql="select * from en_web_menu where view_yes=1 and [position]="&rs("id")&" order by [order]"
rss.open(sql),cn,1,1
if not rss.eof then
web_TopMenu=web_TopMenu&"<li><a href='"&rs("url")&"' onmouseover=mopen('m"&i&"') onmouseout='mclosetime()'>"&rs("name")&"</a> "
web_TopMenu=web_TopMenu&"<div id='m"&i&"' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'>"
do while not rss.eof
web_TopMenu=web_TopMenu&"<a href='"&rss("url")&"'>"&rss("name")&"</a> "
rss.movenext
loop
web_TopMenu=web_TopMenu&"</div></li> "
else
web_TopMenu=web_TopMenu&"<li><a href='"&rs("url")&"'>"&rs("name")&"</a></li> "
end if
rss.close
set rss=nothing

end if
rs.movenext
next
web_TopMenu=web_TopMenu&"</ul>"
end if
rs.close
set rs=nothing

replace_code=replace(replace_code,"$web_TopMenu$",web_TopMenu)


'新闻动态
set rs=server.createobject("adodb.recordset")
sql="select top 1  [id] from [en_category] where ClassType=1 and ppid=1 order by [order]"
rs.open(sql),cn,1,1
if not rs.eof then
NewsID=rs("id")
end if
rs.close
set rs=nothing
set rs=server.createobject("adodb.recordset")
sql="select top 5 title,content,file_path,[url],time from [en_article]  where  cid='"&NewsID&"'  and view_yes=1  and ArticleType=1 order by [time] desc"
rs.open(sql),cn,1,1
if not rs.eof and not rs.bof then
Block01_LeftItem=Block01_LeftItem&"<dl>"
for i=1 to rs.recordcount
rs_url=""
if rs("url")<>"" then
rs_url=rs("url")
else
rs_url=ArticleContent_FolderName&"/"&rs("file_path")
end if 

Block01_LeftItem=Block01_LeftItem&"<dd><a href='"&rs_url&"' target='_blank' title='"&rs("title")&"'>"&left(rs("title"),60)&"</a></dd>"
rs_0.close
set rs_0=nothing
rs.movenext
next
Block01_LeftItem=Block01_LeftItem&"</dl>"
else
Block01_LeftItem=Block01_LeftItem&"No Information."
end if
rs.close
set rs=nothing

replace_code=replace(replace_code,"$Block01_LeftItem$",Block01_LeftItem)


'顶部广告读取
set rs=server.createobject("adodb.recordset")
sql="select top 1 [id],ADtype,[ADcode] from en_web_ads  where [position]=35 and view_yes=1 order by [time] desc"
rs.open(sql),cn,1,1
if not rs.eof then
if rs("ADtype")=4 then
Inner_BannerTop=Inner_BannerTop&rs("ADcode")
else
Inner_BannerTop=Inner_BannerTop&"<script src='/English/ADs/"&rs("id")&".js' type='text/javascript'></script>"
end if 
else
Inner_BannerTop=Inner_BannerTop&""
end if
rs.close
set rs=nothing

replace_code=replace(replace_code,"$Inner_BannerTop$",Inner_BannerTop)


'分类信息
set rs1=server.createobject("adodb.recordset")
sql="select [id],[pid],[ppid],[name],[title],[content],[description],[folder],[keywords] from [en_category] where [id]="&ClassID&""
rs1.open(sql),cn,1,1
if not rs1.eof then
Class_Name=rs1("name")
Class_Content=rs1("content")
Class_FolderName=rs1("folder")
Class_Keywords=rs1("keywords")
CLass_Description=rs1("Description")
Class_PPid=rs1("ppid")
if rs1("title")<>"" then
Class_Title=rs1("title")
else
Class_Title=rs1("name")
end if

Select Case Class_PPid
'一级分类情况下的当前位置
case 1
'----------------------
MainClass_FolderName="/"&rs1("folder")
class_position=""
class_position=class_position&"<a href='/English/"&rs1("folder")&"/'>"&rs1("name")&"</a>"
ClassName1=rs1("name")
ClassFolder1=rs1("folder")
ClassID1=rs1("id")
'二级分类情况下的当前位置
case 2
'--------------------
set rs_1=server.createobject("adodb.recordset")
sql="select [id],[name],[folder] from [en_category] where [id]="&rs1("pid")&" and ppid=1"
rs_1.open(sql),cn,1,1
if not rs_1.eof then
MainClass_FolderName="/"&rs_1("folder")
class_position=""
class_position=class_position&"<a href='/English/"&rs_1("folder")&"/'>"&rs_1("name")&"</a>"
class_position=class_position&" > <a href='/English/"&rs_1("folder")&"/"&rs1("folder")&"/'>"&rs1("name")&"</a>"
ClassName1=rs_1("name")
ClassFolder1=rs_1("folder")
ClassID1=rs_1("id")
end if
rs_1.close
set rs_1=nothing

'三级分类情况下的当前位置
case 3
'--------------------
set rs_1=server.createobject("adodb.recordset")
sql="select [id],[pid],[name],[folder] from [en_category] where [id]="&rs1("pid")&" and ppid=2"
rs_1.open(sql),cn,1,1
if not rs_1.eof then
set rs_1_1=server.createobject("adodb.recordset")
sql="select [id],[name],[folder] from [en_category] where [id]="&rs_1("pid")&" and ppid=1"
rs_1_1.open(sql),cn,1,1
if not rs_1_1.eof then
ClassName1=rs_1_1("name")
ClassFolder1=rs_1_1("folder")
ClassID1=rs_1_1("id")
MainClass_FolderName="/"&rs_1_1("folder")
SubClass_FolderName="/"&rs_1_1("folder")&"/"&rs_1("folder")
class_position=""
class_position=class_position&"<a href='/English/"&rs_1_1("folder")&"/'>"&rs_1_1("name")&"</a>"
end if

class_position=class_position&" > <a href='/English/"&rs_1_1("folder")&"/"&rs_1("folder")&"/'>"&rs_1("name")&"</a>"
end if

class_position=class_position&" > <a href='/English/"&rs_1_1("folder")&"/"&rs_1("folder")&"/"&rs1("folder")&"/'>"&rs1("name")&"</a>"

rs_1.close
set rs_1=nothing
rs_1_1.close
set rs_1_1=nothing
end select
end if 
rs1.close
set rs1=nothing

'侧边栏当前栏目列表
Block_LeftClassList=""
set rsl=server.createobject("adodb.recordset")
sql="select [name],[folder],[id],[pid],[ppid] from [en_category] where pid="&ClassID1&" order by [order] "
rsl.open(sql),cn,1,1
Block_LeftClassList=Block_LeftClassList&"<ul>"
if not rsl.eof then
for i=1 to rsl.recordcount
if rsl("name")=Class_Name then
Block_LeftClassList=Block_LeftClassList&"<li class='current'><A href='/English/"&ClassFolder1&"/"&rsl("Folder")&"'>"&rsl("name")&"</A></li> "
else
Block_LeftClassList=Block_LeftClassList&"<li><A href='/English/"&ClassFolder1&"/"&rsl("Folder")&"'>"&rsl("name")&"</A></li> "
end if
rsl.movenext
next
else
Block_LeftClassList=Block_LeftClassList&""
end if
Block_LeftClassList=Block_LeftClassList&"</ul>"
rsl.close
set rsl=nothing

replace_code=replace(replace_code,"$Class_Title$",Class_Title)
replace_code=replace(replace_code,"$Class_Name$",Class_Name)
replace_code=replace(replace_code,"$Class_Keywords$",Class_Keywords)
replace_code=replace(replace_code,"$Class_Description$",Class_Description)
replace_code=replace(replace_code,"$Class_Content$",Class_Content)
replace_code=replace(replace_code,"$Block_Title$",Block_Title)
replace_code=replace(replace_code,"$Block_Link$",Block_Link)
replace_code=replace(replace_code,"$category_position$",class_position)
replace_code=replace(replace_code,"$Block_LeftClassList$",Block_LeftClassList)
replace_code=replace(replace_code,"$ClassName1$",ClassName1)
replace_code=replace(replace_code,"$ClassFolder1$",ClassFolder1)
%>

<% '判断模板文件夹是否存在，否则创建
Set Fso=Server.CreateObject("Scripting.FileSystemObject") 
If Fso.FolderExists(Server.MapPath(Model_FolderName))=false Then
NewFolderDir=Model_FolderName
call CreateFolderB(NewFolderDir)
end if
%>
<% 
Select case Class_PPid
	case 1
Set Fso=Server.CreateObject("Scripting.FileSystemObject") 
If Fso.FolderExists(Server.MapPath("/English"&Model_FolderName&"/"&Class_FolderName))=false Then
NewFolderDir="/English"&Model_FolderName&"/"&Class_FolderName
call CreateFolderB(NewFolderDir)
end if
filepath_index="/English"&Model_FolderName&"/"&Class_FolderName&"/index.html"	
	case 2
Set Fso=Server.CreateObject("Scripting.FileSystemObject") 
If Fso.FolderExists(Server.MapPath("/English"&Model_FolderName&MainClass_FolderName))=false Then
NewFolderDir="/English"&Model_FolderName&MainClass_FolderName
call CreateFolderB(NewFolderDir)
end if

Set Fso=Server.CreateObject("Scripting.FileSystemObject") 
If Fso.FolderExists(Server.MapPath("/English"&Model_FolderName&MainClass_FolderName&"/"&Class_FolderName))=false Then
NewFolderDir="/English"&Model_FolderName&MainClass_FolderName&"/"&Class_FolderName
call CreateFolderB(NewFolderDir)
end if
filepath_index="/English"&Model_FolderName&MainClass_FolderName&"/"&Class_FolderName&"/index.html"
	case 3
Set Fso=Server.CreateObject("Scripting.FileSystemObject") 
If Fso.FolderExists(Server.MapPath("/English"&Model_FolderName&MainClass_FolderName))=false Then
NewFolderDir="/English"&Model_FolderName&MainClass_FolderName
call CreateFolderB(NewFolderDir)
end if

Set Fso=Server.CreateObject("Scripting.FileSystemObject") 
If Fso.FolderExists(Server.MapPath("/English"&Model_FolderName&SubClass_FolderName))=false Then
NewFolderDir="/English"&Model_FolderName&SubClass_FolderName
call CreateFolderB(NewFolderDir)
end if

Set Fso=Server.CreateObject("Scripting.FileSystemObject") 
If Fso.FolderExists(Server.MapPath("/English"&Model_FolderName&SubClass_FolderName&"/"&Class_FolderName))=false Then
NewFolderDir="/English"&Model_FolderName&SubClass_FolderName&"/"&Class_FolderName
call CreateFolderB(NewFolderDir)
end if
filepath_index="/English"&Model_FolderName&SubClass_FolderName&"/"&Class_FolderName&"/index.html" 

end select
%>
<%
'读取模板
'******************************************
'功能：生成UTF-8文件
'******************************************
mappath =filepath_index
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
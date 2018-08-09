<!-- #include file="../access.asp" -->
<!-- #include file="../html_clear.asp" -->

<%'容错处理
function SiteMap_to_html()
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

'文章内容文件夹获取
set rs_1=server.createobject("adodb.recordset")
sql="select FileName,FolderName from web_Models_type where [id]=39"
rs_1.open(sql),cn,1,1
if not rs_1.eof and rs_1("FolderName")<>"" then
ArticleContent_FolderName="/English/"&rs_1("FolderName")
end if
rs_1.close
set rs_1=nothing

'搜索文件夹获取
set rs_1=server.createobject("adodb.recordset")
sql="select FileName,FolderName from web_Models_type where [id]=45"
rs_1.open(sql),cn,1,1
if not rs_1.eof and rs_1("FolderName")<>"" then
Search_FolderName="/English/"&rs_1("FolderName")
end if
rs_1.close
set rs_1=nothing
%>

<% '读取模板内容
'模板类型获取
set rs_1=server.createobject("adodb.recordset")
sql="select FileName,FolderName from web_Models_type where [id]=46"
rs_1.open(sql),cn,1,1
if not rs_1.eof then
Model_FileName=rs_1("FileName")
if rs_1("FolderName")<>"" then
Model_FolderName="/"&rs_1("FolderName")
end if
end if
rs_1.close
set rs_1=nothing

Set fso=Server.CreateObject("Scripting.FileSystemObject") 
Set htmlwrite=fso.OpenTextFile(Server.MapPath("/templates/"&web_theme&"/"&Model_FileName)) 
replace_code=htmlwrite.ReadAll() 
htmlwrite.close 
%>
<%
replace_code=replace(replace_code,"$web_name$",web_name)
replace_code=replace(replace_code,"$web_url$",web_url)
replace_code=replace(replace_code,"$web_image$",web_image)
replace_code=replace(replace_code,"$web_title$",web_title)
replace_code=replace(replace_code,"$web_keywords$",web_keywords)
replace_code=replace(replace_code,"$web_description$",web_description)
replace_code=replace(replace_code,"$web_copyright$",web_copyright)
replace_code=replace(replace_code,"$web_theme$",web_theme)
replace_code=replace(replace_code,"$web_consult$",web_consult)
replace_code=replace(replace_code,"$web_TopHTML$",web_TopHTML)
replace_code=replace(replace_code,"$web_BottomHTML$",web_BottomHTML)
replace_code=replace(replace_code,"$web_tel$",web_tel)
replace_code=replace(replace_code,"$search_FolderName$",search_FolderName)


'顶部导航
set rs=server.createobject("adodb.recordset")
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


'顶部广告读取
set rs=server.createobject("adodb.recordset")
sql="select top 1 [id],ADtype,[ADcode] from en_web_ads  where [position]=35 and view_yes=1 order by [time] desc"
rs.open(sql),cn,1,1
if not rs.eof then
if rs("ADtype")=4 then
Inner_BannerTop=Inner_BannerTop&rs("ADcode")
else
Inner_BannerTop=Inner_BannerTop&"<script src='/ADs/"&rs("id")&".js' type='text/javascript'></script>"
end if 
else
Inner_BannerTop=Inner_BannerTop&""
end if
rs.close
set rs=nothing

replace_code=replace(replace_code,"$Inner_BannerTop$",Inner_BannerTop)


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


'网站地图
SiteMap_List=""
set rsl=server.createobject("adodb.recordset")
sql="select [name],[folder],[id],[pid],[ppid] from [en_category] where ppid=1 order by [order] "
rsl.open(sql),cn,1,1
if not rsl.eof then
do while not rsl.eof
SiteMap_List=SiteMap_List&"<li><A href='/English/"&rsl("Folder")&"'>"&rsl("name")&"</A> "

set rs=server.createobject("adodb.recordset")
sql="select [id],[name],[folder] from [en_category] where ppid=2 and pid="&rsl("id")&" "
rs.open(sql),cn,1,1
if not rs.eof then
SiteMap_List=SiteMap_List&"<ul>"
do while not rs.eof 
SiteMap_List=SiteMap_List&"<li><a href='/English/"&rsl("Folder")&"/"&rs("folder")&"/' >"&rs("name")&"</a>"
set rs3=server.createobject("adodb.recordset")
sql="select  [name],[folder] from [en_category] where ppid=3 and pid="&rs("id")&" "
rs3.open(sql),cn,1,1
if not rs3.eof then
do while not rs3.eof 
SiteMap_List=SiteMap_List&"&nbsp;&nbsp;&nbsp;&nbsp;<a href='/English/"&rsl("Folder")&"/"&rs("folder")&"/"&rs3("folder")&"/' >"&rs3("name")&"</a>"
rs3.movenext
loop
end if
rs3.close
set rs3=nothing
SiteMap_List=SiteMap_List&"</li> "
rs.movenext
loop
SiteMap_List=SiteMap_List&"</ul>"
end if
rs.close
set rs=nothing
SiteMap_List=SiteMap_List&"</li> "
rsl.movenext
loop
end if
rsl.close
set rsl=nothing

replace_code=replace(replace_code,"$SiteMap_List$",SiteMap_List)






%>

<% '判断文件夹是否存在，否则创建
Set Fso=Server.CreateObject("Scripting.FileSystemObject") 
If Fso.FolderExists(Server.MapPath("/English"&Model_FolderName))=false Then
NewFolderDir="/English"&Model_FolderName
call CreateFolderB(NewFolderDir)
end if
%>

<%'声明HTML文件名,指定文件路径
filepath="/English"&Model_FolderName&"/index.html"
%>

<% '生成首页静态文件
'读取模板
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
end function
%>
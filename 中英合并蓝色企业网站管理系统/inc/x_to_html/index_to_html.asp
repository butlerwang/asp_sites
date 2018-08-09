<!-- #include file="../access.asp" -->
<!-- #include file="../html_clear.asp" -->

<%'容错处理
function index_to_html()
On Error Resume Next
%>
<%
'首页基本信息内容读取替换
set rs=server.createobject("adodb.recordset")
sql="select web_name,web_url,web_image,web_title,web_keywords,web_contact,web_tel,web_TopHTML,web_BottomHTML,web_description,web_copyright,web_theme from web_settings"
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
sql="select FileName,FolderName from web_Models_type where [id]=5"
rs_1.open(sql),cn,1,1
if not rs_1.eof and rs_1("FolderName")<>"" then
ArticleContent_FolderName="/"&rs_1("FolderName")
end if
rs_1.close
set rs_1=nothing

'产品内容文件夹获取
set rs_1=server.createobject("adodb.recordset")
sql="select FileName,FolderName from web_Models_type where [id]=6"
rs_1.open(sql),cn,1,1
if not rs_1.eof and rs_1("FolderName")<>"" then
ProductContent_FolderName="/"&rs_1("FolderName")
end if
rs_1.close
set rs_1=nothing
%>

<% '读取模板内容
'模板类型获取
set rs_1=server.createobject("adodb.recordset")
sql="select FileName,FolderName from web_Models_type where [id]=1"
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


'顶部导航
set rs=server.createobject("adodb.recordset")
sql="select * from web_menu_type where TopNav=1 order by [order]"
rs.open(sql),cn,1,1
if not rs.eof then
web_TopMenu=web_TopMenu&"<ul id='sddm'>"
for i=1 to rs.recordcount
if i=1 then
web_TopMenu=web_TopMenu&"<li class='CurrentLi'><a href='"&rs("url")&"'>"&rs("name")&"<p>"&rs("EnName")&"</p></a></li> "
else

set rss=server.createobject("adodb.recordset")
sql="select * from web_menu where view_yes=1 and [position]="&rs("id")&" order by [order]"
rss.open(sql),cn,1,1
if not rss.eof then
web_TopMenu=web_TopMenu&"<li><a href='"&rs("url")&"' onmouseover=mopen('m"&i&"') onmouseout='mclosetime()'>"&rs("name")&"<p>"&rs("EnName")&"</p></a> "
web_TopMenu=web_TopMenu&"<div id='m"&i&"' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'>"
do while not rss.eof
web_TopMenu=web_TopMenu&"<a href='"&rss("url")&"'>"&rss("name")&"</a> "
rss.movenext
loop
web_TopMenu=web_TopMenu&"</div></li> "
else
web_TopMenu=web_TopMenu&"<li><a href='"&rs("url")&"'>"&rs("name")&"<p>"&rs("EnName")&"</p></a></li> "
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



'首页顶部幻灯广告读取替换
set rs=server.createobject("adodb.recordset")
sql="select top 5 name,url,image from web_ads  where [position]=36 and view_yes=1 order by [order]"
rs.open(sql),cn,1,1
if not rs.eof then
for i=1 to rs.recordcount
web_TopIMGAD=web_TopIMGAD&"<li><a href='"&rs("url")&"' target='_blank'><img src='/images/up_images/"&rs("image")&"' alt='"&rs("name")&"'/></a></li>"
rs.movenext
next
else
web_TopIMGAD=web_TopIMGAD&"无数据"
end if
rs.close
set rs=nothing

replace_code=replace(replace_code,"$web_TopIMGAD$",web_TopIMGAD)


'企业介绍
set rs=server.createobject("adodb.recordset")
sql="select top 1  [name],[folder],[id],[pid],[ppid],[image],[content] from [category] where ClassType=5 and ppid=1 order by [order]"
rs.open(sql),cn,1,1
if not rs.eof then
ItemID=rs("id")
WebAboutName=rs("name")
WebAboutFolderName=rs("folder")
WebAboutImage=rs("image")
WebAboutContent=left(ClearHtml(rs("content")),90)

select case rs("ppid")
case 1
ClassSQL="cid"
WebAboutFolder="/"&rs("folder") 
case 2
ClassSQL="pid"
set rs0=server.createobject("adodb.recordset")
sql="select [name],[folder],[id],[pid],[ppid] from [category] where id="&rs("pid")
rs0.open(sql),cn,1,1
if not rs0.eof then
WebAboutFolder="/"&rs0("folder")&"/"&rs("folder")
end if
rs0.close
set rs0=nothing
case 3
ClassSQL="ppid"
set rs0=server.createobject("adodb.recordset")
sql="select [name],[folder],[id],[pid],[ppid] from [category] where id="&rs("pid")
rs0.open(sql),cn,1,1
if not rs0.eof then
set rs00=server.createobject("adodb.recordset")
sql="select [name],[folder],[id],[pid],[ppid] from [category] where id="&rs0("pid")
rs00.open(sql),cn,1,1
if not rs00.eof then
WebAboutFolder="/"&rs00("folder")&"/"&rs0("folder")&"/"&rs("folder") 
end if
rs00.close
set rs00=nothing
end if
rs0.close
set rs0=nothing
end select
end if
rs.close
set rs=nothing

replace_code=replace(replace_code,"$WebAboutName$",WebAboutName)
replace_code=replace(replace_code,"$WebAboutFolderName$",WebAboutFolderName)
replace_code=replace(replace_code,"$WebAboutFolder$",WebAboutFolder)
replace_code=replace(replace_code,"$WebAboutImage$",WebAboutImage)
replace_code=replace(replace_code,"$WebAboutContent$",WebAboutContent)



'新闻动态
set rs=server.createobject("adodb.recordset")
sql="select top 1  [name],[folder],[id],[pid],[ppid] from [category] where ClassType=1 and ppid=1 order by [order]"
rs.open(sql),cn,1,1
if not rs.eof then
ItemID=rs("id")
WebNewNewsName=rs("name")
WebNewNewsFolderName=rs("folder")
select case rs("ppid")
case 1
ClassSQL="cid"
WebNewNewsFolder="/"&rs("folder") 
case 2
ClassSQL="pid"
set rs0=server.createobject("adodb.recordset")
sql="select [name],[folder],[id],[pid],[ppid] from [category] where id="&rs("pid")
rs0.open(sql),cn,1,1
if not rs0.eof then
WebNewNewsFolder="/"&rs0("folder")&"/"&rs("folder")
end if
rs0.close
set rs0=nothing
case 3
ClassSQL="ppid"
set rs0=server.createobject("adodb.recordset")
sql="select [name],[folder],[id],[pid],[ppid] from [category] where id="&rs("pid")
rs0.open(sql),cn,1,1
if not rs0.eof then
set rs00=server.createobject("adodb.recordset")
sql="select [name],[folder],[id],[pid],[ppid] from [category] where id="&rs0("pid")
rs00.open(sql),cn,1,1
if not rs00.eof then
WebNewNewsFolder="/"&rs00("folder")&"/"&rs0("folder")&"/"&rs("folder") 
end if
rs00.close
set rs00=nothing
end if
rs0.close
set rs0=nothing
end select
end if
rs.close
set rs=nothing
'content

set rs=server.createobject("adodb.recordset")
sql="select top 9 title,content,file_path,[url],time from [article]  where   "&ClassSQL&"='"&ItemID&"'  and view_yes=1  and ArticleType=1 order by [time] desc"
rs.open(sql),cn,1,1
if not rs.eof and not rs.bof then
for i=1 to 9
rs_url=""
if rs("url")<>"" then
rs_url=rs("url")
else
rs_url=ArticleContent_FolderName&"/"&rs("file_path")
end if 
WebNewNewsList=WebNewNewsList&"<tr><td width='75%' class='ListTitle'><a href='"&rs_url&"' target='_blank' title='"&rs("title")&"'>"&left(rs("title"),25)&"</a></td>"
WebNewNewsList=WebNewNewsList&"<td width='25%'><span>"&year(rs("time"))&" / "&month(rs("time"))&" / "&day(rs("time"))&"</span></td></tr>"
rs_0.close
set rs_0=nothing
rs.movenext
next
else
WebNewNewsList=WebNewNewsList&"暂无信息。"
end if
rs.close
set rs=nothing

replace_code=replace(replace_code,"$WebNewNewsList$",WebNewNewsList)
replace_code=replace(replace_code,"$WebNewNewsName$",WebNewNewsName)
replace_code=replace(replace_code,"$WebNewNewsFolderName$",WebNewNewsFolderName)
replace_code=replace(replace_code,"$WebNewNewsFolder$",WebNewNewsFolder)



'品牌产品
set rs=server.createobject("adodb.recordset")
sql="select top 1  [name],[folder],[id],[pid],[ppid] from [category] where ClassType=2 and ppid=1 order by [order]"
rs.open(sql),cn,1,1
if not rs.eof then
ItemID=rs("id")
WebProductName=rs("name")
WebProductFolderName=rs("folder")
WebProductFolder="/"&rs("folder")
end if
rs.close
set rs=nothing

set rs=server.createobject("adodb.recordset")
sql="select top 8 [name],[folder],[id],[pid],[ppid] from [category] where pid="&ItemID&" and ClassType=2 and ppid=2 order by [order] "
rs.open(sql),cn,1,1
if not rs.eof then
for i=1 to 8
Block03_LeftItem_Title=Block03_LeftItem_Title&"<li class='Tabs"&i&"'>"
Block03_LeftItem_Title=Block03_LeftItem_Title&"<A onmousemove='easytabs(1, "&i&");' onfocus='easytabs(1, "&i&");'  "
Block03_LeftItem_Title=Block03_LeftItem_Title&"title='"&rs("name")&"' "
Block03_LeftItem_Title=Block03_LeftItem_Title&"id='tablink"&i&"' href='"&WebProductFolderName&"/"
Block03_LeftItem_Title=Block03_LeftItem_Title&rs("Folder")
Block03_LeftItem_Title=Block03_LeftItem_Title&"'>"
Block03_LeftItem_Title=Block03_LeftItem_Title&rs("name")
Block03_LeftItem_Title=Block03_LeftItem_Title&"</A>"
Block03_LeftItem_Title=Block03_LeftItem_Title&"</li> "


Block03_LeftItem=Block03_LeftItem&"<div id='tabcontent"&i&"'>"
Block03_LeftItem=Block03_LeftItem&"<DIV class='blk_29'>"
Block03_LeftItem=Block03_LeftItem&"<DIV class='LeftBotton' id='LeftArr"&i&"'></DIV>"
Block03_LeftItem=Block03_LeftItem&"<DIV class='Cont' id='ISL_Cont_"&i&"'>"

set rsp=server.createobject("adodb.recordset")
sql="select top 12 [title],file_path,[image] from [article] where ArticleType=2 and cid='"&ItemID&"'  and pid='"&rs("id")&"' and view_yes=1 order by [time] desc"
rsp.open(sql),cn,1,1
if not rsp.eof then
do while not rsp.eof 
rs_url=""
rs_url=ProductContent_FolderName&"/"&rsp("file_path")
Block03_LeftItem=Block03_LeftItem&"<DIV class='box'><A class='imgBorder'  href='"&rs_url&"' target='_blank' title='"&rsp("title")&"'><IMG alt='"&rsp("title")&"' src='/images/up_images/"&rsp("image")&"'><p>"&left(rsp("title"),14)&"</p></A> </DIV>"
rsp.movenext
loop
end if 
rsp.close
set rsp=nothing

Block03_LeftItem=Block03_LeftItem&"</DIV>"
Block03_LeftItem=Block03_LeftItem&"<DIV class='RightBotton' id='RightArr"&i&"'></DIV></DIV>"
Block03_LeftItem=Block03_LeftItem&"<SCRIPT language='javascript' type='text/javascript'>"
Block03_LeftItem=Block03_LeftItem&"var scrollPic_02 = new ScrollPic();"
Block03_LeftItem=Block03_LeftItem&"scrollPic_02.scrollContId   = 'ISL_Cont_"&i&"';"
Block03_LeftItem=Block03_LeftItem&"scrollPic_02.arrLeftId      = 'LeftArr"&i&"';"
Block03_LeftItem=Block03_LeftItem&"scrollPic_02.arrRightId     = 'RightArr"&i&"'; "
Block03_LeftItem=Block03_LeftItem&"scrollPic_02.frameWidth     = 888;"
Block03_LeftItem=Block03_LeftItem&"scrollPic_02.pageWidth      = 888; "
Block03_LeftItem=Block03_LeftItem&"scrollPic_02.speed          = 3; "
Block03_LeftItem=Block03_LeftItem&"scrollPic_02.space          = 50; "
Block03_LeftItem=Block03_LeftItem&"scrollPic_02.autoPlay       = true; "
Block03_LeftItem=Block03_LeftItem&"scrollPic_02.autoPlayTime   = 3; "
Block03_LeftItem=Block03_LeftItem&"scrollPic_02.initialize(); "
Block03_LeftItem=Block03_LeftItem&"</SCRIPT>"
Block03_LeftItem=Block03_LeftItem&"<div class='clearfix'></div> </div> "

rs.movenext
next
end if
rs.close
set rs=nothing



replace_code=replace(replace_code,"$Block03_LeftItem$",Block03_LeftItem)
replace_code=replace(replace_code,"$Block03_LeftItem_Title$",Block03_LeftItem_Title)
replace_code=replace(replace_code,"$WebProductName$",WebProductName)
replace_code=replace(replace_code,"$WebProductFolderName$",WebProductFolderName)
replace_code=replace(replace_code,"$WebProductFolder$",WebProductFolder)




'友情链接
set rs=server.createobject("adodb.recordset")
sql="select  [name],[url],[image],follow_yes from [web_link] where view_yes=1 order by [order]"
rs.open(sql),cn,1,1
if not rs.eof then
do while not rs.eof
if rs("follow_yes")=1 then
NoFollow="rel='nofollow'"
else
NoFollow=""
end if 
web_link=web_link&"<a href='"&rs("url")&"' target='_blank' "&NoFollow&">"&rs("name")&"</a>"
rs.movenext
loop
else
web_link=web_link&"暂无链接。"
end if
rs.close
set rs=nothing

replace_code=replace(replace_code,"$web_link$",web_link)
%>

<% '判断文件夹是否存在，否则创建
Set Fso=Server.CreateObject("Scripting.FileSystemObject") 
If Fso.FolderExists(Server.MapPath(Model_FolderName))=false Then
NewFolderDir=Model_FolderName
call CreateFolderB(NewFolderDir)
end if
%>

<%'声明HTML文件名,指定文件路径
filepath=Model_FolderName&"/index.html"
%>

<% '生成首页静态文件
Set fout = fso.CreateTextFile(Server.MapPath(filepath))
fout.WriteLine replace_code
fout.close
fso.close
set fso=nothing
end function
%>
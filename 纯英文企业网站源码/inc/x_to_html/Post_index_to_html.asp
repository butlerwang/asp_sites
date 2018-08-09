<!-- #include file="../html_clear.asp" -->
<%'容错处理
function Post_index_to_html()
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
web_consult=rs("web_TopHTML")
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
sql="select FileName,FolderName from web_Models_type where [id]=35"
rs_1.open(sql),cn,1,1
if not rs_1.eof and rs_1("FolderName")<>"" then
Search_FolderName="/"&rs_1("FolderName")
end if
rs_1.close
set rs_1=nothing

'文章内容文件夹获取
set rs_1=server.createobject("adodb.recordset")
sql="select FileName,FolderName from web_Models_type where [id]=5"
rs_1.open(sql),cn,1,1
if not rs_1.eof and rs_1("FolderName")<>"" then
ArticleContent_FolderName="/"&rs_1("FolderName")
end if
rs_1.close
set rs_1=nothing


'问吧模板类型获取
set rs_1=server.createobject("adodb.recordset")
sql="select FileName,FolderName from web_Models_type where [id]=8"
rs_1.open(sql),cn,1,1
if not rs_1.eof then
Model_FileName=rs_1("FileName")
if rs_1("FolderName")<>"" then
Model_FolderName="/"&rs_1("FolderName")
end if
end if
rs_1.close
set rs_1=nothing
%>
<%
'顶部导航
web_TopMenu=""
set rs=server.createobject("adodb.recordset")
sql="select * from web_menu_type where TopNav=1 order by [order]"
rs.open(sql),cn,1,1
if not rs.eof then
web_TopMenu=web_TopMenu&"<ul id='sddm'>"
for i=1 to rs.recordcount
if i=1 then
web_TopMenu=web_TopMenu&"<li class='CurrentLi'><a href='"&rs("url")&"'>"&rs("name")&"</a></li> "
else

set rss=server.createobject("adodb.recordset")
sql="select * from web_menu where view_yes=1 and [position]="&rs("id")&" order by [order]"
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


'新闻动态
set rs=server.createobject("adodb.recordset")
sql="select top 1  [id] from [category] where ClassType=1 and ppid=1 order by [order]"
rs.open(sql),cn,1,1
if not rs.eof then
NewsID=rs("id")
end if
rs.close
set rs=nothing
set rs=server.createobject("adodb.recordset")
sql="select top 5 title,content,file_path,[url],time from [article]  where  cid='"&NewsID&"'  and view_yes=1  and ArticleType=1 order by [time] desc"
rs.open(sql),cn,1,1
if not rs.eof and not rs.bof then
Block01_LeftItem=Block01_LeftItem&"<dl>"
for i=1 to 5
rs_url=""
if rs("url")<>"" then
rs_url=rs("url")
else
rs_url=ArticleContent_FolderName&"/"&rs("file_path")
end if 
Block01_LeftItem=Block01_LeftItem&"<dt>"&year(rs("time"))&"/"&month(rs("time"))&"/"&day(rs("time"))&"</dt>"
Block01_LeftItem=Block01_LeftItem&"<dd><a href='"&rs_url&"' target='_blank' title='"&rs("title")&"'>"&left(rs("title"),14)&"</a></dd>"
rs_0.close
set rs_0=nothing
rs.movenext
next
Block01_LeftItem=Block01_LeftItem&"</dl>"
else
Block01_LeftItem=Block01_LeftItem&"No Data."
end if
rs.close
set rs=nothing


'顶部广告读取
set rs=server.createobject("adodb.recordset")
sql="select top 1 [id],ADtype,[ADcode] from web_ads  where [position]=35 and view_yes=1 order by [time] desc"
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


'侧边栏当前栏目列表
set rs=server.createobject("adodb.recordset")
sql="select [id],[name],[folder] from [category] where ClassType=5 and ppid=1 order by [order]"
rs.open(sql),cn,1,1
if not rs.eof then
ClassID=rs("id")
ClassName1=rs("name")
ClassFolder1=rs("folder")
Block_LeftClassList=""
set rsl=server.createobject("adodb.recordset")
sql="select [name],[folder],[id],[pid],[ppid] from [category] where pid="&ClassID&" order by [order] "
rsl.open(sql),cn,1,1
Block_LeftClassList=Block_LeftClassList&"<ul>"
if not rsl.eof then
for i=1 to rsl.recordcount
if rsl("id")=ClassID then
Block_LeftClassList=Block_LeftClassList&"<li class='current'><A href='/"&ClassFolder1&"/"&rsl("Folder")&"'>"&rsl("name")&"</A></li> "
else
Block_LeftClassList=Block_LeftClassList&"<li><A href='/"&ClassFolder1&"/"&rsl("Folder")&"'>"&rsl("name")&"</A></li> "
end if
rsl.movenext
next
else
Block_LeftClassList=Block_LeftClassList&""
end if
Block_LeftClassList=Block_LeftClassList&"</ul>"
rsl.close
set rsl=nothing

end if
rs.close
set rs=nothing
%>
<!--common use end-->

<%
 '列表读取替换
sql="select [id] from web_article_comment where view_yes=1 and article_id=0  order by [time]"
set rs1=server.createObject("ADODB.Recordset")
rs1.open sql,cn,1,1
%><%
if not rs1.eof then
rs1.pagesize=5
totalpage=rs1.pagecount

for j=1 to totalpage
rs_order=1
sql="select * from web_article_comment where view_yes=1 and article_id=0  order by [time]"
set rs=server.createObject("ADODB.Recordset")
rs.open sql,cn,1,1
if not rs.eof then

rscount=rs.recordcount
whichpage=j 
rs.pagesize=5
totalpage=rs.pagecount
rs.absolutepage=whichpage
howmanyrecs=0
list_block=""
do while not rs.eof and howmanyrecs<rs.pagesize
%><%
if rs("name")<>"" then
comment_name=rs("name")
else
comment_name=rs("ip")
end if

if rs("recontent")<>"" then
comment_replay="<div class='Freply'><div class='FRtitle'>Reply</div><p>"&rs("recontent")&"</p></div>"
else
comment_replay=""
end if
rs_Pagesize=(j-1)*rs.pagesize
rs_order=rs_order+rs_pagesize

list_block=list_block&"<div class='FeedBlock'><div class='Fleft'>"
list_block=list_block&"<div class='Ficon'><img src='/images/PostIcon.gif'></div>"
list_block=list_block&"<div class='Fname'>"&rs("name")&"</div></div>"
list_block=list_block&"<div class='Fright'><div class='Fcontent'>"
list_block=list_block&"<div class='Ftime'>"&rs("time")&"</div>"
list_block=list_block&"<p>"&rs("content")&"</p>"
list_block=list_block&comment_replay
list_block=list_block&"</div><div class='Fline'></div><div class='clearfix'></div>"
list_block=list_block&"</div></div><div class='clearfix'></div> "
%>
<%
rs.movenext
rs_order=rs_order+1-rs_Pagesize
howmanyrecs=howmanyrecs+1
loop
end if
rs.close
set rs=nothing
%>

<%
'分页部分
list_page=""
list_page=list_page&"<div class='t_page ColorLink'>"
list_page=list_page&"Total: "&rscount&"&nbsp;&nbsp;Current Page: <span class='FontRed'>" & j & "</span>/" & totalpage &""
list_page=list_page&"<a href=index.html"&">" & "First Page" & "</a>"
if whichpage=1 then
list_page=list_page&"&nbsp;&nbsp;Pre Page&nbsp;&nbsp;"
else
if j=2  then
list_page=list_page&"<a href=index.html"&">" & "Pre Page" & "</a>"
else
list_page=list_page&"<a href=list_"&j-1&".html"&">" & "Pre Page" & "</a>"
end if
end if

if totalpage-j>4 then
PageNO=j+4
else
PageNO=totalpage
end if

for counter=j to PageNO
if counter=1 then
list_page=list_page&"<a href=index.html"&">" & counter & "</a> "
else
if counter=j then
list_page=list_page&"<a href=list_"&counter&".html"&"><span class='FontRed'>" & counter & "</span></a> "
else
list_page=list_page&"<a href=list_"&counter&".html"&">" & counter & "</a> "
end if
end if
next

if (whichpage>totalpage or whichpage=totalpage) then
list_page=list_page&"&nbsp;&nbsp;Next&nbsp;&nbsp;"
else
list_page=list_page&"<a href=list_"&j+1&".html"&">" & "Next" & "</a>"
end if
if totalpage=1 then
list_page=list_page&"<a href=index.html"&">" & "Last Page" & "</a></div>"
else
list_page=list_page&"<a href=list_"&totalpage&".html"&">" & "Last Page" & "</a></div>"
end if
%>

<%
 '读取模板内容
TemplatePath="/templates/"&web_theme&"/"&Model_FileName
replace_code=ReadFromUTF(TemplatePath,"utf-8") %>

<%'循环列表替换内容
replace_code=replace(replace_code,"$list_block$",list_block)  
replace_code=replace(replace_code,"$list_page$",list_page)    
replace_code=replace(replace_code,"$RScount$",RScount)   


%>
<%'循环其它替换内容
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

replace_code=replace(replace_code,"$web_link$",web_link)
replace_code=replace(replace_code,"$InnerAD_Bottom$",InnerAD_Bottom)
replace_code=replace(replace_code,"$InnerAD_Top$",InnerAD_Top)
replace_code=replace(replace_code,"$web_BottomNav$",web_BottomNav)
replace_code=replace(replace_code,"$web_menu$",web_menu)
replace_code=replace(replace_code,"$web_indexsearch$",web_indexsearch)
replace_code=replace(replace_code,"$Block_LeftClassList$",Block_LeftClassList)
replace_code=replace(replace_code,"$ClassName1$",ClassName1)
replace_code=replace(replace_code,"$ClassFolder1$",ClassFolder1)

replace_code=replace(replace_code,"$web_TopMenu$",web_TopMenu)
replace_code=replace(replace_code,"$Inner_BannerTop$",Inner_BannerTop)
replace_code=replace(replace_code,"$Block01_LeftItem$",Block01_LeftItem)
replace_code=replace(replace_code,"$Block02_LeftItem$",Block02_LeftItem)
%>

<% '判断文件夹是否存在，否则创建
Set Fso=Server.CreateObject("Scripting.FileSystemObject") 
If Fso.FolderExists(Server.MapPath(Model_FolderName))=false Then
NewFolderDir=Model_FolderName
call CreateFolderB(NewFolderDir)
end if
%>
<%
filepath=Model_FolderName&"/list_"&j&".html"
filepath_index=Model_FolderName&"/index.html"
%>
<%
if j>1 then
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
end if

if j=1 then
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
end if
%>
<%
next
else
j=1
%>
<%
 '读取模板内容
TemplatePath="/templates/"&web_theme&"/"&Model_FileName
replace_code=ReadFromUTF(TemplatePath,"utf-8") 
%>

<%'循环列表替换内容
replace_code=replace(replace_code,"$list_block$","<p align='center'></p>") 
replace_code=replace(replace_code,"$list_page$",list_page)      
replace_code=replace(replace_code,"$RScount$","0")   
%>
<%'循环其它替换内容
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

replace_code=replace(replace_code,"$web_link$",web_link)
replace_code=replace(replace_code,"$InnerAD_Bottom$",InnerAD_Bottom)
replace_code=replace(replace_code,"$InnerAD_Top$",InnerAD_Top)
replace_code=replace(replace_code,"$web_BottomNav$",web_BottomNav)
replace_code=replace(replace_code,"$web_menu$",web_menu)
replace_code=replace(replace_code,"$web_indexsearch$",web_indexsearch)
replace_code=replace(replace_code,"$Block_LeftClassList$",Block_LeftClassList)
replace_code=replace(replace_code,"$ClassName1$",ClassName1)
replace_code=replace(replace_code,"$ClassFolder1$",ClassFolder1)

replace_code=replace(replace_code,"$web_TopMenu$",web_TopMenu)
replace_code=replace(replace_code,"$Inner_BannerTop$",Inner_BannerTop)
replace_code=replace(replace_code,"$Block01_LeftItem$",Block01_LeftItem)
replace_code=replace(replace_code,"$Block02_LeftItem$",Block02_LeftItem)
%>

<% '判断文件夹是否存在，否则创建
Set Fso=Server.CreateObject("Scripting.FileSystemObject") 
If Fso.FolderExists(Server.MapPath(Model_FolderName))=false Then
NewFolderDir=Model_FolderName
call CreateFolderB(NewFolderDir)
end if
%>
<%
filepath_index=Model_FolderName&"/index.html"
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



<%end if
rs1.close
set rs1=nothing%>

<%end function%>
<!-- #include file="../access.asp" -->
<!-- #include file="../html_clear.asp" -->
<%'�ݴ���
function Case_List_to_html(ClassID)
On Error Resume Next
%>
<%
'��ҳ������Ϣ���ݶ�ȡ�滻
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

<% '�ļ��л�ȡ
'�����ļ��л�ȡ
set rs_1=server.createobject("adodb.recordset")
sql="select FileName,FolderName from web_Models_type where [id]=35"
rs_1.open(sql),cn,1,1
if not rs_1.eof and rs_1("FolderName")<>"" then
Search_FolderName="/"&rs_1("FolderName")
end if
rs_1.close
set rs_1=nothing

'���������ļ��л�ȡ
set rs_1=server.createobject("adodb.recordset")
sql="select FileName,FolderName from web_Models_type where [id]=5"
rs_1.open(sql),cn,1,1
if not rs_1.eof and rs_1("FolderName")<>"" then
ArticleContent_FolderName=web_url&"/"&rs_1("FolderName")
end if
rs_1.close
set rs_1=nothing

'��Ʒ�����ļ��л�ȡ
set rs_1=server.createobject("adodb.recordset")
sql="select FileName,FolderName from web_Models_type where [id]=6"
rs_1.open(sql),cn,1,1
if not rs_1.eof and rs_1("FolderName")<>"" then
ProductContent_FolderName=web_url&"/"&rs_1("FolderName")
end if
rs_1.close
set rs_1=nothing

'�б�ģ�����ͻ�ȡ
set rs_1=server.createobject("adodb.recordset")
sql="select FileName,FolderName from web_Models_type where [id]=24"
rs_1.open(sql),cn,1,1
if not rs_1.eof then
Model_FileName=rs_1("FileName")
Model_FolderName=rs_1("FolderName")
end if
rs_1.close
set rs_1=nothing
%>
<%
'��������
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

'�ײ�����
web_BottomMenu=""
set rs=server.createobject("adodb.recordset")
sql="select * from web_menu_type where BottomNav=1 order by [order]"
rs.open(sql),cn,1,1
if not rs.eof then
for i=1 to rs.recordcount

set rss=server.createobject("adodb.recordset")
sql="select * from web_menu where view_yes=1 and [position]="&rs("id")&" order by [order]"
rss.open(sql),cn,1,1
if not rss.eof then
web_BottomMenu=web_BottomMenu&"<div class='box_240px_left'><div class='post'><h2>"&rs("name")&"</span> </h2> "
web_BottomMenu=web_BottomMenu&"<ul>"
do while not rss.eof
web_BottomMenu=web_BottomMenu&"<li><a href='"&rss("url")&"'>"&rss("name")&"</a></li> "
rss.movenext
loop
web_BottomMenu=web_BottomMenu&"</ul></div></div> "
rss.close
set rss=nothing

end if
rs.movenext
next
end if
rs.close
set rs=nothing

'�ȵ��Ʒ
set rs=server.createobject("adodb.recordset")
sql="select top 1  [id] from [category] where ClassType=2 and ppid=1 order by [order]"
rs.open(sql),cn,1,1
if not rs.eof then
NewsID=rs("id")
end if
rs.close
set rs=nothing
set rs=server.createobject("adodb.recordset")
sql="select top 5 title,content,file_path,[url],time from [article]  where  cid='"&NewsID&"'  and view_yes=1  and ArticleType=2 order by [Hit] desc"
rs.open(sql),cn,1,1
if not rs.eof and not rs.bof then
Block01_LeftItem=Block01_LeftItem&"<dl>"
for i=1 to 5
rs_url=""
if rs("url")<>"" then
rs_url=rs("url")
else
rs_url=ProductContent_FolderName&"/"&rs("file_path")
end if 
Block01_LeftItem=Block01_LeftItem&"<dt>"&year(rs("time"))&"/"&month(rs("time"))&"/"&day(rs("time"))&"</dt>"
Block01_LeftItem=Block01_LeftItem&"<dd><a href='"&rs_url&"' target='_blank' title='"&rs("title")&"'>"&left(rs("title"),14)&"</a></dd>"
rs_0.close
set rs_0=nothing
rs.movenext
next
Block01_LeftItem=Block01_LeftItem&"</dl>"
else
Block01_LeftItem=Block01_LeftItem&"������Ϣ��"
end if
rs.close
set rs=nothing


'��������ȡ
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
%>
<!--common use end-->

<%'list_common use
'���⣬������ͷ���ʣ������ڵ�λ�ö�ȡ�滻
set rs1=server.createobject("adodb.recordset")
sql="select [id],[pid],[ppid],[name],[title],[content],[description],[folder],[keywords] from [category] where [id]="&ClassID&""
rs1.open(sql),cn,1,1
if not rs1.eof then
Class_Name=rs1("name")
Class_Intro=left(nohtml(rs1("content")),200)
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
'һ����������µĵ�ǰλ��
case 1
'----------------------
ClassSQL="cid"
ClassName1=rs1("name")
ClassFolder1=rs1("folder")
ClassID1=rs1("id")
MainClass_FolderName="/"&rs1("folder")
class_position=""
class_position=class_position&"<a href='/"&rs1("folder")&"/'>"&rs1("name")&"</a>"
'������������µĵ�ǰλ��
case 2
'--------------------
ClassSQL="pid"
set rs_1=server.createobject("adodb.recordset")
sql="select [id],[name],[folder] from [category] where [id]="&rs1("pid")&" and ppid=1"
rs_1.open(sql),cn,1,1
if not rs_1.eof then
ClassName1=rs_1("name")
ClassFolder1=rs_1("folder")
ClassID1=rs_1("id")
MainClass_FolderName="/"&rs_1("folder")
class_position=""
class_position=class_position&"<a href='/"&rs_1("folder")&"/'>"&rs_1("name")&"</a>"
class_position=class_position&" > <a href='/"&rs_1("folder")&"/"&rs1("folder")&"/'>"&rs1("name")&"</a>"
end if
rs_1.close
set rs_1=nothing

'������������µĵ�ǰλ��
case 3
'--------------------
ClassSQL="ppid"
set rs_1=server.createobject("adodb.recordset")
sql="select [id],[pid],[name],[folder] from [category] where [id]="&rs1("pid")&" and ppid=2"
rs_1.open(sql),cn,1,1
if not rs_1.eof then

set rs_1_1=server.createobject("adodb.recordset")
sql="select [id],[name],[folder] from [category] where [id]="&rs_1("pid")&" and ppid=1"
rs_1_1.open(sql),cn,1,1
if not rs_1_1.eof then
ClassName1=rs_1_1("name")
ClassFolder1=rs_1_1("folder")
ClassID1=rs_1_1("id")
MainClass_FolderName="/"&rs_1_1("folder")
SubClass_FolderName="/"&rs_1_1("folder")&"/"&rs_1("folder")
class_position=""
class_position=class_position&"<a href='/"&rs_1_1("folder")&"/'>"&rs_1_1("name")&"</a>"
end if

class_position=class_position&" > <a href='/"&rs_1_1("folder")&"/"&rs_1("folder")&"/'>"&rs_1("name")&"</a>"
end if

class_position=class_position&" > <a href='/"&rs_1_1("folder")&"/"&rs_1("folder")&"/"&rs1("folder")&"/'>"&rs1("name")&"</a>"

rs_1.close
set rs_1=nothing
rs_1_1.close
set rs_1_1=nothing
end select
end if 
rs1.close
set rs1=nothing

'�������ǰ��Ŀ�б�
Block_LeftClassList=""
set rsl=server.createobject("adodb.recordset")
sql="select [name],[folder],[id],[pid],[ppid] from [category] where pid="&ClassID1&" order by [order] "
rsl.open(sql),cn,1,1
Block_LeftClassList=Block_LeftClassList&"<ul id='suckertree1'>"
if not rsl.eof then
for i=1 to rsl.recordcount
if rsl("id")=ClassID then
Block_LeftClassList=Block_LeftClassList&"<li class='current'><A href='/"&ClassFolder1&"/"&rsl("Folder")&"'>"&rsl("name")&"</A> "
else
Block_LeftClassList=Block_LeftClassList&"<li><A href='/"&ClassFolder1&"/"&rsl("Folder")&"'>"&rsl("name")&"</A> "
end if
set rs=server.createobject("adodb.recordset")
sql="select  [name],[folder] from [category] where ppid=3 and pid="&rsl("id")&" "
rs.open(sql),cn,1,1
if not rs.eof then
Block_LeftClassList=Block_LeftClassList&"<ul>"
do while not rs.eof 
Block_LeftClassList=Block_LeftClassList&"<li><a href='/"&ClassFolder1&"/"&rsl("Folder")&"/"&rs("folder")&"/' >"&rs("name")&"</a></li> "
rs.movenext
loop
Block_LeftClassList=Block_LeftClassList&"</ul>"
end if
rs.close
set rs=nothing
Block_LeftClassList=Block_LeftClassList&"</li> "
rsl.movenext
next
else
Block_LeftClassList=Block_LeftClassList&""
end if
Block_LeftClassList=Block_LeftClassList&"</ul>"
rsl.close
set rsl=nothing

%>

<%
 '�����б��ȡ�滻
sql="select * from [article] where "&ClassSQL&"='"&ClassID&"' and view_yes=1  and ArticleType=2 order by [time] desc"
set rs1=server.createObject("ADODB.Recordset")
rs1.open sql,cn,1,1
%><%
if not rs1.eof then
rs1.pagesize=12
totalpage=rs1.pagecount

for j=1 to totalpage
sql="select * from [article] where "&ClassSQL&"='"&ClassID&"' and view_yes=1  and ArticleType=2 order by [time] desc"
set rs=server.createObject("ADODB.Recordset")
rs.open sql,cn,1,1
if not rs.eof then

rscount=rs.recordcount
whichpage=j 
rs.pagesize=12
totalpage=rs.pagecount
rs.absolutepage=whichpage
howmanyrecs=0
list_block=""
list_block=list_block&"<ul>"
do while not rs.eof and howmanyrecs<rs.pagesize
%><%
rs_url=""
if rs("url")<>"" then
rs_url=rs("url")
else
rs_url=ProductContent_FolderName&"/"&rs("file_path")
end if

list_block=list_block&"<div class='albumblock'><div class='inner'><a href='"&rs_url&"' target='_blank'  title='"&left(rs("title"),39)&"'><img src='/images/up_images/"&rs("image")&"' /><div class='albumtitle'>"&left(rs("title"),9)&"</div></a></div></div>"

%>
<%
rs.movenext
howmanyrecs=howmanyrecs+1
loop
list_block=list_block&"</ul>"
else
list_block=list_block&"�������ݣ�"
end if
rs.close
set rs=nothing
%>
<%
'��ҳ����
list_page=""
list_page=list_page&"<div class='t_page ColorLink'>"
list_page=list_page&"������"&rscount&"��&nbsp;&nbsp;��ǰҳ����<span class='FontRed'>" & j & "</span>/" & totalpage &""
list_page=list_page&"<a href=index.html"&">" & "��ҳ" & "</a>"
if whichpage=1 then
list_page=list_page&"&nbsp;&nbsp;��һҳ&nbsp;&nbsp;"
else
if j=2  then
list_page=list_page&"<a href=index.html"&">" & "��һҳ" & "</a>"
else
list_page=list_page&"<a href=list_"&j-1&".html"&">" & "��һҳ" & "</a>"
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
list_page=list_page&"&nbsp;&nbsp;��һҳ&nbsp;&nbsp;"
else
list_page=list_page&"<a href=list_"&j+1&".html"&">" & "��һҳ" & "</a>"
end if
if totalpage=1 then
list_page=list_page&"<a href=index.html"&">" & "βҳ" & "</a></div>"
else
list_page=list_page&"<a href=list_"&totalpage&".html"&">" & "βҳ" & "</a></div>"
end if
%>

<%
 '��ȡģ������
Set fso=Server.CreateObject("Scripting.FileSystemObject") 
Set htmlwrite=fso.OpenTextFile(Server.MapPath("/templates/"&web_theme&"/"&Model_FileName)) 
replace_code=htmlwrite.ReadAll() 
htmlwrite.close 
%>

<%'ѭ���б��滻����
replace_code=replace(replace_code,"$list_block$",list_block)   
replace_code=replace(replace_code,"$list_page$",list_page)   


%>
<%'ѭ�������滻����
replace_code=replace(replace_code,"$web_name$",web_name)
replace_code=replace(replace_code,"$web_url$",web_url)
replace_code=replace(replace_code,"$web_image$",web_image)
replace_code=replace(replace_code,"$web_title$",web_title)
replace_code=replace(replace_code,"$web_copyright$",web_copyright)
replace_code=replace(replace_code,"$web_theme$",web_theme)
replace_code=replace(replace_code,"$web_consult$",web_consult)
replace_code=replace(replace_code,"$web_TopHTML$",web_TopHTML)
replace_code=replace(replace_code,"$web_BottomHTML$",web_BottomHTML)
replace_code=replace(replace_code,"$web_link$",web_link)
replace_code=replace(replace_code,"$web_tel$",web_tel)
replace_code=replace(replace_code,"$search_FolderName$",search_FolderName)

replace_code=replace(replace_code,"$Block_LeftClassList$",Block_LeftClassList)
replace_code=replace(replace_code,"$ClassName1$",ClassName1)
replace_code=replace(replace_code,"$ClassFolder1$",ClassFolder1)

replace_code=replace(replace_code,"$Class_Position$",Class_Position)
replace_code=replace(replace_code,"$Class_Title$",Class_Title)
replace_code=replace(replace_code,"$Class_Name$",Class_Name)
replace_code=replace(replace_code,"$Class_Intro$",Class_Intro)
replace_code=replace(replace_code,"$Class_Description$",Class_Description)
replace_code=replace(replace_code,"$Class_Keywords$",Class_Keywords)

replace_code=replace(replace_code,"$web_TopMenu$",web_TopMenu)
replace_code=replace(replace_code,"$web_BottomMenu$",web_BottomMenu)
replace_code=replace(replace_code,"$Block01_LeftItem$",Block01_LeftItem)
replace_code=replace(replace_code,"$Block02_LeftItem$",Block02_LeftItem)
replace_code=replace(replace_code,"$Inner_BannerTop$",Inner_BannerTop)

%>

<% '�ж�ģ���ļ����Ƿ���ڣ����򴴽�
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
If Fso.FolderExists(Server.MapPath(Model_FolderName&"/"&Class_FolderName))=false Then
NewFolderDir=Model_FolderName&"/"&Class_FolderName
call CreateFolderB(NewFolderDir)
end if
filepath=Model_FolderName&"/"&Class_FolderName&"/list_"&j&".html"
filepath_index=Model_FolderName&"/"&Class_FolderName&"/index.html"	
	
	case 2
Set Fso=Server.CreateObject("Scripting.FileSystemObject") 
If Fso.FolderExists(Server.MapPath(Model_FolderName&MainClass_FolderName))=false Then
NewFolderDir=Model_FolderName&MainClass_FolderName
call CreateFolderB(NewFolderDir)
end if

Set Fso=Server.CreateObject("Scripting.FileSystemObject") 
If Fso.FolderExists(Server.MapPath(Model_FolderName&MainClass_FolderName&"/"&Class_FolderName))=false Then
NewFolderDir=Model_FolderName&MainClass_FolderName&"/"&Class_FolderName
call CreateFolderB(NewFolderDir)
end if
filepath=Model_FolderName&MainClass_FolderName&"/"&Class_FolderName&"/list_"&j&".html"
filepath_index=Model_FolderName&MainClass_FolderName&"/"&Class_FolderName&"/index.html"
	case 3
Set Fso=Server.CreateObject("Scripting.FileSystemObject") 
If Fso.FolderExists(Server.MapPath(Model_FolderName&MainClass_FolderName))=false Then
NewFolderDir=Model_FolderName&MainClass_FolderName
call CreateFolderB(NewFolderDir)
end if

Set Fso=Server.CreateObject("Scripting.FileSystemObject") 
If Fso.FolderExists(Server.MapPath(Model_FolderName&SubClass_FolderName))=false Then
NewFolderDir=Model_FolderName&SubClass_FolderName
call CreateFolderB(NewFolderDir)
end if

Set Fso=Server.CreateObject("Scripting.FileSystemObject") 
If Fso.FolderExists(Server.MapPath(Model_FolderName&SubClass_FolderName&"/"&Class_FolderName))=false Then
NewFolderDir=Model_FolderName&SubClass_FolderName&"/"&Class_FolderName
call CreateFolderB(NewFolderDir)
end if
filepath=Model_FolderName&SubClass_FolderName&"/"&Class_FolderName&"/list_"&j&".html" 
filepath_index=Model_FolderName&SubClass_FolderName&"/"&Class_FolderName&"/index.html" 
end select
%>

<%
if j>1 then
Set fso=Server.CreateObject("Scripting.FileSystemObject")
Set f=fso.CreateTextFile(Server.MapPath(filepath),true)
f.WriteLine replace_code
f.close
end if

if j=1 then
Set f=fso.CreateTextFile(Server.MapPath(filepath_index),true)
f.WriteLine replace_code
f.close
end if
%>
<%
next
else
j=1
%>
<%
 '��ȡģ������
Set fso=Server.CreateObject("Scripting.FileSystemObject") 
Set htmlwrite=fso.OpenTextFile(Server.MapPath("/templates/"&web_theme&"/"&Model_FileName)) 
replace_code=htmlwrite.ReadAll() 
htmlwrite.close 
%>

<%'ѭ���б��滻����
replace_code=replace(replace_code,"$list_block$","������Ϣ��")   
replace_code=replace(replace_code,"$list_page$","")   
%>
<%'ѭ�������滻����
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

replace_code=replace(replace_code,"$Class_Position$",Class_Position)
replace_code=replace(replace_code,"$Class_Title$",Class_Title)
replace_code=replace(replace_code,"$Class_Name$",Class_Name)
replace_code=replace(replace_code,"$Class_Intro$",Class_Intro)
replace_code=replace(replace_code,"$Class_Description$",Class_Description)
replace_code=replace(replace_code,"$Class_Keywords$",Class_Keywords)
replace_code=replace(replace_code,"$Block_LeftClassList$",Block_LeftClassList)
replace_code=replace(replace_code,"$ClassName1$",ClassName1)
replace_code=replace(replace_code,"$ClassFolder1$",ClassFolder1)

replace_code=replace(replace_code,"$web_TopMenu$",web_TopMenu)
replace_code=replace(replace_code,"$web_BottomMenu$",web_BottomMenu)
replace_code=replace(replace_code,"$Block01_LeftItem$",Block01_LeftItem)
replace_code=replace(replace_code,"$Block02_LeftItem$",Block02_LeftItem)
replace_code=replace(replace_code,"$Inner_BannerTop$",Inner_BannerTop)
%>

<% '�ж�ģ���ļ����Ƿ���ڣ����򴴽�
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
If Fso.FolderExists(Server.MapPath(Model_FolderName&"/"&Class_FolderName))=false Then
NewFolderDir=Model_FolderName&"/"&Class_FolderName
call CreateFolderB(NewFolderDir)
end if
filepath_index=Model_FolderName&"/"&Class_FolderName&"/index.html"	
	case 2
Set Fso=Server.CreateObject("Scripting.FileSystemObject") 
If Fso.FolderExists(Server.MapPath(Model_FolderName&MainClass_FolderName))=false Then
NewFolderDir=Model_FolderName&MainClass_FolderName
call CreateFolderB(NewFolderDir)
end if

Set Fso=Server.CreateObject("Scripting.FileSystemObject") 
If Fso.FolderExists(Server.MapPath(Model_FolderName&MainClass_FolderName&"/"&Class_FolderName))=false Then
NewFolderDir=Model_FolderName&MainClass_FolderName&"/"&Class_FolderName
call CreateFolderB(NewFolderDir)
end if
filepath_index=Model_FolderName&MainClass_FolderName&"/"&Class_FolderName&"/index.html"
	case 3
Set Fso=Server.CreateObject("Scripting.FileSystemObject") 
If Fso.FolderExists(Server.MapPath(Model_FolderName&MainClass_FolderName))=false Then
NewFolderDir=Model_FolderName&MainClass_FolderName
call CreateFolderB(NewFolderDir)
end if

Set Fso=Server.CreateObject("Scripting.FileSystemObject") 
If Fso.FolderExists(Server.MapPath(Model_FolderName&SubClass_FolderName))=false Then
NewFolderDir=Model_FolderName&SubClass_FolderName
call CreateFolderB(NewFolderDir)
end if

Set Fso=Server.CreateObject("Scripting.FileSystemObject") 
If Fso.FolderExists(Server.MapPath(Model_FolderName&SubClass_FolderName&"/"&Class_FolderName))=false Then
NewFolderDir=Model_FolderName&SubClass_FolderName&"/"&Class_FolderName
call CreateFolderB(NewFolderDir)
end if
filepath_index=Model_FolderName&SubClass_FolderName&"/"&Class_FolderName&"/index.html" 

end select
%>
<%
Set f=fso.CreateTextFile(Server.MapPath(filepath_index),true)
f.WriteLine replace_code
f.close
%>

<%end if
rs1.close
set rs1=nothing%>

<%end function%>
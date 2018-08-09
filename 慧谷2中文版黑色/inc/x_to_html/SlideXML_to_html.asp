<!-- #include file="../access.asp" -->
<!-- #include file="../html_clear.asp" -->
<%'容错处理
function SlideXML_to_html()
On Error Resume Next
%>

<!--common use start-->
<!--common use end-->

<%  
'首页顶部幻灯广告读取替换
set rs=server.createobject("adodb.recordset")
sql="select top 6 name,url,image from web_ads  where [position]=30 and view_yes=1 order by [order]"
rs.open(sql),cn,1,1
if not rs.eof then
for i=1 to rs.recordcount
'web_TopIMGAD=web_TopIMGAD&"box.add({""url"":""images/up_images/"&rs("image")&""",""href"":"""&rs("url")&""",""title"":"""&rs("name")&"""}); "
web_TopIMGAD=web_TopIMGAD&"<menu url="""&rs("url")&""" frame=""_parent"" imageUrl=""../images/up_images/"&rs("image")&"""/>"
rs.movenext
next
else
web_TopIMGAD=web_TopIMGAD&"无数据"
end if
rs.close
set rs=nothing
 %>

<% '读取模板内容
Set fso=Server.CreateObject("Scripting.FileSystemObject") 
Set htmlwrite=fso.OpenTextFile(Server.MapPath("/Templates/common/template.xml")) 
replace_code=htmlwrite.ReadAll() 
htmlwrite.close 
%>

<%'替换内容
replace_code=replace(replace_code,"$web_TopIMGAD$",web_TopIMGAD)
   %>
 <% 
filepath="/xml/images.xml"
 Set fso=Server.CreateObject("Scripting.FileSystemObject")
Set f=fso.CreateTextFile(Server.MapPath(filepath),true,true)
f.WriteLine replace_code
f.close
%>
<%end function%>
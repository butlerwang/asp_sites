<!-- #include file="../access.asp" -->
<!-- #include file="../html_clear.asp" -->
<%'�ݴ���
function SlideXML_to_html()
On Error Resume Next
%>

<!--common use start-->
<!--common use end-->

<%  
'��ҳ�����õƹ���ȡ�滻
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
web_TopIMGAD=web_TopIMGAD&"������"
end if
rs.close
set rs=nothing
 %>

<% '��ȡģ������
Set fso=Server.CreateObject("Scripting.FileSystemObject") 
Set htmlwrite=fso.OpenTextFile(Server.MapPath("/Templates/common/template.xml")) 
replace_code=htmlwrite.ReadAll() 
htmlwrite.close 
%>

<%'�滻����
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
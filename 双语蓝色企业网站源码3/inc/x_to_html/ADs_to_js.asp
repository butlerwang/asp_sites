<!-- #include file="../access.asp" -->

<%'�ݴ���
function ADs_to_js(l_id)
On Error Resume Next
%>
<%
'ģ�����ݶ�ȡ�滻
set rs=server.createobject("adodb.recordset")
sql="select * from web_ads where id="&l_id
rs.open(sql),cn,1,1
if not rs.eof  then

select case rs("ADtype")
'������
case 1
ADs_Content=ADs_Content&"<a href='"&rs("url")&"' target='_blank' title='"&rs("name")&"'>"&rs("name")&"</a>"
'ͼƬ���
case 2
ADs_Content=ADs_Content&"<a href='"&rs("url")&"' target='_blank' title='"&rs("name")&"'><img src='/images/up_images/"&rs("image")&"'  width='"&rs("ADWidth")&"' height='"&rs("ADHeight")&"' ></a>"
'Flash���
case 3
ADs_Content=ADs_Content&"<embed src='"&rs("FlashUrl")&"' width='"&rs("ADWidth")&"' height='"&rs("ADHeight")&"'></embed>"
'case 4 ������˴����棬һ����JS���룬����Ҫ����JS�ļ���
end select

end if
rs.close
set rs=nothing
%>
<% '��ȡģ������
Set fso=Server.CreateObject("Scripting.FileSystemObject") 
Set htmlwrite=fso.OpenTextFile(Server.MapPath("/templates/common/template.js")) 
replace_code=htmlwrite.ReadAll() 
htmlwrite.close 
%>
<%
replace_code=replace(replace_code,"$ADs_Content$",ADs_Content)

%>
<% '�ж�ADs�ļ����Ƿ���ڣ����򴴽�
Set Fso=Server.CreateObject("Scripting.FileSystemObject") 
If Fso.FolderExists(Server.MapPath("/ADs"))=false Then
NewFolderDir="/ADs"
call CreateFolderB(NewFolderDir)
end if
%>

<%'����HTML�ļ���,ָ���ļ�·��
filepath="/ADs/"&l_id&".js"
%>

<% '���ɾ�̬�ļ�
Set fout = fso.CreateTextFile(Server.MapPath(filepath))
fout.WriteLine replace_code
fout.close
fso.close
set fso=nothing
end function
%>
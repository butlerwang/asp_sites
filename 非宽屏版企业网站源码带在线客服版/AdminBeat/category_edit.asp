<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="inc/pingyin.asp" -->
<!-- #include file="../inc/x_to_html/index_to_html.asp" -->
<!-- #include file="../inc/x_to_html/Blank_Content_to_html.asp" -->

<% '��ȡ����
page=request.querystring("page")
act=request.querystring("act")
keywords=request.querystring("keywords")
article_id=cint(request.querystring("id"))

id1=cint(request.querystring("id"))
pid_name=request.querystring("pid_name")
pid_name2=request.querystring("pid_name2")
ppid=cint(request.querystring("ppid"))

act1=Request("act1")
If act1="save" Then 
id2=trim(request.form("id2"))
c_name=trim(request.form("c_name"))
c_title=trim(request.form("c_title"))
c_folder=trim(request.form("c_folder"))
c_image=trim(request.form("web_image"))
c_keywords=trim(request.form("c_keywords"))
c_description=trim(request.form("c_description"))
c_content=trim(request.form("a_content"))
c_index_push=trim(request.form("c_index_push"))
'c_Html_Yes=trim(request.form("c_Html_Yes"))
c_ClassType=trim(request.form("ClassType"))
c_order=trim(request.form("order"))

c_time=now()
%>

<% '����ת����ƴ��
if c_folder="" then
c_folder=trans_letters(c_name)
end if
%>
<%
'����ļ������Ƿ���ϵͳ�ļ������ظ�
			set rs_f=server.createobject("adodb.recordset")
			sql="select [name] from web_SystemFolder where [name]='"&c_folder&"'"
			rs_f.open(sql),cn,1,1
if not rs_f.eof then
response.Write "<script language='javascript'>alert('��Ŀ�ļ���������ϵͳ�ļ������ظ���������������');history.go(-1);</script>"
else
%>
<% 
set rs_1=server.createobject("adodb.recordset")
sql="select [id] from category where ( [folder]='"&c_folder&"' and [folder]<>'' ) and [id]<>"&id2
rs_1.open(sql),cn,1,3
if not rs_1.eof then
response.Write "<script language='javascript'>alert('���ļ��������Ѿ����ڣ�������������');history.go(-1);</script>"
else
rs_1.close
set rs_1=nothing
'��ӵ����ݿ�
set rs=server.createobject("adodb.recordset")
sql="select * from category where id="&id2&""
rs.open(sql),cn,1,3
c_folderold=rs("folder")
rs("name")=c_name
rs("title")=c_title
rs("folder")=c_folder
rs("image")=c_image
rs("keywords")=c_keywords
rs("description")=c_description
rs("content")=c_content
'rs("Html_Yes")=c_Html_Yes
'rs("index_push")=c_index_push
rs("ClassType")=c_ClassType
if c_order<>"" then
rs("order")=cint(c_order)
end if
rs("time")=c_time
rs.update
rs.close
set rs=nothing
%>

<% '������Ŀ�ļ���
set rs=server.createobject("adodb.recordset")
sql="select [folder] from [category] where [name]='"&pid_name&"'"
rs.open(sql),cn,1,1
if not rs.eof and not rs.bof then
folder_0=rs("folder")
Set Fso=Server.CreateObject("Scripting.FileSystemObject") 
If Fso.FolderExists(Server.MapPath("/"&rs("folder")))=false Then
NewFolderDir="/"&rs("folder")
call CreateFolderB(NewFolderDir)
end if


set rs2=server.createobject("adodb.recordset")
sql="select [folder] from [category] where [name]='"&pid_name2&"'"
rs2.open(sql),cn,1,1
if not rs2.eof and not rs2.bof then
folder_2=rs2("folder")
Set Fso=Server.CreateObject("Scripting.FileSystemObject") 
If Fso.FolderExists(Server.MapPath("/"&rs("folder")&"/"&rs2("folder")))=false Then
NewFolderDir="/"&rs("folder")&"/"&rs2("folder")
call CreateFolderB(NewFolderDir)
end if

rs2.close
set rs2=nothing
end if

rs.close
set rs=nothing
end if

if ppid=1 then
NewFolderDir1="/"&c_folder
OldFolderDir="/"&c_folderold
end if

if ppid=2 then
NewFolderDir1="/"&folder_0&"/"&c_folder
OldFolderDir="/"&folder_0&"/"&c_folderold
end if

if ppid=3 then
NewFolderDir1="/"&folder_0&"/"&folder_2&"/"&c_folder
OldFolderDir="/"&folder_0&"/"&folder_2&"/"&c_folderold
end if

'���ԭ�ļ����Ƿ���ڣ����򴴽�
Set Fso=Server.CreateObject("Scripting.FileSystemObject") 
If Fso.FolderExists(Server.MapPath(OldFolderDir))=false Then
NewFolderDir=OldFolderDir
call CreateFolderB(NewFolderDir)
end if
'������ļ����Ƿ���ԭ�ļ��в�ͬ�����������
if c_folder<>c_folderold  then
NewFolderDir=NewFolderDir1
call renamefolder(OldFolderDir,NewFolderDir) 
end if
%>

<%'������Ŀ��ҳ
if c_ClassType=5  then
ClassID=id2
call Blank_Content_to_html(ClassID)
 end if
 %>
<%
response.Write "<script language='javascript'>alert('�޸ĳɹ���');location.href='category_list.asp?page="&page&"&act="&act&"&keywords="&keywords&"';</script>"
end if
end if
rs_f.close
set rs_f=nothing

end if 

 %>
	<script charset="utf-8" src="Keditor/kindeditor.js"></script>
	<script charset="utf-8" src="Keditor/lang/zh_CN.js"></script>
	<script charset="utf-8" src="Keditor/editor.js"></script> 
	<%
Call header()

%>
<%
set rs2=server.createobject("adodb.recordset")
sql="select * from category where id="&id1&""
rs2.open(sql),cn,1,3
if not rs2.eof and not rs2.bof then
%>
  <form id="form1" name="form1" method="post" action="?act1=save&page=<%=page%>&act=<%=act%>&keywords=<%=keywords%>&pid_name=<%=pid_name%>&pid_name2=<%=pid_name2%>&ppid=<%=ppid%>">
         <script language='javascript'>
function checksignup1() {
if ( document.form1.c_name.value == '' ) {
window.alert('��������Ŀ����^_^');
document.form1.c_name.focus();
return false;}

return true;}
</script>
	<table cellpadding='3' cellspacing='1' border='0' class='tableBorder' align=center>
	<tr>
	  <th class='tableHeaderText' colspan=2 height=25>��Ŀ����</th>
	<tr>
	<td width="15%" height=23 class='forumRow'>��Ŀ���� (����)</td>
	<td class='forumRow'><input name='c_name' type='text' id='c_name' value="<%=rs2("name")%>" size='40'>
	<input name='id2' type='hidden' id='id2' size='40' value="<%=id1%>">
	
	  &nbsp;<span style="color: #FF0000"><%
	  if ppid=2 then
response.write "��ǰΪ������Ŀ:&nbsp;"&pid_name&"&nbsp;>"
elseif ppid=3 then
response.write "��ǰΪ������Ŀ:&nbsp;"&pid_name&"&nbsp;>&nbsp;"&pid_name2&"&nbsp;>"
else
response.write "��ǰΪһ����Ŀ"
end if%></span></td>
	</tr>
	<tr>
	  <td class='forumRow' height=23>��Ŀ���� (ѡ��) </td>
	  <td class='forumRow'>
	    <input name='c_title' type='text' id='c_title' size='50' maxlength="200" value="<%=rs2("title")%>"/>
	 </td>
	  </tr>
	<tr>
	<td class='forumRowHighLight' height=23>��Ŀ�ļ�������</td>
    <td class='forumRowHighLight'><input type='text' id='c_folder' name='c_folder' value="<%=rs2("folder")%>" size='40'  >
      <span style="color: #FF0000">��ʹ��Ӣ������������Ϊ�ս��Զ�ʹ����Ŀ���Ƶ�ƴ������,������ַ�����Ч������ϵͳ�ļ������ظ���</span> </td>
	</tr>
	  <tr>
	    <td class='forumRow' height=23>��ĿͼƬ</td>
	    <td width="85%" class='forumRow'><table width="100%" border="0" cellspacing="0" cellpadding="0">
         <tr>
           <td width="22%"  class='forumRow'><input name="web_image" type="text" id="web_image" value="<%=rs2("image")%>" size="30"></td>
           <td width="78%"  class='forumRow'><iframe width="500" name="ad" frameborder=0 height=30 scrolling=no src=upload.asp></iframe></td>
         </tr>
       </table></td>
      </tr>

      <td class='forumRowHighLight' height=11>��Ŀ�ؼ���</td>
	      <td class='forumRowHighLight'><input type='text' id='v3' name='c_keywords' value="<%=rs2("keywords")%>" size='80'>
	  &nbsp;���ԣ�����</td>
	</tr><tr>
	  <td class='forumRow' height=11>��Ŀ����</td>
	  <td class='forumRow'><textarea name='c_description'  cols="100" rows="4" id="c_description" ><%=rs2("description")%></textarea></td>
	</tr>
	<tr>
	  <td class='forumRowHighLight' height=23>��Ŀ���</td>
	  <td class='forumRowHighLight'>  <textarea name="a_content" id="a_content" style=" width:100%; height:400px; visibility:hidden;"><%=rs2("content")%></textarea>
</td>
	</tr>
	<tr>
	  <td class='forumRowHighLight' height=23><span style="color: #FF0000">��Ŀ���� (��ѡ)</span></td>
	  <td class='forumRowHighLight'><label>
	    <input type="radio" name="ClassType" value="1" <%
		if rs2("ClassType")=1 then
		response.write "checked"
		end if%>>
      ����
      &nbsp;&nbsp;
      <input name="ClassType" type="radio" value="2" <%if rs2("ClassType")=2 then
		response.write "checked"
		end if%>>
      ��Ʒ
      &nbsp;&nbsp;
      <input name="ClassType" type="radio" value="3" <%if rs2("ClassType")=3 then
		response.write "checked"
		end if%>>
      ����      &nbsp;&nbsp;
      <input name="ClassType" type="radio" value="4" <%if rs2("ClassType")=4 then
		response.write "checked"
		end if%>>
      ��Ƹ	  
	&nbsp;&nbsp;		
      <input name="ClassType" type="radio" value="5" <%if rs2("ClassType")=5 then
		response.write "checked"
		end if%>>
      ��ҳ</label></td>
	  </tr>	
<tr>
	    <td class='forumRow' height=23>��Ŀ����</td>
	    <td class='forumRow'><span class="forumRow">
	      <input name='order' type='text' id='order' value="<%=rs2("order")%>" size='20' maxlength="5">
	    &nbsp;ֻ�������֣�����ԽС����Խ��ǰ��Ĭ��Ϊ100�������������</span></td>
      </tr>	
	<tr><td height="50" colspan=2  class='forumRow'><div align="center">
	  <INPUT type=submit value='�ύ' onClick='javascript:return checksignup1()' name=Submit>
	  </div></td></tr>
	</table>
</form>
<%
end if
rs2.close
set rs2=nothing
%>
<%
Call DbconnEnd()
 %>
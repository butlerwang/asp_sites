<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="inc/pingyin.asp" -->
<!-- #include file="../inc/x_to_html/index_to_html.asp" -->
<!-- #include file="../inc/x_to_html/Blank_Content_to_html.asp" -->

<% '��ȡ����
pid_name=request.querystring("pid_name")
pid_name2=request.querystring("pid_name2")
pid=cint(request.querystring("pid"))
ppid=cint(request.querystring("ppid"))

act=Request("act")
If act="save" Then 
pid=trim(request.form("pid"))
ppid=trim(request.form("ppid"))
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

<% '�ж��Ƿ������ظ�����ӵ����ݿ�

set rs=server.createobject("adodb.recordset")
sql="select * from category where [folder]='"&c_folder&"'"
rs.open(sql),cn,1,3
if not rs.eof then
response.Write "<script language='javascript'>alert('����Ŀ�ļ��������Ѿ����ڣ�������������');history.go(-1);</script>"
else
rs.addnew
if ppid=2 then
rs("pid")=pid
rs("ppid")=2
elseif ppid=3 then
rs("pid")=pid
rs("ppid")=3
else
rs("ppid")=1
end if
rs("name")=c_name
rs("title")=c_title
rs("folder")=c_folder
rs("image")=c_image
rs("keywords")=c_keywords
rs("description")=c_description
rs("content")=c_content
'rs("index_push")=c_index_push
'rs("Html_Yes")=c_Html_Yes
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
if c_folder<>""  then

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
NewFolderDir="/"&c_folder
end if

if ppid=2 then
NewFolderDir="/"&folder_0&"/"&c_folder
end if

if ppid=3 then
NewFolderDir="/"&folder_0&"/"&folder_2&"/"&c_folder
end if


 '�ж��ļ����Ƿ���ڣ����򴴽�
Set Fso=Server.CreateObject("Scripting.FileSystemObject") 
If Fso.FolderExists(Server.MapPath(NewFolderDir))=false Then
call CreateFolderB(NewFolderDir)
end if

end if
%>
<%'������Ŀ��ҳ
if c_ClassType=5  then
set rs2=server.createobject("adodb.recordset")
sql="select top 1 [id] from [category] where [name]='"&c_name&"' order by [time] desc"
rs2.open(sql),cn,1,1
if not rs2.eof  then
ClassID=rs2("id")
call Blank_Content_to_html(ClassID)
end if
rs2.close
set rs2=nothing
end if %>
<%
response.Write "<script language='javascript'>alert('��ӳɹ���');location.href='category_list.asp';</script>"
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

  <form id="form1" name="form1" method="post" action="?act=save&pid_name=<%=pid_name%>&pid_name2=<%=pid_name2%>">
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
	  <th class='tableHeaderText' colspan=2 height=25>�����Ŀ</th>
	<tr>
	  <td height=23 colspan="2" class='forumRow'><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td height="20" class='TipTitle'>&nbsp;�� ������ʾ</td>
        </tr>
        <tr>
          <td height="30" valign="top" class="TipWords"><p>1��һ������£�������Ҫ��д��Ŀ���Ƽ���,��Ŀ�ļ������Ʋ�����Զ���ƴ��������</p>
            <p>2����Ŀ�ļ��н����ڸ�Ŀ¼�½��д������뾡��ʹ����ϵͳ�ļ��в�ͬ������������</p>
            <p>3��ϵͳ�ļ�����<%
			set rs_f=server.createobject("adodb.recordset")
			sql="select [name] from web_SystemFolder"
			rs_f.open(sql),cn,1,1
			do while not rs_f.eof
			response.write "��"&rs_f("name")
			rs_f.movenext
			loop
			rs_f.close
			set rs_f=nothing
			%>��</p>
            </td>
        </tr>
        <tr>
          <td height="10">&nbsp;</td>
        </tr>
      </table></td>
	  </tr>
	<tr>
	<td width="15%" height=23 class='forumRowHighLight'>��Ŀ���� (����) </td>
	<td class='forumRowHighLight'><input name='c_name' type='text' id='c_name' size='40'>
	<input name='pid' type='hidden' id='pid' size='40' value="<%=pid%>">
	<input name='ppid' type='hidden' id='ppid' size='40' value="<%=ppid%>">
	  &nbsp;<span style="color: #FF0000"><%
	  if ppid=2 then
response.write "��ǰΪ������Ŀ:&nbsp;"&pid_name
elseif ppid=3 then
response.write "��ǰΪ������Ŀ:&nbsp;"&pid_name&"&nbsp;>&nbsp;"&pid_name2
else
response.write "��ǰΪһ����Ŀ"
end if%></span></td>
	</tr>
	<tr>
	  <td class='forumRow' height=23>��Ŀ���� (ѡ��) </td>
	  <td class='forumRow'>
	    <input name='c_title' type='text' id='c_title' size='50' maxlength="200" />
	 </td>
	  </tr>
	<tr>
	<td class='forumRowHighLight' height=23>��Ŀ�ļ������� (ѡ��)</td>
    <td class='forumRowHighLight'><input type='text' id='c_folder' name='c_folder' size='40'>
      &nbsp;<span style="color: #FF0000">��ʹ��Ӣ������������Ϊ�ս��Զ�ʹ����Ŀ���Ƶ�ƴ������,������ַ�����Ч��</span><span style="color: #FF0000">����ϵͳ�ļ������ظ���</span></td>
	</tr>
	  <tr>
	    <td class='forumRow' height=23>��ĿͼƬ</td>
	    <td width="85%" class='forumRow'><table width="100%" border="0" cellspacing="0" cellpadding="0">
         <tr>
           <td width="22%"  class='forumRow'><input name="web_image" type="text" id="web_image"  size="30"></td>
           <td width="78%"  class='forumRow'><iframe width="500" name="ad" frameborder=0 height=30 scrolling=no src=upload.asp></iframe></td>
         </tr>
       </table></td>
      </tr>

      <td class='forumRowHighLight' height=11>��Ŀ�ؼ���</td>
	      <td class='forumRowHighLight'><input type='text' id='v3' name='c_keywords' size='80'>
	  &nbsp;���ԣ�����</td>
	</tr><tr>
	  <td class='forumRow' height=11>��Ŀ����</td>
	  <td class='forumRow'><textarea name='c_description'  cols="100" rows="4" id="c_description" ></textarea></td>
	</tr>
	<tr>
	  <td class='forumRowHighLight' height=23>��Ŀ���</td>
	  <td class='forumRowHighLight'>  <textarea name="a_content" id="a_content" style=" width:100%; height:400px; visibility:hidden;"></textarea>
</td>
	</tr>
	<tr>
	  <td class='forumRow' height=23>��Ŀ����</td>
	  <td class='forumRow'><label>
	    <input name="ClassType" type="radio" value="1" checked="checked" >
	    ����
	    &nbsp;&nbsp;
        <input name="ClassType" type="radio" value="2" >
      ��Ʒ
      &nbsp;&nbsp;
      <input name="ClassType" type="radio" value="4" >
      ��Ƹ	  
&nbsp;&nbsp;
<input name="ClassType" type="radio" value="5" />      
��ҳ</label></td>
	  </tr>	
<tr>
	    <td class='forumRow' height=23>��Ŀ����</td>
	    <td class='forumRow'><span class="forumRow">
	      <input name='order' type='text' id='order' value="100" size='20' maxlength="5">
	    &nbsp;ֻ�������֣�����ԽС����Խ��ǰ��Ĭ��Ϊ100�������������</span></td>
      </tr><tr>	
	<tr><td height="50" colspan=2  class='forumRow'><div align="center">
	  <INPUT type=submit value='�ύ' onClick='javascript:return checksignup1()' name=Submit>
	  </div></td></tr>
	</table>
</form>

<%
Call DbconnEnd()
 %>
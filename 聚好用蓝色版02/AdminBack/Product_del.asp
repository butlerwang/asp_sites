<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="../inc/x_to_html/index_to_html.asp" -->
<!-- #include file="../inc/x_to_html/Case_List_to_html.asp" -->
	<%
Call header()
%>
<%
'��Ʒ�����ļ��л�ȡ
set rs_1=server.createobject("adodb.recordset")
sql="select FileName,FolderName from web_Models_type where [id]=6"
rs_1.open(sql),cn,1,1
if not rs_1.eof and rs_1("FolderName")<>"" then
ProductContent_FolderName="/"&rs_1("FolderName")
end if
rs_1.close
set rs_1=nothing
%>

	<table cellpadding='3' cellspacing='1' border='0' class='tableBorder' align=center>
	<tr>
	  <th width="100%" height=25 class='tableHeaderText'>ɾ����Ʒ</th>
	
	<tr><td height="400" valign="top"  class='forumRow'><br>
	    <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="25" bgcolor="#B1CFF8"><div align="center"></div></td>
          </tr>
          <tr>
            <td height="100">
			<%page=request.querystring("page")
			act=request.querystring("act")
			keywords=request.querystring("keywords")

juhaoyongListDelStringPid=""
juhaoyongListDelStringPpid=""

if request("action")="AllDel" then
Num=request.form("Selectitem").count 
if Num=0 then 
response.Write "<script language='javascript'>alert('��ѡ��Ҫɾ�������ݣ�');location.href='Product_list.asp?page="&page&"&act="&act&"&keywords="&keywords&"';</script>"
Response.End 
end if 
Selectitem=request.Form("Selectitem") 
article_id=split(Selectitem,",")
c=ubound(article_id)
for i=0 to c		
			set rs=server.createobject("adodb.recordset")
sql="select id,cid,pid,ppid,file_path from Article where id="&cint(article_id(i))&""
rs.open(sql),cn,1,3
FilePath=rs("file_path")
ClassID=rs("cid")

if trim(rs("pid"))<>"" and rs("pid")<>tempJuhaoyongPid then
juhaoyongListDelStringPid=juhaoyongListDelStringPid&rs("pid")&","
end if

if trim(rs("ppid"))<>"" and rs("ppid")<>tempJuhaoyongPpid then
juhaoyongListDelStringPpid=juhaoyongListDelStringPpid&rs("ppid")&","
end if

tempJuhaoyongPid=rs("pid")
tempJuhaoyongPpid=rs("ppid")

rs.delete
rs.close
set rs=nothing
'���ж��ļ��Ƿ���ڣ�����ɾ��
Set fso=Server.CreateObject("Scripting.FileSystemObject")
If fso.FileExists(Server.MapPath(ProductContent_FolderName&"/"&FilePath)) then
FilePath=ProductContent_FolderName&"/"&FilePath
call DelFile(FilePath)
end if
fso.close
set fso=nothing

next

else
			article_id=cint(request.querystring("id"))
			set rs=server.createobject("adodb.recordset")
sql="select id,cid,pid,ppid,file_path from article where id="&article_id&""
rs.open(sql),cn,1,3
FilePath=rs("file_path")
ClassID=rs("cid")
juhaoyongClassPid=rs("pid")
juhaoyongClassPpid=rs("ppid")
rs.delete
rs.close
set rs=nothing
'���ж��ļ��Ƿ���ڣ�����ɾ��
Set fso=Server.CreateObject("Scripting.FileSystemObject")
If fso.FileExists(Server.MapPath(ProductContent_FolderName&"/"&FilePath)) then
FilePath=ProductContent_FolderName&"/"&FilePath
call DelFile(FilePath)
end if
fso.close
set fso=nothing

end if

'����������ҳ���б�ҳ��ʼ

'����������ҳ��һ���б�ҳ
call index_to_html()
call Case_List_to_html(ClassID)

'����ɾ�����������ɶ���Ŀ¼�б�
if trim(juhaoyongListDelStringPid)<>"" then
jhyListDelArrayPid=split(juhaoyongListDelStringPid,",")
	for juhaoyong_ii=0 to ubound(jhyListDelArrayPid)
		if trim(jhyListDelArrayPid(juhaoyong_ii))<>"" then
		call Case_List_to_html(jhyListDelArrayPid(juhaoyong_ii))
		end if
	next	
end if

'����ɾ����������������Ŀ¼�б�
if trim(juhaoyongListDelStringPpid)<>"" then
jhyListDelArrayPpid=split(juhaoyongListDelStringPpid,",")
	for juhaoyong_ii=0 to ubound(jhyListDelArrayPpid)
		if trim(jhyListDelArrayPpid(juhaoyong_ii))<>"" then
		call Case_List_to_html(jhyListDelArrayPpid(juhaoyong_ii))
		end if
	next	
end if

'����ɾ�����ж϶���Ŀ¼id�Ƿ��ظ���գ������ɶ���Ŀ¼�б�
if trim(juhaoyongClassPid)<>"" then
call Case_List_to_html(juhaoyongClassPid)
end if

'����ɾ�����ж�����Ŀ¼id�Ƿ��ظ���գ�����������Ŀ¼�б�
if trim(juhaoyongClassPpid)<>"" then
call Case_List_to_html(juhaoyongClassPpid)
end if

'����������ҳ���б�ҳ����


juhaoyong_cid=request.QueryString("juhaoyong_cid")
juhaoyong_pid=request.QueryString("juhaoyong_pid")
juhaoyong_ppid=request.QueryString("juhaoyong_ppid")

if juhaoyong_ppid>0 then
response.Write "<script language='javascript'>alert('ɾ���ɹ���');location.href='Product_list.asp?ppid="&juhaoyong_ppid&"&page="&page&"&act="&act&"&keywords="&keywords&"';</script>"
elseif juhaoyong_pid>0 then
response.Write "<script language='javascript'>alert('ɾ���ɹ���');location.href='Product_list.asp?pid="&juhaoyong_pid&"&page="&page&"&act="&act&"&keywords="&keywords&"';</script>"
elseif juhaoyong_cid>0 then
response.Write "<script language='javascript'>alert('ɾ���ɹ���');location.href='Product_list.asp?cid="&juhaoyong_cid&"&page="&page&"&act="&act&"&keywords="&keywords&"';</script>"
end if

			%></td>
          </tr>
        </table>
	    </td>
	</tr>
	</table>


<%
Call DbconnEnd()
 %>
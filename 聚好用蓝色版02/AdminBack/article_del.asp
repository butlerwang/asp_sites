<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="../inc/x_to_html/index_to_html.asp" -->
<!-- #include file="../inc/x_to_html/Article_list_to_html.asp" -->
<!-- #include file="../inc/x_to_html/article_to_html.asp" -->
	<%
Call header()
%>
<%
'���������ļ��л�ȡ
set rs_1=server.createobject("adodb.recordset")
sql="select FileName,FolderName from web_Models_type where [id]=5"
rs_1.open(sql),cn,1,1
if not rs_1.eof and rs_1("FolderName")<>"" then
ArticleContent_FolderName="/"&rs_1("FolderName")
end if
rs_1.close
set rs_1=nothing
%>

	<table cellpadding='3' cellspacing='1' border='0' class='tableBorder' align=center>
	<tr>
	  <th width="100%" height=25 class='tableHeaderText'>ɾ������</th>
	
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

set rs=server.createobject("adodb.recordset")
Set fso=Server.CreateObject("Scripting.FileSystemObject")

if request("action")="AllDel" then
	Num=request.form("Selectitem").count 
	if Num=0 then 
	response.Write "<script language='javascript'>alert('��ѡ��Ҫɾ�������ݣ�');location.href='article_list.asp?page="&page&"&act="&act&"&keywords="&keywords&"';</script>"
	Response.End 
	end if 
	
	Selectitem=request.Form("Selectitem") 
	article_id=split(Selectitem,",")
	
	c=ubound(article_id)
	for i=0 to c
		if i=0 then daArticleId=article_id(i)
		if i=c then xiaoArticleId=article_id(i)
	sql="select id,cid,pid,ppid,file_path from article where id="&cint(article_id(i))&""
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
	
	'�ж��ļ��Ƿ���ڣ�����ɾ��
	If fso.FileExists(Server.MapPath(ArticleContent_FolderName&"/"&FilePath)) then
	FilePath=ArticleContent_FolderName&"/"&FilePath
	call DelFile(FilePath)
	end if
	fso.close
	
	next

else
	article_id=cint(request.querystring("id"))
	sql="select id,cid,pid,ppid,file_path from article where id="&article_id&""
	rs.open(sql),cn,1,3
	FilePath=rs("file_path")
	ClassID=rs("cid")
	juhaoyongClassPid=rs("pid")
	juhaoyongClassPpid=rs("ppid")
	rs.delete
	rs.close

	'���ж��ļ��Ƿ���ڣ�����ɾ��
	If fso.FileExists(Server.MapPath(ArticleContent_FolderName&"/"&FilePath)) then
	FilePath=ArticleContent_FolderName&"/"&FilePath
	call DelFile(FilePath)
	end if
	fso.close

end if
set rs=nothing
set fso=nothing

'����������ҳ���б�ҳ��ʼ

'����������ҳ��һ���б�ҳ
call index_to_html()
call Article_list_to_html(ClassID)

'����ɾ�����������ɶ���Ŀ¼�б�
if trim(juhaoyongListDelStringPid)<>"" then
jhyListDelArrayPid=split(juhaoyongListDelStringPid,",")
	for juhaoyong_ii=0 to ubound(jhyListDelArrayPid)
		if trim(jhyListDelArrayPid(juhaoyong_ii))<>"" then
		call Article_list_to_html(jhyListDelArrayPid(juhaoyong_ii))
		end if
	next	
end if

'����ɾ����������������Ŀ¼�б�
if trim(juhaoyongListDelStringPpid)<>"" then
jhyListDelArrayPpid=split(juhaoyongListDelStringPpid,",")
	for juhaoyong_ii=0 to ubound(jhyListDelArrayPpid)
		if trim(jhyListDelArrayPpid(juhaoyong_ii))<>"" then
		call Article_list_to_html(jhyListDelArrayPpid(juhaoyong_ii))
		end if
	next	
end if

'����ɾ�����ж϶���Ŀ¼id�Ƿ��ظ���գ������ɶ���Ŀ¼�б�
if trim(juhaoyongClassPid)<>"" then
call Article_list_to_html(juhaoyongClassPid)
end if

'����ɾ�����ж�����Ŀ¼id�Ƿ��ظ���գ�����������Ŀ¼�б�
if trim(juhaoyongClassPpid)<>"" then
call Article_list_to_html(juhaoyongClassPpid)
end if

'����������ҳ���б�ҳ����

juhaoyong_cid=request.QueryString("juhaoyong_cid")
juhaoyong_pid=request.QueryString("juhaoyong_pid")
juhaoyong_ppid=request.QueryString("juhaoyong_ppid")

'��������ǰ����м����¿�ʼ
'��ȡ������С����id
if request("action")="AllDel" then
daArticleId=juhaoyongGetQianOrHouArticleId(juhaoyong_cid,juhaoyong_pid,juhaoyong_ppid,daArticleId,"qian")
xiaoArticleId=juhaoyongGetQianOrHouArticleId(juhaoyong_cid,juhaoyong_pid,juhaoyong_ppid,xiaoArticleId,"hou")
else
daArticleId=juhaoyongGetQianOrHouArticleId(juhaoyong_cid,juhaoyong_pid,juhaoyong_ppid,article_id,"qian")
xiaoArticleId=juhaoyongGetQianOrHouArticleId(juhaoyong_cid,juhaoyong_pid,juhaoyong_ppid,article_id,"hou")
end if


'��������
sql="select id from [article] where cid='"&juhaoyong_cid&"' and pid='"&juhaoyong_pid&"' and ppid='"&juhaoyong_ppid&"' and [id]>="&xiaoArticleId&" and [id]<="&daArticleId&" and view_yes=1 and ArticleType=1 order by [id] desc"
'sql="select [id],[ArticleType] from [article]  where view_yes=1 order by [time] desc"
set rs_create=server.createobject("adodb.recordset")
rs_create.open(sql),cn,1,1
	do while not rs_create.eof 
	a_id=rs_create("id")
	call article_to_html(a_id)
	rs_create.movenext
	loop
rs_create.close
set rs_create=nothing

'��������ǰ����м����½���

response.Write "<script language='javascript'>alert('ɾ���ɹ���');location.href='article_list.asp?cid="&juhaoyong_cid&"&pid="&juhaoyong_pid&"&ppid="&juhaoyong_ppid&"&page="&page&"&act="&act&"&keywords="&keywords&"';</script>"

%>
			</td>
          </tr>
        </table>
	    </td>
	</tr>
	</table>


<%
Call DbconnEnd()
 %>
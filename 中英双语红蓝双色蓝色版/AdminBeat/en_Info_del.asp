<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="../inc/en_x_to_html/index_to_html.asp" -->
<!-- #include file="../inc/en_x_to_html/Recruit_list_to_html.asp" -->
	<%
Call header()
%>
<%
'��Ƹְλ����ҳ�ļ��л�ȡ
set rs_1=server.createobject("adodb.recordset")
sql="select FileName,FolderName from web_Models_type where [id]=44"
rs_1.open(sql),cn,1,1
if not rs_1.eof and rs_1("FolderName")<>"" then
CategoryContent_FolderName="/"&rs_1("FolderName")
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
			
if request("action")="AllDel" then
Num=request.form("Selectitem").count 
if Num=0 then 
response.Write "<script language='javascript'>alert('��ѡ��Ҫɾ�������ݣ�');location.href='en_info_list.asp?page="&page&"&act="&act&"&keywords="&keywords&"';</script>"
Response.End 
end if 
Selectitem=request.Form("Selectitem") 
article_id=split(Selectitem,",")
c=ubound(article_id)
for i=0 to c		
			set rs=server.createobject("adodb.recordset")
sql="select id,cid,file_path from en_web_info where id="&cint(article_id(i))&""
rs.open(sql),cn,1,3
FilePath=rs("file_path")
ClassID=rs("cid")
rs.delete
rs.close
set rs=nothing
'�ж��ļ��Ƿ���ڣ�����ɾ��
Set fso=Server.CreateObject("Scripting.FileSystemObject")
If fso.FileExists(Server.MapPath(CategoryContent_FolderName&"/"&FilePath)) then
FilePath=CategoryContent_FolderName&"/"&FilePath
call DelFile(FilePath)
end if
next

else
			article_id=cint(request.querystring("id"))
			set rs=server.createobject("adodb.recordset")
sql="select id,cid,file_path from en_web_info where id="&article_id&""
rs.open(sql),cn,1,3
FilePath=rs("file_path")
ClassID=rs("cid")
rs.delete
rs.close
set rs=nothing
'���ж��ļ��Ƿ���ڣ�����ɾ��
Set fso=Server.CreateObject("Scripting.FileSystemObject")
If fso.FileExists(Server.MapPath(CategoryContent_FolderName&"/"&FilePath)) then
FilePath=CategoryContent_FolderName&"/"&FilePath
call DelFile(FilePath)
end if

end if
call Recruit_list_to_html(ClassID)
response.Write "<script language='javascript'>alert('ɾ���ɹ���');location.href='en_info_list.asp?page="&page&"&act="&act&"&keywords="&keywords&"';</script>"
			%></td>
          </tr>
        </table>
	    </td>
	</tr>
	</table>


<%
Call DbconnEnd()
 %>
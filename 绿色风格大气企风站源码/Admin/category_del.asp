<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->

	<%
Call header()
%>
<%
'��Ŀ�ļ��л�ȡ
set rs_1=server.createobject("adodb.recordset")
sql="select FolderName from web_Models_type where [id]=3"
rs_1.open(sql),cn,1,1
if not rs_1.eof then
BlogClass_FolderName=rs_1("FolderName")
end if
rs_1.close
set rs_1=nothing%>

	<table cellpadding='3' cellspacing='1' border='0' class='tableBorder' align=center>
	<tr>
	  <th width="100%" height=25 class='tableHeaderText'>ɾ����Ŀ</th>
	
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
			article_id=cint(request.querystring("id"))

'�������Ŀ��������Ŀ��Ҫ��ɾ������Ŀ�µ��¼���Ŀ
	set rs=server.createobject("adodb.recordset")
sql="select [id] from [category] where pid="&article_id
rs.open(sql),cn,1,1
if not rs.eof then
response.Write "<script language='javascript'>alert('����Ŀ�´�������Ŀ������ɾ������Ŀ��');history.go(-1);</script>"
else
rs.close
set rs=nothing
	set rs=server.createobject("adodb.recordset")
sql="select [id],[folder],pid,ppid from [category] where id="&article_id
rs.open(sql),cn,1,3

if rs("ppid")=1 then
FolderPath="/"&rs("folder")
end if

if rs("ppid")=2 then
set rs2=server.createobject("adodb.recordset")
sql="select [folder] from [category] where id="&rs("pid")
rs2.open(sql),cn,1,1
if not rs2.eof then
FolderPath="/"&rs2("folder")&"/"&rs("folder")
end if
rs2.close 
set rs2=nothing
end if

if rs("ppid")=3 then
set rs3=server.createobject("adodb.recordset")
sql="select [folder],pid from [category] where id="&rs("pid")
rs3.open(sql),cn,1,1
if not rs3.eof then
set rs2=server.createobject("adodb.recordset")
sql="select [folder] from [category] where id="&rs3("pid")
rs2.open(sql),cn,1,1
if not rs2.eof then

FolderPath="/"&rs2("folder")&"/"&rs3("folder")&"/"&rs("folder")
end if
rs2.close 
set rs2=nothing
end if
rs3.close 
set rs3=nothing
end if

Set Fso=Server.CreateObject("Scripting.FileSystemObject") 
If Fso.FolderExists(Server.MapPath(FolderPath)) Then
call DelFolder(FolderPath)
end if


rs.delete
response.Write "<script language='javascript'>alert('ɾ���ɹ���');location.href='category_list.asp?page="&page&"&act="&act&"&keywords="&keywords&"';</script>"
rs.close
set rs=nothing
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
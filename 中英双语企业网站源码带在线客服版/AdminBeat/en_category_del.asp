﻿<%@ LANGUAGE=VBScript CodePage=65001%>
<% response.charset="utf-8" %>
<% session.codepage=65001 %>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->

	<%
Call header()
%>
<%
'栏目文件夹获取
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
	  <th width="100%" height=25 class='tableHeaderText'>删除栏目</th>
	
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

'如果该栏目下有子栏目需要先删除该栏目下的下级栏目
	set rs=server.createobject("adodb.recordset")
sql="select [id] from [en_category] where pid="&article_id
rs.open(sql),cn,1,1
if not rs.eof then
response.Write "<script language='javascript'>alert('该栏目下存在子栏目，请先删除子栏目！');history.go(-1);</script>"
else
rs.close
set rs=nothing
	set rs=server.createobject("adodb.recordset")
sql="select [id],[folder],pid,ppid from [en_category] where id="&article_id
rs.open(sql),cn,1,3

if rs("ppid")=1 then
FolderPath="/"&rs("folder")
end if

if rs("ppid")=2 then
set rs2=server.createobject("adodb.recordset")
sql="select [folder] from [en_category] where id="&rs("pid")
rs2.open(sql),cn,1,1
if not rs2.eof then
FolderPath="/"&rs2("folder")&"/"&rs("folder")
end if
rs2.close 
set rs2=nothing
end if

if rs("ppid")=3 then
set rs3=server.createobject("adodb.recordset")
sql="select [folder],pid from [en_category] where id="&rs("pid")
rs3.open(sql),cn,1,1
if not rs3.eof then
set rs2=server.createobject("adodb.recordset")
sql="select [folder] from [en_category] where id="&rs3("pid")
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
response.Write "<script language='javascript'>alert('删除成功！');location.href='en_category_list.asp?page="&page&"&act="&act&"&keywords="&keywords&"';</script>"
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
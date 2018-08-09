<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="../inc/x_to_html/index_to_html.asp" -->
<!-- #include file="../inc/x_to_html/article_to_html.asp" -->
<!-- #include file="../inc/x_to_html/Article_list_to_html.asp" -->
	<%
Call header()
%>
<%
'文章内容文件夹获取
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
	  <th width="100%" height=25 class='tableHeaderText'>审核文章</th>
	
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
			set rs_v=server.createobject("adodb.recordset")
sql="select id,view_yes,cid,file_path from article where id="&article_id&""
rs_v.open(sql),cn,1,3
FilePath=rs_v("file_path")
ClassID=rs_v("cid")
if rs_v("view_yes")=0 then
rs_v("view_yes")=1
a_id=rs_v("id")
call article_to_html(a_id)
else
rs_v("view_yes")=0
'先判断文件是否存在，否则删除
Set fso=Server.CreateObject("Scripting.FileSystemObject")
If fso.FileExists(Server.MapPath(ArticleContent_FolderName&"/"&FilePath)) then
FilePath=ArticleContent_FolderName&"/"&FilePath
call DelFile(FilePath)
end if
end if
rs_v.update
rs_v.close
set rs_v=nothing

call index_to_html()
call Article_list_to_html(ClassID)

juhaoyong_cid=request.QueryString("juhaoyong_cid")
juhaoyong_pid=request.QueryString("juhaoyong_pid")
juhaoyong_ppid=request.QueryString("juhaoyong_ppid")

'重新生成前后和中间文章开始
'获取最大和最小文章id
daArticleId=juhaoyongGetQianOrHouArticleId(juhaoyong_cid,juhaoyong_pid,juhaoyong_ppid,article_id,"qian")
xiaoArticleId=juhaoyongGetQianOrHouArticleId(juhaoyong_cid,juhaoyong_pid,juhaoyong_ppid,article_id,"hou")

'生成文章
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

'重新生成前后和中间文章结束

response.Write "<script language='javascript'>alert('修改成功！');location.href='article_list.asp?cid="&juhaoyong_cid&"&pid="&juhaoyong_pid&"&ppid="&juhaoyong_ppid&"&page="&page&"&act="&act&"&keywords="&keywords&"';</script>"

			%></td>
          </tr>
        </table>
	    </td>
	</tr>
	</table>


<%
Call DbconnEnd()
 %>
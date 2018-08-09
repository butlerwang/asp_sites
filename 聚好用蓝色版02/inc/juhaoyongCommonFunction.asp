<%
function juhaoyongArticleNextCodeHtml(articleId,ArticleContent_FolderName)
'On Error Resume Next

juhaoyongCid=""
juhaoyongPid=""
juhaoyongPPid=""

article_next=""
rs_url=""

'上一篇，下一篇读取替换
sql="select cid,pid,ppid from [article] where id="&articleId
set rs=server.createobject("adodb.recordset")
rs.open(sql),cn,1,1
if not rs.eof and not rs.bof then
juhaoyongCid=rs("cid")
juhaoyongPid=rs("pid")
juhaoyongPPid=rs("ppid")
end if
rs.close


'生成上一篇链接
sql="select top 1 cid,pid,ppid,[title],[file_path],url,[time] from [article] where cid='"&juhaoyongCid&"' and pid='"&juhaoyongPid&"' and ppid='"&juhaoyongPPid&"' and  [id]>"&articleId&"  and view_yes=1 and ArticleType=1 order by [id] asc"
rs.open(sql),cn,1,1
article_next=article_next&"<ul><li>上一篇："
if not rs.eof and not rs.bof then

		if rs("url")<>"" then
		rs_url=rs("url")
		else
		rs_url=ArticleContent_FolderName&"/"&rs("file_path")
		end if 
		
article_next=article_next&"<a href='"&rs_url&"' target='_blank' title='"&rs("title")&"'>"&left(rs("title"),30)&"</a> <span class='ListDate'>&nbsp;["&year(rs("time"))&"/"&month(rs("time"))&"/"&day(rs("time"))&"]</span>"
else
article_next=article_next&"没有啦"
end if
article_next=article_next&"</li>"
rs.close


'生成下一篇链接
sql="select top 1 cid,pid,ppid,[title],[file_path],url,[time] from [article] where cid='"&juhaoyongCid&"' and pid='"&juhaoyongPid&"' and ppid='"&juhaoyongPPid&"' and  [id]<"&articleId&"  and view_yes=1 and ArticleType=1 order by [id] desc"
rs.open(sql),cn,1,1
article_next=article_next&"<li>下一篇："
if not rs.eof and not rs.bof then

		if rs("url")<>"" then
		rs_url=rs("url")
		else
		rs_url=ArticleContent_FolderName&"/"&rs("file_path")
		end if 
		
article_next=article_next&"<a href='"&rs_url&"' target='_blank' title='"&rs("title")&"'>"&left(rs("title"),30)&"</a> <span class='ListDate'>&nbsp;["&year(rs("time"))&"/"&month(rs("time"))&"/"&day(rs("time"))&"]</span>"
else
article_next=article_next&"没有啦"
end if
article_next=article_next&"</li></ul>"
rs.close
set rs=nothing

juhaoyongArticleNextCodeHtml=article_next
end function
%>






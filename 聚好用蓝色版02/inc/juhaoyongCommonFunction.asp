<%
function juhaoyongArticleNextCodeHtml(articleId,ArticleContent_FolderName)
'On Error Resume Next

juhaoyongCid=""
juhaoyongPid=""
juhaoyongPPid=""

article_next=""
rs_url=""

'��һƪ����һƪ��ȡ�滻
sql="select cid,pid,ppid from [article] where id="&articleId
set rs=server.createobject("adodb.recordset")
rs.open(sql),cn,1,1
if not rs.eof and not rs.bof then
juhaoyongCid=rs("cid")
juhaoyongPid=rs("pid")
juhaoyongPPid=rs("ppid")
end if
rs.close


'������һƪ����
sql="select top 1 cid,pid,ppid,[title],[file_path],url,[time] from [article] where cid='"&juhaoyongCid&"' and pid='"&juhaoyongPid&"' and ppid='"&juhaoyongPPid&"' and  [id]>"&articleId&"  and view_yes=1 and ArticleType=1 order by [id] asc"
rs.open(sql),cn,1,1
article_next=article_next&"<ul><li>��һƪ��"
if not rs.eof and not rs.bof then

		if rs("url")<>"" then
		rs_url=rs("url")
		else
		rs_url=ArticleContent_FolderName&"/"&rs("file_path")
		end if 
		
article_next=article_next&"<a href='"&rs_url&"' target='_blank' title='"&rs("title")&"'>"&left(rs("title"),30)&"</a> <span class='ListDate'>&nbsp;["&year(rs("time"))&"/"&month(rs("time"))&"/"&day(rs("time"))&"]</span>"
else
article_next=article_next&"û����"
end if
article_next=article_next&"</li>"
rs.close


'������һƪ����
sql="select top 1 cid,pid,ppid,[title],[file_path],url,[time] from [article] where cid='"&juhaoyongCid&"' and pid='"&juhaoyongPid&"' and ppid='"&juhaoyongPPid&"' and  [id]<"&articleId&"  and view_yes=1 and ArticleType=1 order by [id] desc"
rs.open(sql),cn,1,1
article_next=article_next&"<li>��һƪ��"
if not rs.eof and not rs.bof then

		if rs("url")<>"" then
		rs_url=rs("url")
		else
		rs_url=ArticleContent_FolderName&"/"&rs("file_path")
		end if 
		
article_next=article_next&"<a href='"&rs_url&"' target='_blank' title='"&rs("title")&"'>"&left(rs("title"),30)&"</a> <span class='ListDate'>&nbsp;["&year(rs("time"))&"/"&month(rs("time"))&"/"&day(rs("time"))&"]</span>"
else
article_next=article_next&"û����"
end if
article_next=article_next&"</li></ul>"
rs.close
set rs=nothing

juhaoyongArticleNextCodeHtml=article_next
end function
%>






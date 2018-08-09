<!--#include file="../conn.asp"-->
<%

'==============================================================
'请根据你的需要自行修改以下代码
'本文件调用方式:<script src='/plus/total.asp'></script>
'=================================================================
Response.Write("document.writeln('文章数量：<font color=""red"">"&Conn.Execute("Select Count(id) From ks_Article")(0) & "</font> 篇<br />');")
Response.Write("document.writeln('软件数量：<font color=""red"">"&Conn.Execute("Select Count(id) From ks_download")(0) & "</font> 个<br />');") 
Response.Write("document.writeln('今日更新文章：<font color=""red"">" & conn.execute("select count(id) from ks_article where datediff(" & DataPart_D & ",adddate,getdate())<1 ")(0) & "</font> 篇<br />');")
Response.Write("document.writeln('今日更新软件：<font color=""red"">" & conn.execute("select count(id) from ks_download where datediff(" & DataPart_D & ",adddate,getdate())<1 ")(0) & "</font> 个<br />');")
Response.Write("document.writeln('文章总浏览次数：<font color=""red"">" & conn.execute("select sum(hits) from ks_article")(0) & "</font> 次<br />');")
Response.Write("document.writeln('下载总次数：<font color=""red"">" & conn.execute("select sum(hits) from ks_download")(0) & "</font> 次<br />');")
%>
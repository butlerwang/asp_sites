<%'------------------------------生成客服代码文件开始------------------------------%>

<%
function juhaoyongKefuHtmlCode()
'On Error Resume Next
juhaoyongKefu_html_code=""



'<!----------------------在线悬浮客服代码开始---------------------->
juhaoyongKefu_html_code=juhaoyongKefu_html_code&"<DIV id=juhaoyong_xuanfukefu>"
juhaoyongKefu_html_code=juhaoyongKefu_html_code&"<DIV id=juhaoyong_xuanfukefuBut onmouseover='ShowJhyXuanfu()'><table class=juhaoyong_xuanfukefuBut_table border=0 cellspacing=0 cellpadding=0><tr><td> </td></tr></table></DIV>"
juhaoyongKefu_html_code=juhaoyongKefu_html_code&"<DIV id=juhaoyong_xuanfukefuContent>"
juhaoyongKefu_html_code=juhaoyongKefu_html_code&"<table width=143 border=0 cellspacing=0 cellpadding=0>"
juhaoyongKefu_html_code=juhaoyongKefu_html_code&"<tr><td class=juhaoyong_xuanfukefuContent01 valign=top> </td></tr>"
juhaoyongKefu_html_code=juhaoyongKefu_html_code&"<tr><td class=juhaoyong_xuanfukefuContent02 align=center>"
juhaoyongKefu_html_code=juhaoyongKefu_html_code&"	<table border=0 cellspacing=0 cellpadding=0 align=center>"

'循环开始
set rs = server.CreateObject("adodb.recordset")
s_sql="select * from web_ads_position order by  id"
rs.Open (s_sql),cn,1,1
if not rs.eof then
do while not rs.eof

juhaoyongKefu_html_code=juhaoyongKefu_html_code&"    <tr><td class=jhykefu_box1>"&rs("name")&"</td></tr>"
juhaoyongKefu_html_code=juhaoyongKefu_html_code&"    <tr><td class=jhykefu_box2>"&rs("memo")&"</td></tr>"

rs.movenext
loop
'循环结束

juhaoyongKefu_html_code=juhaoyongKefu_html_code&"	</table>"
juhaoyongKefu_html_code=juhaoyongKefu_html_code&"</td></tr>	"
juhaoyongKefu_html_code=juhaoyongKefu_html_code&"<tr><td class=juhaoyong_xuanfukefuContent03 onclick=window.location.href='/Contact/'> </td></tr>"
juhaoyongKefu_html_code=juhaoyongKefu_html_code&"</table>"
juhaoyongKefu_html_code=juhaoyongKefu_html_code&"</DIV>"
juhaoyongKefu_html_code=juhaoyongKefu_html_code&"</DIV>"
'<!----------------------在线悬浮客服代码结束---------------------->

else
juhaoyongKefu_html_code=""
end if 
rs.close
set rs=nothing

juhaoyongKefuHtmlCode=juhaoyongKefu_html_code
end function
%>

<%
function juhaoyongKefu_to_html(jhyCodeString)
juhaoyongKefu_FolderName="/juhaoyong-kfimgs"
 '判断文件夹是否存在，否则创建
Set Fso=Server.CreateObject("Scripting.FileSystemObject") 
If Fso.FolderExists(Server.MapPath(juhaoyongKefu_FolderName))=false Then
call juhaoyongCreateFolder(juhaoyongKefu_FolderName)
end if


'声明HTML文件径和文件名
filepath=juhaoyongKefu_FolderName&"/juhaoyongKefu.html"


'生成在线客服静态文件
Set fout = fso.CreateTextFile(Server.MapPath(filepath))
fout.WriteLine jhyCodeString
fout.close
fso.close
set fso=nothing
end function
%>

<%
Function juhaoyongCreateFolder(NewFolderDir)
Set Fso=Server.CreateObject("Scripting.FileSystemObject") 
If Fso.FolderExists(Server.MapPath(NewFolderDir)) Then

else
Fso.CreateFolder(Server.MapPath(NewFolderDir))
end if
set fso=nothing
End Function
%>
<%'------------------------------生成客服代码文件结束------------------------------%>






<%'------------------------------���ɿͷ������ļ���ʼ------------------------------%>

<%
function juhaoyongKefuHtmlCode()
'On Error Resume Next
juhaoyongKefu_html_code=""



'<!----------------------���������ͷ����뿪ʼ---------------------->
juhaoyongKefu_html_code=juhaoyongKefu_html_code&"<DIV id=juhaoyong_xuanfukefu>"
juhaoyongKefu_html_code=juhaoyongKefu_html_code&"<DIV id=juhaoyong_xuanfukefuBut onmouseover='ShowJhyXuanfu()'><table class=juhaoyong_xuanfukefuBut_table border=0 cellspacing=0 cellpadding=0><tr><td> </td></tr></table></DIV>"
juhaoyongKefu_html_code=juhaoyongKefu_html_code&"<DIV id=juhaoyong_xuanfukefuContent>"
juhaoyongKefu_html_code=juhaoyongKefu_html_code&"<table width=143 border=0 cellspacing=0 cellpadding=0>"
juhaoyongKefu_html_code=juhaoyongKefu_html_code&"<tr><td class=juhaoyong_xuanfukefuContent01 valign=top> </td></tr>"
juhaoyongKefu_html_code=juhaoyongKefu_html_code&"<tr><td class=juhaoyong_xuanfukefuContent02 align=center>"
juhaoyongKefu_html_code=juhaoyongKefu_html_code&"	<table border=0 cellspacing=0 cellpadding=0 align=center>"

'ѭ����ʼ
set rs = server.CreateObject("adodb.recordset")
s_sql="select * from web_ads_position order by  id"
rs.Open (s_sql),cn,1,1
if not rs.eof then
do while not rs.eof

juhaoyongKefu_html_code=juhaoyongKefu_html_code&"    <tr><td class=jhykefu_box1>"&rs("name")&"</td></tr>"
juhaoyongKefu_html_code=juhaoyongKefu_html_code&"    <tr><td class=jhykefu_box2>"&rs("memo")&"</td></tr>"

rs.movenext
loop
'ѭ������

juhaoyongKefu_html_code=juhaoyongKefu_html_code&"	</table>"
juhaoyongKefu_html_code=juhaoyongKefu_html_code&"</td></tr>	"
juhaoyongKefu_html_code=juhaoyongKefu_html_code&"<tr><td class=juhaoyong_xuanfukefuContent03 onclick=window.location.href='/Contact/'> </td></tr>"
juhaoyongKefu_html_code=juhaoyongKefu_html_code&"</table>"
juhaoyongKefu_html_code=juhaoyongKefu_html_code&"</DIV>"
juhaoyongKefu_html_code=juhaoyongKefu_html_code&"</DIV>"
'<!----------------------���������ͷ��������---------------------->

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
 '�ж��ļ����Ƿ���ڣ����򴴽�
Set Fso=Server.CreateObject("Scripting.FileSystemObject") 
If Fso.FolderExists(Server.MapPath(juhaoyongKefu_FolderName))=false Then
call juhaoyongCreateFolder(juhaoyongKefu_FolderName)
end if


'����HTML�ļ������ļ���
filepath=juhaoyongKefu_FolderName&"/juhaoyongKefu.html"


'�������߿ͷ���̬�ļ�
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
<%'------------------------------���ɿͷ������ļ�����------------------------------%>






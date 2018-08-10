<%
if session("Urule")<>"a" then
response.redirect "error.asp?id=admin"
end if
%>
<!--#INCLUDE FILE="data.asp" -->
<!--#include file="html.asp"-->
<%
set rs=server.createobject("ADODB.recordset") 
rs.Open "SELECT * FROM soft Where ID is null",conn,1,3 
rs.addnew
 nowtime=now()
sj=cstr(year(nowtime))+"-"+right("0"+cstr(month(nowtime)),2)+"-"+right("0"+cstr(day(nowtime)),2)
	rs("name") =htmlencode2(request("softname"))
	rs("content") =htmlencode2(request("content"))
	rs("url") =request("url")
	rs("time")=sj
    rs("best")=request("best")
	rs.Update
	rs.close
	Set rs=nothing
	Conn.Close
	Set Conn=nothing
%>
<LINK href="oa.css" rel=stylesheet>
<script language=javascript>
opener.location=opener.location
</script>
<BODY>
        <TABLE border=1 bordercolorlight='000000' bordercolordark=#ffffff cellspacing=0 cellpadding=0 align=center>
	  <TR> 
            <TD>文件已经成功加上，是否继续添加……<BR>
<P><P><A HREF="addsoft.asp">继续添加</A>&nbsp;&nbsp;<A HREF="javascript:window.close()">关闭窗口</A></TD>
            
      </TR>
	  
      </table>

<%
set file=nothing
set upload=nothing  ''删除此对象
%>
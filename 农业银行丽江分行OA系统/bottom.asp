<!--#INCLUDE FILE="data.asp" -->
<html><head><title>bottom_main</title>
<script>
function OpenWindows(url,widthx,heighx)
{
  var 
 newwin=window.open(url,"_blank","toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,top=20,left=60,width=320,height=160");
 return false;
 
}
function OpenWindows1(url,widthx,heighx)
{
  var 
 newwin=window.open(url,"_blank","toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,top=20,left=60,width=500,height=300");
 return false;
 
}
</script>







<script language=javascript>

	<% '-----------------------------邮箱系统收信通告代码------------------------------------------------------------------------------------------------------------------------------------------------------------------

	dim strr
	set con = Server.CreateObject("ADODB.Connection") 
	ConnStr="DBQ=" & Server.Mappath("db/mails1.mdb") & ";DRIVER={Microsoft Access Driver (*.mdb)};"
	con.Open(ConnStr)
	'创建并打开 Recordset 对象。
	set Record = Server.CreateObject("ADODB.Recordset")
	Record.ActiveConnection = con
	Record.CursorType = 1
	Record.LockType =1
	Record.open("select * from recived"+Session("id")+" where iread='f'")
		if Session("num")< Record.recordcount then 
			''strr="var nwindow=window.showModalDialog('',null,'dialogwidth=250,dialogheight=150,center=1');"
			strr="var nwindow=window.open('',null,'width=250,height=150,top=300,left=450');"
			strr=strr+"nwindow.document.write('<body bgcolor=#aaaaaa ><font color=blue><b>"+Session("id")+"</b></font><font color=black style=fontsize=9px>，您好！<br>您的信箱有了新邮件!</font><br><form method=post action=mailbox.asp?mailbox=recived target=main> <input type=submit value=立即查看 onclick=window.close()><input type=button value=待会儿再看 onclick=window.close()></form></body>')"
			response.write(strr)
			Session("num")=Record.recordcount
		end if
	
	Record.Close
	set Record=nothing
	con.close
	set con=nothing'--------------------------邮箱系统收信通告代码-----------------------------------------------------------------------------------------------------------------------------------------------

	%>

</script>





<meta content="text/html; charset=gb2312" http-equiv=Content-Type>
<meta HTTP-EQUIV=REFRESH CONTENT=15;URL=bottom.asp>
<meta content="MSHTML 5.00.2614.3500" name=GENERATOR></head>
   <%
nowtime=now()
shijian=cstr(year(nowtime))+right("0"+cstr(month(nowtime)),2)+right("0"+cstr(day(nowtime)),2)+right("0"+cstr(hour(nowtime)),2)+right("0"+cstr(minute(nowtime)),2)
strSQL ="SELECT * FROM calendar where (userid="&session("Uid")&" and state=false and remindtime<'"&shijian&"') ORDER BY id DESC"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open strSQL,Conn,1,1
if not rs.eof then
    do while not (rs.eof or rs.bof)
%>
<script LANGUAGE="JavaScript">
<!--
OpenWindows1('checkcalendar.asp?id='+<%=rs("id")%>,'width=500','height=300')
//-->
</script>
 <%rs.movenext 
loop 
else%>

<body bgColor=#276db2 leftMargin=0 topMargin=0 marginheight="0" 
marginwidth="0">
<%
end if

    set rs=server.createobject("ADODB.recordset")
    rs.open "select top 5 * from chat where receive='"&session("Uid")&"' and zt=false order by id",conn,1,1
    if not rs.eof then
    %>

    <%
    do while not (rs.eof or rs.bof)
    %>
<script LANGUAGE="JavaScript">
<!--
OpenWindows('AlertMsg.asp?id='+<%=rs("id")%>,'width=320','height=160')
//-->
</script>

<%rs.movenext 
loop 
else%>


<%
end if
Conn.execute"delete from chat WHERE endtime<=date()-30"
  rs.close
  set rs=nothing
  conn.close
  set conn=nothing%>
</body></html>

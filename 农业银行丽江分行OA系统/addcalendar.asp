<!--#INCLUDE FILE="data.asp" --> 
<!--#INCLUDE FILE="check.asp" -->
 <!--#INCLUDE FILE="html.asp" --> 
 <%
 if request("title")<>"" then
set rs=server.createobject("ADODB.recordset") 
rs.Open "SELECT * FROM calendar Where ID is null",conn,1,3 
rs.addnew

rs("title")=htmlencode2(request("title"))
rs("content")=htmlencode2(request("content"))
rs("time")=request("year1")&request("month1")&request("day1")&request("hour1")&request("minute1")
rs("remindtime")=request("year2")&request("month2")&request("day2")&request("hour2")&request("minute2")
rs("userid")=session("Uid")
rs.update 
rs.close
set rs=nothing
%> <SCRIPT LANGUAGE="JavaScript">
opener.location=opener.location;window.close();
</SCRIPT> <%
else
%> <script>
function OpenWindows(url)
{
  var 
 newwin=window.open(url,"_blank","toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,top=50,left=120,width=600,height=400");
 return false;
 
}
function OpenSmallWindows(url)
{
  var 
 newwin=window.open(url,"_blank","toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,top=50,left=120,width=550,height=270");
 return false;
 
}
</script> <html> <head> </P><title>日程添加</title> <meta http-equiv="Content-Type" content="text/html; charset=gb2312"> 
<link rel="stylesheet" href="oa.css"> </head> <body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"> 
<center> <br> <form action="addcalendar.asp" method="post" name="calendar" onsubmit="return ChkSubmit();"> 
<table width="90%" border="0" cellpadding="1" cellspacing="1" bgcolor="#666666" > 
<tr><td colspan=2 bgcolor="#000000" class="heading"><font color="#FFFFFF"><b>添加日程</b></font></td></tr> 
<tr> <td bgcolor="#EFEFEF" >活动名称：</td><td bgcolor="#FFFFFF"><input name="title" size=50></td></tr> 
<tr> <td bgcolor="#EFEFEF">活动时间：</td><td bgcolor="#FFFFFF"> <select name="year1"> 
<%for i=year(now()) to year(now())+4%> <option value="<%=i%>"><%=i%></option> 
<%next%> </select> 年 <select name="month1"> <%for i=1 to 12
		    j="0"&i
		  %> <option value="<%=right(j,2)%>" <%if i=month(now()) then response.write" selected"%>><%=right(j,2)%></option> 
<%next%> </select> 月 <select name="day1"> <%for i=1 to 31
		    j="0"&i
		  %> <option value="<%=right(j,2)%>" <%if i=day(now()) then response.write" selected"%>><%=right(j,2)%></option> 
<%next%> </select> 日 <select name="hour1"> <%for i=1 to 24
		    j="0"&i
		  %> <option value="<%=right(j,2)%>" <%if i=hour(now()) then response.write" selected"%>><%=right(j,2)%></option> 
<%next%> </select> 点 <input name=minute1 size=2 value=00>分 </td></tr> <tr> <td bgcolor="#EFEFEF">提醒时间：</td><td bgcolor="#FFFFFF"> 
<select name="year2"> <%for i=year(now()) to year(now())+4%> <option value="<%=i%>"><%=i%></option> 
<%next%> </select> 年 <select name="month2"> <%for i=1 to 12
		    j="0"&i
		  %> <option value="<%=right(j,2)%>" <%if i=month(now()) then response.write" selected"%>><%=right(j,2)%></option> 
<%next%> </select> 月 <select name="day2"> <%for i=1 to 31
		    j="0"&i
		  %> <option value="<%=right(j,2)%>" <%if i=day(now()) then response.write" selected"%>><%=right(j,2)%></option> 
<%next%> </select> 日 <select name="hour2"> <%for i=1 to 24
		    j="0"&i
		  %> <option value="<%=right(j,2)%>" <%if i=hour(now()) then response.write" selected"%>><%=right(j,2)%></option> 
<%next%> </select> 点 <input name=minute2 size=2 value=00>分 </td></tr> <tr> <td valign=top bgcolor="#EFEFEF">活动内容：</td><td bgcolor="#FFFFFF"><TEXTAREA NAME="content" ROWS="6" COLS="50"></TEXTAREA></td></tr> 
<tr> <td bgcolor="#EFEFEF"></td><td bgcolor="#FFFFFF"><input value="提交" name="sub1" type="submit">&nbsp;<input value="重写" name="set1" type="reset"></td></tr> 
</table></form></center> 
</body>
</html>
<script language="Javascript">
	function ChkSubmit()
	{
		if(isNaN(document.calendar.minute1.value)||isNaN(document.calendar.minute2.value)|| (document.calendar.minute1.value)>"59" || (document.calendar.minute2.value)>"59" || (document.calendar.minute1.value)<"00" || (document.calendar.minute2.value)<"00")
		{
			alert("请正确输入时间！");
			return false;
		}
		if(document.calendar.title.value=="")
		{
			alert("标题不能为空！");
			return false;
		}
		
	}
</script>
<%end if%>
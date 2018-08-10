<!--#INCLUDE FILE="data.asp" -->
<!--#INCLUDE FILE="check.asp" -->
<!--#INCLUDE FILE="html.asp" -->
<%
	set rs=server.createobject("ADODB.recordset")
    rs.open "select * from calendar where id="&request("id")&" and userid="&session("Uid")&" order by id desc",conn,1,3
	if rs.eof then
	response.write "error"
	response.end
	else
      if request("title")<>"" then
	  rs("title")=htmlencode2(request("title"))
      rs("content")=htmlencode2(request("content"))
      rs("time")=request("year1")&request("month1")&request("day1")&request("hour1")&request("minute1")
      rs("remindtime")=request("year2")&request("month2")&request("day2")&request("hour2")&request("minute2")
      rs.update 
%>
<SCRIPT LANGUAGE="JavaScript">
opener.location=opener.location;window.close();
</SCRIPT>

<%
	  else

%>

<script>
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
</script>

<html>
<head>
<title>日程修改</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="oa.css">
<script language="JavaScript">
<!--
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}

function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}

function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}
//-->
</script>
</head>
<body bgcolor="#efefef" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('images/close_2.gif','images/modify_on.gif','images/delete_on.gif')">

 <form action="modcalendar.asp?id=<%=rs("id")%>" method="post" name=calendar onsubmit="return ChkSubmit();">
 <table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#666666" >
 	<tr bgcolor="#4E5960"> 
      <td colspan=2 class="heading" height=3></td>
    </tr>  
	<tr bgcolor="#4E5960"> 
      <td colspan=2 class="heading"><font color="#FFFFFF"><b>添加修改</b></font></td>
    </tr>
	</table>
  <table width="100%" border="1" cellspacing="0" cellpadding="1" bordercolordark="#FFFFFF" bordercolor="#000000">
    <tr>	
		    <input type="hidden" name=calid value="100">
			
      <td bgcolor="#bfbfbf">活动名称：</td>
      <td bgcolor="#efefef">
        <input name="title" size=50 value="<%=rs("title")%>">
      </td>
		</tr>
		<tr>
		  <td bgcolor="#EFEFEF">活动时间：</td><td bgcolor="#FFFFFF"> <select name="year1">
		  <%for i=year(now()) to year(now())+4%>
          <option value="<%=i%>" <%if right(rs("time"),4)=i then response.write" selected" end if%>><%=i%></option>
		  <%next%>
        </select>
        年 
        <select name="month1">
		  <%for i=1 to 12
		    j="0"&i
		  %>
          <option value="<%=right(j,2)%>" <%if j=mid(rs("time"),5,2) then response.write" selected"%>><%=right(j,2)%></option>
		  <%next%>
        </select>
        月 
        <select name="day1">
		  <%for i=1 to 31
		    j="0"&i
		  %>
          <option value="<%=right(j,2)%>" <%if j=mid(rs("time"),7,2) then response.write" selected"%>><%=right(j,2)%></option>
		  <%next%>
        </select>
        日
        <select name="hour1">          
		  <%for i=1 to 24
		    j="0"&i
		  %>
          <option value="<%=right(j,2)%>" <%if j=mid(rs("time"),9,2) then response.write" selected"%>><%=right(j,2)%></option>
		  <%next%>
        </select>
        点       
        <input name=minute1 size=2 value="<%=right(rs("time"),2)%>">分
        </td>        
		</tr>
		<tr>
			<td bgcolor="#EFEFEF">提醒时间：</td><td bgcolor="#FFFFFF">
			<select name="year2">
		  <%for i=year(now()) to year(now())+4%>
          <option value="<%=i%>" <%if right(rs("remindtime"),4)=i then response.write" selected" end if%>><%=i%></option>
		  <%next%>
        </select>
        年 
        <select name="month2">
		  <%for i=1 to 12
		    j="0"&i
		  %>
          <option value="<%=right(j,2)%>" <%if j=mid(rs("remindtime"),5,2) then response.write" selected"%>><%=right(j,2)%></option>
		  <%next%>
        </select>
        月 
        <select name="day2">
		  <%for i=1 to 31
		    j="0"&i
		  %>
          <option value="<%=right(j,2)%>" <%if j=mid(rs("remindtime"),7,2) then response.write" selected"%>><%=right(j,2)%></option>
		  <%next%>
        </select>
        日
        <select name="hour2">          
		  <%for i=1 to 24
		    j="0"&i
		  %>
          <option value="<%=right(j,2)%>" <%if j=mid(rs("remindtime"),9,2) then response.write" selected"%>><%=right(j,2)%></option>
		  <%next%>
        </select>
        点       
        <input name=minute2 size=2 value="<%=right(rs("remindtime"),2)%>">分
        </td>        
		</tr>
		<tr>
			
      <td valign=top bgcolor="#bfbfbf">活动内容：</td>
      <td bgcolor="#efefef">
<TEXTAREA NAME="content" ROWS="6" COLS="50"><%=rs("content")%></TEXTAREA></td>			
		</tr>				
    </table>  
	  <div align="center"><br></div>  
  <div align="center"><a href='Javascript:DelChk(<%=rs("id")%>);' onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image2','','images/delete_on.gif',1)"><img name="Image2" border="0" src="images/delete_off.gif" width="60" height="19" hspace="5" ></a><a href="Javascript:document.calendar.submit();" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image3','','images/modify_on.gif',1)"><img name="Image3" border="0" src="images/modify_off.gif" width="60" height="19" hspace="5"></a><a href="Javascript:window.close();" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image1','','images/close_2.gif',1)"><img name="Image1" border="0" src="images/close_1.gif" width="85" height="19" onClick="MM_callJS('window.close();')" hspace="5"></a> 
  </div>
</form>
    

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
	function DelChk(calid)
	{
		if(confirm("确认删除吗?"))
			document.location="delcalendar.asp?id="+calid ;
	}
</script>
<%
end if
end if
%>
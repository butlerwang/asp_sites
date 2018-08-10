<!--#INCLUDE FILE="data.asp" -->
<!--#INCLUDE FILE="check.asp" -->
<%
strSQL ="SELECT * FROM calendar where id="&request("id")
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open strSQL,Conn,1,3
if not rs.eof then
rs("state")=true
rs.update 

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
	function DelChk(calid)
	{
		if(confirm("确认删除吗?"))
			document.location="delcalendar.asp?id="+calid ;
	}

</script>

<html>
<head>
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

function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
</head>
<body bgcolor="#efefef" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('images/delete_on.gif','images/modify_on.gif','images/close_2.gif')" >
<table width="100%" border="0" cellspacing="1" cellpadding="2">
  <tr bgcolor="#4E5960"> 
    <td class="heading" height="20"><font color="#FFFFFF"><b><font color=white>注意：</font>日程活动自动提醒！</b></font></td>
  </tr>
</table>

  
<table width="100%" border="1" cellspacing="0" cellpadding="1" bordercolordark="#FFFFFF" bordercolor="#000000">
  <tr> 
    <td width=15% bgcolor="#bfbfbf">日程主题 
      ：</td>
    <td><%=rs("title")%></td>
  </tr>
  <tr> 
    <td width=15% bgcolor="#bfbfbf">活动时间 
      ：</td>
    <td><%=left(rs("time"),4)%>/<%=mid(rs("time"),5,2)%>/<%=mid(rs("time"),7,2)%>&nbsp;&nbsp;<%=mid(rs("time"),9,2)%>:<%=right(rs("time"),2)%></td>
  </tr>
  <tr> 
    <td width=15% bgcolor="#bfbfbf">提醒时间 
      ：</td>
    <td><%=left(rs("remindtime"),4)%>/<%=mid(rs("remindtime"),5,2)%>/<%=mid(rs("remindtime"),7,2)%>&nbsp;&nbsp;<%=mid(rs("remindtime"),9,2)%>:<%=right(rs("remindtime"),2)%></td>
  </tr>
  <tr> 
    <td width=15% bgcolor="#bfbfbf">详细内容 
      ：</td>
    <td><%=rs("content")%></td>
  </tr>
  
</table>
<div align="center"><br></div>
<div align="center"><a href="Javascript:DelChk(<%=rs("id")%>);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image1','','images/delete_on.gif',1)"><img name="Image1" border="0" src="images/delete_off.gif" width="60" height="19" hspace="5"></a> 
  <a href="modCalendar.asp?id=<%=rs("id")%>" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image2','','images/modify_on.gif',1)"><img name="Image2" border="0" src="images/modify_off.gif" width="60" height="19" hspace="5"></a> 
  <a href="Javascript:window.close();" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image3','','images/close_2.gif',1)"><img name="Image3" border="0" src="images/close_1.gif" width="85" height="19" hspace="5"></a> 
</div>
</body>
</html>
<%
else 
response.write "error"
end if
%>
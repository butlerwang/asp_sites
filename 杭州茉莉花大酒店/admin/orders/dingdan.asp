<!--#include file="../check.asp"-->
<!--#include file="../include/conn.asp"-->
<!--#include file="../include/lib.asp"-->
<%
Call OpenData()
%>
<%
If Session("name") = "" then
response.Write "<script LANGUAGE=javascript>alert('网络超时，或者你还没有登陆! ');this.location.href='../login.asp';<'/script>"
end if
 	temp_check_rights = Split(session("manconfig"),",")
	for temp_rights_count=0 to ubound(temp_check_rights)
	    if trim(temp_check_rights(temp_rights_count)) = "8" then
			rights_check_passkey = trim(temp_check_rights(temp_rights_count))
		end if
	next
	if rights_check_passkey <> "8" then
Session.Abandon()
	Response.Write "<Script Language=JavaScript>alert('您的权限不足，请不要非法调用其它管理模块，否则您的账号将被系统自动删除!');this.location.href='../login.asp';</Script>"
	Response.end
	end if%>
<%
if Request("method")="ok" then
	if Request("chkid") = "" then
	Response.Write("<script>alert(""参出错误,你要删除的订单ID不能为空。"");location.href="""& Request.ServerVariables("HTTP_REFERER") &""";</script>")
	   Response.end     
	end if
	chkid = Request("chkid")
	Conn.Execute("delete from Sbe_order where ID in("& Request("chkid") &")")
	Response.Write("<script>alert(""删除成功"");location.href="""& Request.ServerVariables("HTTP_REFERER") &""";</script>") 
	Response.End()
end if
if request("keyword")="ok" then
   keyword=request("keyword")
   ly_time1=Cdate(request("ly_time1"))
   ly_time2=Cdate(request("ly_time2"))
   flag=request("flag")
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>订单管理</title>
<link href="../include/style.css" rel="stylesheet" type="text/css">
<SCRIPT language=JavaScript>
// 检测浏览器
NS4 = document.layers && true;
IE4 = document.all && parseInt(navigator.appVersion) >= 4;
function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.name != 'chkall')
       e.checked = form.chkall.checked;
    }
  }
</script>
<script language="JavaScript" src="../include/meizzDate.js"></script>
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="19%" height="25"><font color="#6A859D">订单管理 &gt;&gt;订单管理 </font></td>   
      <td width="81%">&nbsp;       </td>   
  </tr>
  <tr> 
    <td height="1" colspan="2" background="../images/dot.gif"></td>
  </tr>
</table>
<table width="98%"  border="0" cellpadding="0" cellspacing="1" align="center" bgcolor="#99CCFF">
  <tr bgcolor="#E8F1FF">
    <td height="22" colspan="9" class="table_1_2"><table width="100%"  border="0" cellpadding="0" cellspacing="0" class="table_1_2">
      <tr align="left" class="table_1_3_1">
        <form name="newsearch" action="dingdan.asp" method="post">
          <td width="88" align="right"  valign="middle" height="25" class="sbe_table_title">订单查找：</td>
          <td width="806"  colspan="9" align="left"  valign="middle" class="sbe_table_title">时间：
            <input name="keyword" type="hidden" value="ok">
                <input name="ly_time1" type="text" class="sbe_button" id="ly_time1"   style="ime-mode:Disabled" onFocus="setday(this)" onKeyPress="return   event.keyCode>=48&&event.keyCode<=57||event.keyCode==45"  <%if ly_time1<>"" then response.Write("value='"&ly_time1&"'") else response.Write("value='"&date()-6&"'") end if%> size="15" onpaste="return!clipboardData.getData('text').match(/\D/)" ondragenter="return false">
            至
            <input name="ly_time2" type="text" class="sbe_button" id="ly_time2"   style="ime-mode:Disabled" onFocus="setday(this)" onKeyPress="return   event.keyCode>=48&&event.keyCode<=57||event.keyCode==45" <%if ly_time2<>"" then response.Write("value='"&ly_time2&"'") else response.Write("value='"&date()&"'") end if%> size="15" onpaste="return!clipboardData.getData('text').match(/\D/)" ondragenter="return false">
            &nbsp;	状态：
            <select name="flag" size="1" class="sbe_button" style="width:150">
              <option value="" <%if flag ="" then response.Write("selected") end if%>>全部订单</option>
              <option value="0" <%if flag ="0" then response.Write("selected") end if%>>未处理的订单</option>
              <!--                      <option value="1" <%'if flag ="1" then response.Write("selected") end if%>>未回复的订单</option>-->
              <option value="1" <%if flag ="1" then response.Write("selected") end if%>>已处理的订单</option>
            </select>
            <input type="submit" name="Submit2" value="开始查找" class="anniugaodu" title="订单查找">
            &nbsp;
            <input type="button" name="ref" value="刷新页面" onClick="location.href='dingdan.asp'"  class="anniugaodu" title="不能自动刷新点击"></td>
        </form>
      </tr>
    </table></td>
  </tr>
  <tr align="left" bgcolor="#E8F1FF">
    <td  width="4%" align="center" valign="middle" nowrap bgcolor="#FFFFFF" class="sbe_button" ><strong>&nbsp;操作</strong>&nbsp;</td>
    <td width="16%"  height="22" align="center" valign="middle" nowrap bgcolor="#FFFFFF" class="sbe_button"><strong>&nbsp;房间类型&nbsp;</strong></td>
    <td width="20%"  align="center" valign="middle" nowrap bgcolor="#FFFFFF" class="sbe_button"><strong>&nbsp;姓名&nbsp;</strong></td>
    <td width="9%"  height="22" align="center" valign="middle" nowrap bgcolor="#FFFFFF" class="sbe_button"><strong>&nbsp;入住时间&nbsp;</strong></td>
    <td width="18%"  align="center" valign="middle" nowrap bgcolor="#FFFFFF" class="sbe_button"><strong>离开时间</strong></td>
    <td width="12%"  align="center" valign="middle" nowrap bgcolor="#FFFFFF" class="sbe_button"><strong>电话</strong>&nbsp;</td>
    <td width="11%"  align="center" valign="middle" nowrap bgcolor="#FFFFFF" class="sbe_button"><strong>提交日期</strong></td>
    <td width="5%"  align="center" valign="middle" nowrap bgcolor="#FFFFFF" class="sbe_button"><strong>&nbsp;状态&nbsp;</strong></td>
    <td width="5%" align="center"  valign="middle" nowrap bgcolor="#FFFFFF" class="sbe_button">&nbsp;<strong>查看</strong>&nbsp;</td>
  </tr>
  <%
		'-----修改每页显示个数 Start--------
	const MaxPerPage=15
	'-----修改每页显示个数 End  --------
   	dim totalPut
   	dim CurrentPage
	if not isempty(request("page")) then
      		currentPage=cint(request("page"))
   	else
      		currentPage=1
   	end if 
  set rs=server.createobject("adodb.recordset")
  Sql = "Select * from Sbe_order where 5=5 "
  if flag <> "" then Sql = Sql&"and status ="&flag&" "  
  if ly_time1 <> "" and ly_time2<>"" then Sql = Sql&" and timing between #"&ly_time1&"# and #"&ly_time2&"# "
  Sql = Sql&" order by ID desc"
  rs.open Sql,conn,1,1
  if rs.eof and rs.bof then
'     if flag="2" then
'	    Response.Write "<tr align=left bgcolor=#E8F1FF><td height=20 colspan=12 valign=middle>"&ly_time&"还没有<font color='#FF0000'>已回复</font>的订单信息。</td></tr>"
    if flag="1" then
        response.Write("<tr align=left bgcolor=#E8F1FF><td height=20 colspan=12 valign=middle>"&ly_time1&"至"&ly_time2&"还没有<font color='#FF0000'>已处理</font>的订单信息</td></tr>")
      elseif flag="0" then
        response.Write("<tr align=left bgcolor=#E8F1FF><td height=20 colspan=12 valign=middle>"&ly_time1&"至"&ly_time2&"还没有<font color='#FF0000'>未处理</font>的订单信息</td></tr>")
	  else
	    response.Write("<tr align=left bgcolor=#E8F1FF><td height=20 colspan=12 valign=middle>还没有订单信息</td></tr>")
     end if
	 else
	  %>
  <tr align="left" bgcolor="#E8F1FF" style="display:none">
    <td height="20" colspan="9" valign="middle" bgcolor="#FFFFFF" class="sbe_button"><%
 
       		totalPut=rs.recordcount
      		if currentpage<1 then
          		currentpage=1
      		end if
      		if (currentpage-1)*MaxPerPage>totalput then
	   		if (totalPut mod MaxPerPage)=0 then
	     			currentpage= totalPut \ MaxPerPage
	  		else
	      			currentpage= totalPut \ MaxPerPage + 1
	   		end if
      		end if
       		if currentPage=1 then
           		showpage totalput,MaxPerPage,"dingdan.asp"
            		showContent
            		showpage totalput,MaxPerPage,"dingdan.asp"
       		else
          		if (currentPage-1)*MaxPerPage<totalPut then
            			rs.move  (currentPage-1)*MaxPerPage
            			dim bookmark
            			bookmark=rs.bookmark
           			showpage totalput,MaxPerPage,"dingdan.asp"
            			showContent
             			showpage totalput,MaxPerPage,"dingdan.asp"
        		else
	        		currentPage=1
           			showpage totalput,MaxPerPage,"dingdan.asp"
           			showContent
           			showpage totalput,MaxPerPage,"dingdan.asp"
	      		end if
	   		end if
		rs.close
		set rs = nothing	
   	end if 
	sub showContent
	dim i 
	   	i=0
  %></td>
  </tr>
  <form action="dingdan.asp?method=ok" name="form1" method="post">
    <%do while not rs.eof%>
    <tr align="left" bgcolor="#E8F1FF">
      <td align="center" valign="middle" bgcolor="#FFFFFF"><input type="checkbox" name="chkid" value="<%=rs("ID")%>"></td>
      <td height="22" align="center" valign="middle" bgcolor="#FFFFFF">	 
	   <%set rs2=server.CreateObject("adodb.recordset")
	  sql="select * from sbe_product_class where id="&rs("roomtype")&""
	  rs2.open sql,conn,1,1
	  fang=rs2("classname")
	  rs2.close%>
	  <%=fang%>
</td>
      <td align="center" valign="middle" nowrap bgcolor="#FFFFFF">&nbsp;<%=rs("username")%>&nbsp;</td> 
      <td height="22" align="center" valign="middle" bgcolor="#FFFFFF"><a href="dd_show.asp?id=<%=rs("ID")%>"><%=gotTopic(rs("kssj"),20)%></a></td>
      <td align="center" valign="middle" nowrap bgcolor="#FFFFFF"><%=gotTopic(rs("lksj"),30)%></td>
      <td align="center" valign="middle" nowrap bgcolor="#FFFFFF"><%=(rs("tel"))%></td>
      <td align="center" valign="middle" bgcolor="#FFFFFF"><%=rs("time")%></td>
      <td align="center" valign="middle" nowrap bgcolor="#FFFFFF">&nbsp;
          <%if rs("status") = 1 then 
	       response.Write("<font color='#FF0000'>已处理</font>")
        else
	       response.Write("<font color='#000099'>未处理</font>")
	   end if
	 %>
        &nbsp;</td>
      <td align="center" valign="middle" nowrap bgcolor="#FFFFFF"><a href="dd_show.asp?id=<%=rs("ID")%>"><img src="../images/edit.gif" border="0"></a></td>
    </tr>
    <%
  i=i+1
	      if i>=MaxPerPage then exit do
  rs.movenext
  loop
  %>
    <tr align="left" bgcolor="#E8F1FF">
      <td  height="25" colspan="9" valign="middle" bgcolor="#FFFFFF">&nbsp;管理操作：全部选择
        <input type="checkbox" name="chkall" onClick="javascript:CheckAll(this.form)">
          <input type="submit" name="Submit" value="删除" class="sbe_button" >
        &nbsp;&nbsp; </td>
    </tr>
  </form>
  <tr align="left" bgcolor="#E8F1FF">
    <td  height="25" colspan="9" valign="middle" bgcolor="#FFFFFF"><%
    end sub
	function showpage(totalnumber,maxperpage,filename)
  	dim n

  	if totalnumber mod maxperpage=0 then
     		n= totalnumber \ maxperpage
  	else
     		n= totalnumber \ maxperpage+1          
  	end if
  	response.write "<table cellspacing=1 width='100%' border=0 colspan='4' ><form method=Post action="""&filename&"?keyword="& keyword&"&flag="& flag &"&year_time="& year_time &"&day_time="& day_time &"&month_time="& month_time &"""><tr><td align=right> "
  	if CurrentPage<2 then
    		response.write "共<b><font color=red>"&totalnumber&"</font></b>条记录&nbsp;首页 上一页&nbsp;"
  	else
    		response.write "共<b><font color=red>"&totalnumber&"</font></b>条记录&nbsp;<a href="&filename&"?page=1&keyword="& keyword&"&flag="& flag &"&year_time="& year_time &"&day_time="& day_time &"&month_time="& month_time &">首页</a>&nbsp;"
    		response.write "<a href="&filename&"?page="&CurrentPage-1&"&keyword="& keyword&"&flag="& flag &"&year_time="& year_time &"&day_time="& day_time &"&month_time="& month_time &">上一页</a>&nbsp;"
  	end if

  	if n-currentpage<1 then
    		response.write "下一页 尾页"
  	else
    		response.write "<a href="&filename&"?page="&(CurrentPage+1)&"&keyword="& keyword &"&flag="& flag &"&year_time="& year_time &"&day_time="& day_time &"&month_time="& month_time &">"
    		response.write "下一页</a> <a href="&filename&"?page="&n&"&keyword="& keyword&"&flag="& flag &"&year_time="& year_time &"&day_time="& day_time &"&month_time="& month_time &">尾页</a>"
  	end if
   	response.write "&nbsp;页次：<strong><font color=red>"&CurrentPage&"</font>/"&n&"</strong>页 "
    	response.write "&nbsp;<b>"&maxperpage&"</b>条记录/页 "
%>
      转到：
      <select name='page' size='1' class="sbe_button" style="font-size: 9pt" onChange='javascript:submit()'>
          <%for i = 1 to n%>
          <option value='<%=i%>' <%if CurrentPage=cint(i) then%> selected <%end if%>>第<%=i%>页</option>
          <%next%>
        </select>
        <%   
	response.write "</td></tr></FORM></table>"
end function
%></td>
  </tr>
</table>
<%Call CloseDataBase()%>
</body>
</html>

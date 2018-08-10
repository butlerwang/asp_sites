<!--#include file="../check.asp"-->
<!--#include file="../include/conn.asp"-->
<!--#include file="../include/lib.asp"-->
<%Call OpenData()
If Session("name") = "" then
response.Write "<script LANGUAGE=javascript>alert('网络超时，或者你还没有登陆! ');this.location.href='../login.asp';</script>"
end if
  IF instr(webConfig,", 7")=0 or instr(session("manconfig"),", 7")=0 Then'网站功能配置
Session.Abandon()
	Response.Write "<Script Language=JavaScript>alert('您的权限不足，请不要非法调用其它管理模块，否则您的账号将被系统自动删除!');this.location.href='../login.asp';</Script>"
Response.end
end if
if Request("method")="ok" then
	if Request("chkid") = "" then
	Response.Write("<script>alert(""参出错误,你要删除的留言ID不能为空。"");location.href="""& Request.ServerVariables("HTTP_REFERER") &""";</script>")
	   Response.end     
	end if
	chkid = Request("chkid")
	Conn.Execute("delete from Guest_book where ID in("& Request("chkid") &")")
	Response.Write("<script>alert(""删除成功"");location.href="""& Request.ServerVariables("HTTP_REFERER") &""";</script>") 
	Response.End()
end if
if request("keyword")="ok" then
      keyword=request("keyword")
      ly_time1=Cdate(request("ly_time1"))
      ly_time2=Cdate(request("ly_time2"))
	  leibie=request("leibie")
      flag=request("flag")
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>在线留言</title>
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
    <td width="19%" height="25"><font color="#6A859D">服务中心 &gt;&gt; 在线留言 </font></td>   
      <td width="81%">&nbsp;       </td>   
  </tr>
  <tr> 
    <td height="1" colspan="2" background="../images/dot.gif"></td>
  </tr>
</table>
<table width="98%"  border="0" cellpadding="0" cellspacing="1" align="center" bgcolor="#99CCFF">
  <tr bgcolor="#E8F1FF">
    <td height="22" colspan="10" class="table_1_2"><table width="100%"  border="0" cellpadding="0" cellspacing="0" class="table_1_2">
      <tr align="left" class="table_1_3_1">
        <form name="newsearch" action="list.asp" method="post">
          <td width="88" align="right"  valign="middle" height="25" class="sbe_table_title">留言查找：</td>
          <td width="806"  colspan="9" align="left"  valign="middle" class="sbe_table_title">时间：
            <input name="keyword" type="hidden" value="ok">
                <input name="ly_time1" type="text" class="sbe_button" id="ly_time1"   style="ime-mode:Disabled" onFocus="setday(this)" onKeyPress="return   event.keyCode>=48&&event.keyCode<=57||event.keyCode==45"  <%if ly_time1<>"" then response.Write("value='"&ly_time1&"'") else response.Write("value='"&date()-6&"'") end if%> size="15" onpaste="return!clipboardData.getData('text').match(/\D/)" ondragenter="return false">
            至
            <input name="ly_time2" type="text" class="sbe_button" id="ly_time2"   style="ime-mode:Disabled" onFocus="setday(this)" onKeyPress="return   event.keyCode>=48&&event.keyCode<=57||event.keyCode==45" <%if ly_time2<>"" then response.Write("value='"&ly_time2&"'") else response.Write("value='"&date()&"'") end if%> size="15" onpaste="return!clipboardData.getData('text').match(/\D/)" ondragenter="return false">
            &nbsp;	状态：
			<%if hy_message=true then%>
            <select name="flag" size="1" class="sbe_button" style="width:100">
              <option value="" <%if flag ="" then response.Write("selected") end if%>>全部留言</option>
              <option value="0" <%if flag ="0" then response.Write("selected") end if%>>未查看的留言</option>
              <option value="1" <%if flag ="1" then response.Write("selected") end if%>>未回复的留言</option>
              <option value="2" <%if flag ="2" then response.Write("selected") end if%>>已回复的留言</option>
            </select>
			<%else%>
			<select name="flag" size="1" class="sbe_button" style="width:100">
              <option value="" <%if flag ="" then response.Write("selected") end if%>>全部留言</option>
              <option value="0" <%if flag ="0" then response.Write("selected") end if%>>未查看的留言</option>
              <option value="1" <%if flag ="1" then response.Write("selected") end if%>>已查看的留言</option>
            </select>
			<%end if%>
			&nbsp;
			<select <%=banben_display%> name="leibie" size="1" class="sbe_button">
              <option value="" <%if leibie ="" then response.Write("selected") end if%>>全部类别</option>
              <option value="1" <%if leibie ="1" then response.Write("selected") end if%>>中文留言</option>
              <option value="2" <%if leibie ="2" then response.Write("selected") end if%>>英文留言</option>
            </select>
            <input type="submit" name="Submit2" value="开始查找" class="sbe_button" title="信息查找">
            &nbsp;
            <input type="button" name="ref" value="刷新页面" onClick="location.href='list.asp'"  class="sbe_button" title="不能自动刷新点击"></td>
        </form>
      </tr>
    </table></td>
  </tr>
  <tr align="left" bgcolor="#E8F1FF">
    <td  width="5%" align="center" valign="middle" nowrap bgcolor="#FFFFFF" class="sbe_button" ><strong>&nbsp;操作</strong>&nbsp;</td>
    <td width="16%"  align="center" valign="middle" nowrap bgcolor="#FFFFFF" class="sbe_button"  style="display:none"><strong>&nbsp;留言内容</strong></td>
    <td width="17%"  align="center" valign="middle" nowrap bgcolor="#FFFFFF" class="sbe_button"  ><strong>留言主题</strong></td>
    <td width="9%"  height="22" align="center" valign="middle" nowrap bgcolor="#FFFFFF" class="sbe_button"><strong>&nbsp;联系人&nbsp;</strong></td>
    <td width="13%"  height="22" align="center" valign="middle" nowrap bgcolor="#FFFFFF"  style="display:none"><strong>&nbsp;联系电话&nbsp;</strong></td>
    <td width="12%"  align="center" valign="middle" nowrap bgcolor="#FFFFFF"   style="display:none"><strong>&nbsp;E-mail</strong>&nbsp;</td>
    <td width="10%"  align="center" valign="middle" nowrap bgcolor="#FFFFFF" class="sbe_button"><strong>留言日期</strong></td>
    <td width="6%"  align="center" valign="middle" nowrap bgcolor="#FFFFFF" class="sbe_button" <%=banben_display%>><strong>&nbsp;类别&nbsp;</strong></td>
    <td width="6%"  align="center" valign="middle" nowrap bgcolor="#FFFFFF" class="sbe_button"><strong>&nbsp;状态&nbsp;</strong></td>
    <td width="6%" align="center"  valign="middle" nowrap bgcolor="#FFFFFF" class="sbe_button"><strong>查看</strong></td>
  </tr>
  <% 
  set rs1=server.createobject("adodb.recordset")
  Sql = "Select * from Guest_book where flag=1 "
  if flag <> "" then Sql = Sql&"and status ="&flag&" "
  if leibie<>"" then Sql=Sql&" and leibie="&leibie&""
  if ly_time1 <> "" and ly_time2<>"" then Sql = Sql&"and lytime between #"&ly_time1&"# and #"&ly_time2&"# "
  Sql = Sql&" order by ID desc"
  rs1.open Sql,conn,1,1
  if rs1.eof and rs1.bof then
  if hy_message=true then
    if flag="2" then
	    Response.Write "<tr align=left bgcolor=#E8F1FF><td height=20 colspan=12 valign=middle>"&ly_time&"还没有<font color='#FF0000'>已回复</font>的留言信息。</td></tr>"
    elseif flag="1" then
        response.Write("<tr align=left bgcolor=#E8F1FF><td height=20 colspan=12 valign=middle>"&ly_time1&"至"&ly_time2&"还没有<font color='#FF0000'>未回复</font>的留言信息</td></tr>")
      elseif flag="0" then
        response.Write("<tr align=left bgcolor=#E8F1FF><td height=20 colspan=12 valign=middle>"&ly_time1&"至"&ly_time2&"还没有<font color='#FF0000'>未查看</font>的留言信息</td></tr>")
	  else
	    response.Write("<tr align=left bgcolor=#E8F1FF><td height=20 colspan=12 valign=middle>还没有留言信息</td></tr>")
     end if
	else
	 if flag="1" then
        response.Write("<tr align=left bgcolor=#E8F1FF><td height=20 colspan=12 valign=middle>"&ly_time1&"至"&ly_time2&"还没有<font color='#FF0000'>已查看</font>的留言信息</td></tr>")
      elseif flag="0" then
        response.Write("<tr align=left bgcolor=#E8F1FF><td height=20 colspan=12 valign=middle>"&ly_time1&"至"&ly_time2&"还没有<font color='#FF0000'>未查看</font>的留言信息</td></tr>")
	  else
	    response.Write("<tr align=left bgcolor=#E8F1FF><td height=20 colspan=12 valign=middle>还没有留言信息</td></tr>")
     end if
	end if
	 else
	  %>
  <tr align="left" bgcolor="#E8F1FF" style="display:none">
    <td height="23" colspan="10" valign="middle" bgcolor="#FFFFFF" class="sbe_button"><%
	  rs1.pagesize=15
      totalrecord=rs1.recordcount
      totalpage=rs1.pagecount
	  pagenum=rs1.pagesize
      rs1.movefirst
      if request("page")="" then
         nowpage=1
		 elseif cint(request("page"))>totalpage then
         nowpage=totalpage
        else
      nowpage=request("page")
      end if
      nowpage=cint(nowpage)
      rs1.absolutepage=nowpage%></td>
  </tr>
  <form action="list.asp?method=ok" name="form1" method="post">
	<%p=1
	  Do while not rs1.EOF and p<=pagenum%>
    <tr align="left" bgcolor="#E8F1FF">
      <td align="center" valign="middle" bgcolor="#FFFFFF"><input type="checkbox" name="chkid" value="<%=rs1("ID")%>"></td>
      <td align="left" valign="middle" nowrap bgcolor="#FFFFFF" style="display:none">&nbsp;<%=gotTopic(rs1("lyremark"),50)%>&nbsp;</td> 
      <td align="left" valign="middle" nowrap bgcolor="#FFFFFF" ><%=gotTopic(rs1("lytheme"),20)%></td>
      <td height="22" align="center" valign="middle" bgcolor="#FFFFFF"><%=gotTopic(rs1("lyname"),20)%></td>  
      <td height="22" align="center" valign="middle" bgcolor="#FFFFFF"   style="display:none"><%=rs1("lytel")%></td>
      <td align="center" valign="middle" nowrap bgcolor="#FFFFFF"    style="display:none"><%=(rs1("lyemail"))%></td>
      <td align="center" valign="middle" bgcolor="#FFFFFF"><%=rs1("lytime")%></td>
      <td align="center" valign="middle" nowrap bgcolor="#FFFFFF" <%=banben_display%>><%if rs1("leibie")=1 then response.Write("中") else response.Write("英") end if%></td> 
      <td align="center" valign="middle" nowrap bgcolor="#FFFFFF">&nbsp;
      <%
	  if hy_message=true then
	  if rs1("status")=2 then
	       response.Write("已回复")
	    elseif rs1("status") = 1 then 
	       response.Write("<font color='#000099'>未回复</font>")
         else
	       response.Write("<font color='#FF0000'>未查看</font>")
	    end if
	  else
	  if rs1("status")=1 then
	       response.Write("已查看")
         else
	       response.Write("<font color='#FF0000'>未查看</font>")
	    end if
	end if
	 %>
        &nbsp;</td>
      <td align="center" valign="middle" nowrap bgcolor="#FFFFFF"><a href="ly_show.asp?id=<%=rs1("ID")%>"><img src="../images/edit.gif" border="0"></a></td>
    </tr>
	<%p=p+1
      rs1.moveNext
      loop%>
    <tr align="left" bgcolor="#E8F1FF">
      <td  height="25" colspan="10" valign="middle" bgcolor="#FFFFFF">&nbsp;管理操作：全部选择
        <input type="checkbox" name="chkall" onClick="javascript:CheckAll(this.form)">
          <input type="submit" name="Submit" value="删除" class="sbe_button" >
        &nbsp;&nbsp; </td>
    </tr>
  </form>
  <tr>
              <td height="25" colspan="10" align="right" bgcolor="#FFFFFF" class="zi11"><a href="?page=1&keyword=<%=keyword%>&ly_time1=<%=ly_time1%>&ly_time2=<%=ly_time2%>&leibie=<%=leibie%>&flag=<%=flag%>" title="首页" class="zi11">首页</a>&nbsp;&nbsp;
    <%if nowpage>1 then%><a href="?page=<%=nowpage-1%>&keyword=<%=keyword%>&ly_time1=<%=ly_time1%>&ly_time2=<%=ly_time2%>&leibie=<%=leibie%>&flag=<%=flag%>" title="上一页" class="zi11">上一页</a><%else%>上一页<%end if%>&nbsp;&nbsp;<%if nowpage<totalpage then%><a href="?page=<%=nowpage+1%>&keyword=<%=keyword%>&ly_time1=<%=ly_time1%>&ly_time2=<%=ly_time2%>&leibie=<%=leibie%>&flag=<%=flag%>" title="下一页" class="zi11">下一页</a><%else%>下一页<%end if%>&nbsp;&nbsp;<a href="?page=<%=totalpage%>&keyword=<%=keyword%>&ly_time1=<%=ly_time1%>&ly_time2=<%=ly_time2%>&leibie=<%=leibie%>&flag=<%=flag%>" title="最后页" class="zi11">最后页</a>&nbsp;&nbsp;页次：<%=nowpage%>/<%=totalpage%>&nbsp;&nbsp;<%=pagenum%>条记录/页&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  </tr>
  		  <%end if
		  rs1.close%>
</table>
<%Call CloseDataBase()%>
</body>
</html>

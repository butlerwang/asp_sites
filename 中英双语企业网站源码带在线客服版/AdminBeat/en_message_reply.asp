﻿
<%@ LANGUAGE=VBScript CodePage=65001%>
<% response.charset="utf-8" %>
<% session.codepage=65001 %>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="../inc/en_x_to_html/index_to_html.asp" -->
<!-- #include file="../inc/en_x_to_html/post_index_to_html.asp" -->
<% '更新数据到数据表
page=request.querystring("page")
act=request.querystring("act")
keywords=request.querystring("keywords")
article_id=cint(request.querystring("id"))

act1=Request("act1")
If act1="save" Then 
a_id=cint(request.form("a_id"))
a_recontent=trim(request.form("a_recontent"))
a_view_yes=trim(request.form("a_view_yes"))


set rs=server.createobject("adodb.recordset")
sql="select id,view_yes,recontent,retime from en_web_article_comment where id="&a_id&""
rs.open(sql),cn,1,3
rs("view_yes")=a_view_yes
rs("recontent")=a_recontent
rs("retime")=now()
rs.update
rs.close
set rs=nothing
call post_index_to_html()
response.Write "<script language='javascript'>alert('修改成功！');location.href='en_Message_list.asp?page="&page&"&act="&act&"&keywords="&keywords&"';</script>"
end if 
 %>
 

	<%
Call header()

%>

<%
      
			set rs=server.createobject("adodb.recordset")
sql="select id,name,content,recontent,view_yes from [en_web_article_comment] where id="&article_id&""
rs.open(sql),cn,1,1
if not rs.eof and not rs.bof then
%> <form id="form1" name="form1" method="post" action="?act1=save&page=<%=page%>&act=<%=act%>&keywords=<%=keywords%>">

	<table cellpadding='3' cellspacing='1' border='0' class='tableBorder' align=center>
	<tr>
	  <th class='tableHeaderText' colspan=2 height=25>回复留言</th>
	<tr>
	<td width="15%" height=23 class='forumRow'>留言人 </td>
	<td width="85%" class='forumRow'><input name='a_id' type='hidden' id='a_id' value="<%=rs("id")%>" size='70'>&nbsp;<%=rs("name")%></td>
	</tr>
	  <tr>
	    <td height=23 valign="top" class='forumRowHighLight'>留言内容 </td>
	    <td valign="top" class='forumRowHighLight'>&nbsp;<%=rs("content")%></td>
      </tr><tr>
	  <td  class='forumRow' height=11>回复内容</td>
	  <td  class='forumRow'><textarea name='a_recontent'  cols="100" rows="6" id="a_recontent" ><%=rs("recontent")%></textarea></td>
	</tr>
	<tr>
	  <td class='forumRowHighLight' height=23>是否显示</td>
	  <td class='forumRowHighLight'><label>
	    <input type="radio" name="a_view_yes" value="1"<%
		if rs("view_yes")=1 then
		response.write "checked"
		end if%>>
      是
      &nbsp;
      <input name="a_view_yes" type="radio" value="0" <%if rs("view_yes")=0 then
		response.write "checked"
		end if%>>
      否</label></td>
	</tr>
	<tr><td height="50" colspan=2  class='forumRow'><div align="center">
	  <INPUT type=submit value='提交'  name=Submit>
	  </div></td></tr>
	</table>
</form>
<%
else
response.write"未找到数据"
end if%>
<%
Call DbconnEnd()
 %>
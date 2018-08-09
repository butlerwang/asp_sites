<%@ LANGUAGE=VBScript CodePage=65001%>
<% response.charset="utf-8" %>
<% session.codepage=65001 %>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="../inc/en_x_to_html/index_to_html.asp" -->
<% '添加数据到数据表
act=Request("act")
If act="save" Then 
l_name=trim(request.form("name"))
l_memo=trim(request.form("memo"))
l_time=now()

set rs=server.createobject("adodb.recordset")
sql="select * from en_web_ads_position"
rs.open(sql),cn,1,3
rs.addnew
rs("name")=l_name
rs("memo")=l_memo
rs("time")=l_time

rs.update
rs.close
set rs=nothing
response.Write "<script language='javascript'>alert('添加成功！');location.href='en_ads_position_list.asp';</script>"
end if 
 %>
 

	<%
Call header()

%>

  <form id="form1" name="form1" method="post" action="?act=save">
         <script language='javascript'>
function checksignup1() {
if ( document.form1.name.value == '' ) {
window.alert('请输入广告位置^_^');
document.form1.name.focus();
return false;}



return true;}
</script>
	<table cellpadding='3' cellspacing='1' border='0' class='tableBorder' align=center>
	<tr>
	  <th class='tableHeaderText' colspan=2 height=25>添加广告位置</th>
	<tr>
	<td width="15%" height=23 class='forumRow'>广告位置 (必填) </td>
	<td width="85%" class='forumRow'><input name='name' type='text' id='name' size='70' maxlength="30">
	  &nbsp;</td>
	</tr><tr>
	  <td class='forumRow' height=11>备注</td>
	  <td class='forumRow'><textarea name='memo'  cols="100" rows="6" id="memo" ></textarea></td>
	</tr>
	<tr><td height="50" colspan=2  class='forumRow'><div align="center">
	  <INPUT type=submit value='提交' onClick='javascript:return checksignup1()' name=Submit>
	  </div></td></tr>
	</table>
</form>

<%
Call DbconnEnd()
 %>
<%@ LANGUAGE=VBScript CodePage=65001%>
<% response.charset="utf-8" %>
<% session.codepage=65001 %>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="../inc/x_to_html/chat_to_js.asp" -->
<%
act=Request("act")
If act="save" Then 
web_name=trim(request.form("web_name"))
web_wangwang=trim(request.form("web_wangwang"))
web_top=trim(request.form("web_top"))
web_Skype=trim(request.form("web_Skype"))
web_MSN=trim(request.form("web_MSN"))
web_stype=trim(request.form("web_stype"))
web_view_yes=trim(request.form("web_view_yes"))

set rs=server.createobject("adodb.recordset")
sql="select * from web_service"
rs.open(sql),cn,1,3
rs("name")=web_name
rs("wangwang")=web_wangwang
rs("Skype")=web_Skype
rs("MSN")=web_MSN
rs("view_yes")=web_view_yes
rs("top")=web_top
rs("stype")=web_stype
rs.update
rs.close
set rs=nothing

call chat_to_js()
response.Write "<script language='javascript' >alert('OK');location.href='chat_Settings.asp';</script>"
end if
 %>
 
	<%
Call header()

%>
<%set rs=server.createobject("adodb.recordset")
sql="select * from web_service"
rs.open(sql),cn,1,1
if not rs.eof and not rs.bof then
%>
  <form id="form1" name="form1" method="post" action="?act=save">
         <script language='javascript'>
function checksignup1() {
if ( document.form1.web_name.value == '' ) {
window.alert('请输入qq号^_^');
document.form1.web_name.focus();
return false;}

return true;}
</script>

	<table cellpadding='3' cellspacing='1' border='0' class='tableBorder' align=center>
	<tr>
	  <th class='tableHeaderText' colspan=2 height=31>在线客服设置</th>
	<tr>
	  <td height=23 colspan="2" class='forumRow'><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td height="20" class='TipTitle'>&nbsp;√ 操作提示</td>
        </tr>
        <tr>
          <td height="30" valign="top" class="TipWords"><p>1、内容保持为空将不会在控制面板中显示。</p>
          <p>2、设置完成默认只会更新首页，请到 静态管理 处 更新所有页面。</p>
</td>
        </tr>
        <tr>
          <td height="10">&nbsp;</td>
        </tr>
      </table></td>
	  </tr>
	<tr>
	  <td class='forumRowHighLight' height=23>是否显示客服系统</td>
	  <td class='forumRowHighLight'><label>
	       <input type="radio" name="web_view_yes" value="1"<%
		if rs("view_yes")=1 then
		response.write "checked"
		end if%>>
      显示
      &nbsp;
      <input name="web_view_yes" type="radio" value="0" <%if rs("view_yes")=0 then
		response.write "checked"
		end if%>>
      不显示</label></td>
	  </tr>
	<tr>
	  <td class='forumRowHighLight' height=23>客服类型</td>
	  <td class='forumRowHighLight'><label>
	       <input type="radio" name="web_stype" value="1"<%
		if rs("stype")=1 then
		response.write "checked"
		end if%>>
      红色
      &nbsp;
      <input name="web_stype" type="radio" value="2" <%if rs("stype")=2 then
		response.write "checked"
		end if%>>
      蓝色</label></td>
	  </tr>
  	<tr>
	<td width="15%" height=23 class='forumRow'>QQ号码</td>
	<td class='forumRow'><input name='web_name' type='text' id='web_name' value="<%=rs("name")%>" size='80'> 注：多个QQ号请使用 | 隔开。</td>
	</tr>
  	<tr>
	<td width="15%" height=23 class='forumRowHighLight'>旺旺链接</td>
	<td class='forumRowHighLight'><input name='web_wangwang' type='text' id='web_wangwang' value="<%=rs("wangwang")%>" size='80'> 注：更换旺旺号修改此处的uid=hitux,hitux即旺旺账号。 </td>
	</tr>
  	<tr>
	<td width="15%" height=23 class='forumRow'>Skype链接</td>
	<td class='forumRow'><input name='web_Skype' type='text' id='web_Skype' value="<%=rs("Skype")%>" size='80'> </td>
	</tr>    
  	<tr>
	<td width="15%" height=23 class='forumRowHighLight'>MSN链接</td>
	<td class='forumRowHighLight'><input name='web_MSN' type='text' id='web_MSN' value="<%=rs("MSN")%>" size='80'> </td>
	</tr>    
	<tr>
	  <td class='forumRow' height=23>客服图标与网站顶部距离</td>
	  <td class='forumRow'><input type='text' id='web_top' name='web_top'  value="<%=rs("top")%>" size='40'>px (像素)</td>
	</tr>

	<tr><td height="50" colspan=2  class='forumRow'><div align="center">
	  <INPUT type=submit value='提交' onClick='javascript:return checksignup1()' name=Submit>
	  </div></td></tr>
	</table>
</form>

<%
Call DbconnEnd()
else
response.write "暂时无数据"
end if %>
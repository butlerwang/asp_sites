<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="../inc/x_to_html/ADs_to_js.asp" -->
<!-- #include file="../inc/x_to_html/index_to_html.asp" -->
<% '添加数据到数据表
page=request.querystring("page")
act=request.querystring("act")
keywords=request.querystring("keywords")
article_id=cint(request.querystring("id"))


act1=Request("act1")
If act1="save" Then 
l_id=trim(request.form("l_id"))
l_name=trim(request.form("name"))
l_url=trim(request.form("url"))
l_image=trim(request.form("web_image"))
l_order=trim(request.form("order"))
l_ADtype=trim(request.form("ADtype"))
l_FlashUrl=trim(request.form("FlashUrl"))
l_ADcode=trim(request.form("ADcode"))
'l_ADWidth="0"
'l_ADHeight="0"
l_view_yes=trim(request.form("view_yes"))
l_time=now()

set rs=server.createobject("adodb.recordset")
sql="select * from web_ads where id="&l_id&""
rs.open(sql),cn,1,3
rs("name")=l_name
rs("url")=l_url
rs("position")=1
rs("image")=l_image
rs("FlashUrl")=l_FlashUrl
rs("ADcode")=l_ADcode
if l_order<>"" then
rs("order")=cint(l_order)
end if
rs("memo")=l_memo
rs("view_yes")=cint(l_view_yes)
rs("ADtype")=cint(l_ADtype)
'rs("ADWidth")=cint(l_ADWidth)
'rs("ADHeight")=cint(l_ADHeight)
rs.update
rs.close
set rs=nothing
%>
<% '生成广告JS
call ADs_to_js(l_id)
%>
<%
call index_to_html()
response.Write "<script language='javascript'>alert('修改成功！');location.href='ads_list.asp?page="&page&"&act="&act&"&keywords="&keywords&"';</script>"
end if 
 %>
 

	<%
Call header()

%>
<% set rs=server.createobject("adodb.recordset")
sql="select * from web_ads where id="&article_id&""
rs.open(sql),cn,1,1
if not rs.eof and not rs.bof then%>
  <form id="form1" name="form1" method="post" action="?act1=save&page=<%=page%>&act=<%=act%>&keywords=<%=keywords%>">
         <script language='javascript'>
function checksignup1() {
if ( document.form1.name.value == '' ) {
window.alert('请输入标题^_^');
document.form1.name.focus();
return false;}

if ( document.form1.position.value == '' ) {
window.alert('请选择位置^_^');
document.form1.position.focus();
return false;}

if ( document.form1.ADtype.value == '' ) {
window.alert('请选择类型^_^');
document.form1.ADtype.focus();
return false;}

if(document.form1.ADWidth.value.search(/^([0-9]*)([.]?)([0-9]*)$/)   ==   -1)   
      {   
  window.alert("长度只能是数字^_^");   
document.form1.ADWidth.focus();
return false;}

if(document.form1.ADHeight.value.search(/^([0-9]*)([.]?)([0-9]*)$/)   ==   -1)   
      {   
  window.alert("宽度只能是数字^_^");   
document.form1.ADHeight.focus();
return false;}

if(document.form1.order.value.search(/^([0-9]*)([.]?)([0-9]*)$/)   ==   -1)   
      {   
  window.alert("排序只能是数字^_^");   
document.form1.order.focus();
return false;}

return true;}
</script>
	<table cellpadding='3' cellspacing='1' border='0' class='tableBorder' align=center>
	<tr>
	  <th class='tableHeaderText' colspan=2 height=25>修改图片</th>
	<tr>
	  <td height=23 colspan="2" class='forumRow'><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td height="20" class='TipTitle'>&nbsp;√ 操作提示</td>
        </tr>
        <tr>
          <td height="30" valign="top" class="TipWords">
		  
		  	  <p>1、图片尺寸：980 x 300</p>
              <p>2、上传的图片，尽量处理的小点，尽量控制在100K以内，图片太大会影响首页打开速度！</p>
		
		</td>
        </tr>
        <tr>
          <td height="10">&nbsp;</td>
        </tr>
      </table></td>
	  </tr>
	<tr>
	<td width="15%" height=23 class='forumRowHighLight'>标题 (必填) </td>
	<td class='forumRowHighLight'><input name='name' type='text' id='name'  value="<%=rs("name")%>"size='70'>
	 <input name='l_id' type='hidden' id='l_id' value="<%=rs("id")%>" size='70'>
	  &nbsp;</td>
	</tr>

	  <tr>
	    <td height=23 class='forumRowHighLight'>类型</td>
	    <td class='forumRowHighLight'><label>
		<select name="ADtype" id="ADtype">

		<option value="2" selected>图片</option>
						
		</select>&nbsp;（尺寸：<font color="red">980x300</font>）
</label></td>
      </tr>
	  <tr>
	    <td class='forumRow' height=23>链接地址</td>
	    <td class='forumRow'><input name='url' type='text' id='url' value="<%=rs("url")%>" size='70'>填写“站内链接”或者“站外链接”均可</td>
      </tr>

	  <tr>
	    <td class='forumRowHighLight' height=23>上传图片 </td>
	    <td width="85%" class='forumRowHighLight'><table width="100%" border="0" cellspacing="0" cellpadding="0">
         <tr>
           <td width="22%" ><input name="web_image" type="text" id="web_image"  size="30" value="<%=rs("image")%>" readonly></td>
           <td width="78%"  ><iframe width="500" name="ad" frameborder=0 height=30 scrolling=no src=upload.asp?juhaoyongUploadFileName=<%=juhaoyongGetUploadFileName(rs("image"))%>&juhaoyongUpLoadPath=<%=juhaoyongGetweb_theme()%>></iframe><font color="#FF0000">（注：上传的图片名称必须是英文或数字）</font></td>
         </tr>
       </table></td>
      </tr>


<tr>
	    <td class='forumRowHighLight' height=23>排序</td>
	    <td class='forumRowHighLight'><span class="forumRow">
	      <input name='order' type='text' id='order' value="<%=rs("order")%>" size='20'>
	    &nbsp;只能是数字，数字越小排名越靠前</span></td>
      </tr>
	  
	  <tr>
	  <td class='forumRowHighLight' height=23>是否显示</td>
	  <td class='forumRowHighLight'><label>
	       <input type="radio" name="view_yes" value="1"<%
		if rs("view_yes")=1 then
		response.write "checked"
		end if%>>
      是
      &nbsp;
      <input name="view_yes" type="radio" value="0" <%if rs("view_yes")=0 then
		response.write "checked"
		end if%>>
      否</label></td>
	</tr>
	<tr><td height="50" colspan=2  class='forumRow'><div align="center">
	  <INPUT type=submit value='提交' onClick='javascript:return checksignup1()' name=Submit>
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
<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="../inc/x_to_html/ADs_to_js.asp" -->
<!-- #include file="../inc/x_to_html/index_to_html.asp" -->
<% '添加数据到数据表
act=Request("act")
If act="save" Then 
l_name=trim(request.form("name"))
l_url=trim(request.form("url"))
l_image=trim(request.form("web_image"))
l_order=trim(request.form("order"))
l_ADtype=trim(request.form("ADtype"))
l_FlashUrl=trim(request.form("FlashUrl"))
l_ADcode=trim(request.form("ADcode"))
l_ADWidth=trim(request.form("ADwidth"))
l_ADHeight=trim(request.form("ADHeight"))
l_view_yes=trim(request.form("view_yes"))
l_time=now()

set rs=server.createobject("adodb.recordset")
sql="select * from web_ads"
rs.open(sql),cn,1,3
rs.addnew
rs("name")=l_name
rs("url")=l_url
rs("position")=1
rs("FlashUrl")=l_FlashUrl
rs("ADcode")=l_ADcode
rs("image")=l_image
if l_order<>"" then
rs("order")=cint(l_order)
end if
rs("view_yes")=cint(l_view_yes)
rs("ADtype")=cint(l_ADtype)
if l_ADWidth<>"" then
rs("ADWidth")=cint(l_ADWidth)
end if
if l_ADHeight<>"" then
rs("ADHeight")=cint(l_ADHeight)
end if
rs("time")=l_time
rs.update
rs.close
set rs=nothing
%>
<% '生成广告JS
set rs2=server.createobject("adodb.recordset")
sql="select top 1 [id] from [web_ads] where [name]='"&l_name&"' order by [time] desc"
rs2.open(sql),cn,1,1
if not rs2.eof  then
l_id=rs2("id")
call ADs_to_js(l_id)
end if
rs2.close
set rs2=nothing
%>
<%
call index_to_html()
response.Write "<script language='javascript'>alert('添加成功！');location.href='ads_list.asp';</script>"
end if 
 %>
 

	<%
Call header()

%>

  <form id="form1" name="form1" method="post" action="?act=save">
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
<%juhaoyongGetLunboImgTotal=juhaoyongGetLunboImgTotal()%>
	<table cellpadding='3' cellspacing='1' border='0' class='tableBorder' align=center>
	<tr>
	  <th class='tableHeaderText' colspan=2 height=25>添加图片</th>
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
	<td class='forumRowHighLight'><input name='name' type='text' id='name' size='70'>
	  &nbsp;</td>
	</tr>

	  <tr>
	    <td height=23 class='forumRowHighLight'>类型</td>
	    <td class='forumRowHighLight'><label>
		<select name="ADtype" id="ADtype">
		<option value="2" selected >图片</option>
		</select>&nbsp;（尺寸：<font color="red">980x300</font>）</label></td>
      </tr>
	  <tr>
	    <td class='forumRow' height=23>链接地址</td>
	    <td class='forumRow'><input name='url' type='text' id='url' size='70'>填写“站内链接”或者“站外链接”均可</td>
      </tr>

	  <tr>
	    <td class='forumRowHighLight' height=23>上传图片 </td>
	    <td width="85%" class='forumRowHighLight'><table width="100%" border="0" cellspacing="0" cellpadding="0">
         <tr>
           <td width="22%" ><input name="web_image" type="text" id="web_image"  size="30" readonly></td>
           <td width="78%"  ><iframe width="500" name="ad" frameborder=0 height=30 scrolling=no src=upload.asp?juhaoyongUploadFileName=<%=juhaoyongGetLunboImgTotal%>&juhaoyongUpLoadPath=<%=juhaoyongGetweb_theme()%>></iframe><font color="#FF0000">（注：上传的图片名称必须是英文或数字）</font></td>
         </tr>
       </table></td>
      </tr>



	  <tr>
	    <td class='forumRowHighLight' height=23>排序</td>
	    <td class='forumRowHighLight'><span class="forumRow">
	      <input name='order' type='text' id='order' size='20' value="<%=juhaoyongGetLunboImgTotal%>">
	    &nbsp;只能是数字，数字越小排名越靠前</span></td>
      </tr>
	  
	  <tr>
	  <td class='forumRow' height=23>是否显示</td>
	  <td class='forumRow'><label>
	    <input type="radio" name="view_yes" value="1" checked>
      是
      &nbsp;
      <input name="view_yes" type="radio" value="0" >
      否</label></td>
	</tr>
	<tr><td height="50" colspan=2  class='forumRow'><div align="center">
	  <INPUT type=submit value='提交' onClick='javascript:return checksignup1()' name=Submit>
	  </div></td></tr>
	</table>
</form>

<%
Call DbconnEnd()
 %>
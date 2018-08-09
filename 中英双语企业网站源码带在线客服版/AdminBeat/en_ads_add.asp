<%@ LANGUAGE=VBScript CodePage=65001%>
<% response.charset="utf-8" %>
<% session.codepage=65001 %>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="../inc/en_x_to_html/ADs_to_js.asp" -->
<!-- #include file="../inc/en_x_to_html/index_to_html.asp" -->
<% '添加数据到数据表
act=Request("act")
If act="save" Then 
l_name=trim(request.form("name"))
l_url=trim(request.form("url"))
l_position=trim(request.form("position"))
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
sql="select * from en_web_ads"
rs.open(sql),cn,1,3
rs.addnew
rs("name")=l_name
rs("url")=l_url
rs("position")=l_position
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
sql="select top 1 [id] from [en_web_ads] where [name]='"&l_name&"' order by [time] desc"
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
response.Write "<script language='javascript'>alert('添加成功！');location.href='en_ads_list.asp';</script>"
end if 
 %>
 

	<%
Call header()

%>

  <form id="form1" name="form1" method="post" action="?act=save">
         <script language='javascript'>
function checksignup1() {
if ( document.form1.name.value == '' ) {
window.alert('请输入广告标题^_^');
document.form1.name.focus();
return false;}

if ( document.form1.position.value == '' ) {
window.alert('请选择广告位置^_^');
document.form1.position.focus();
return false;}

if ( document.form1.ADtype.value == '' ) {
window.alert('请选择广告类型^_^');
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
	  <th class='tableHeaderText' colspan=2 height=25>添加广告</th>
	<tr>
	  <td height=23 colspan="2" class='forumRow'><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td height="20" class='TipTitle'>&nbsp;√ 操作提示</td>
        </tr>
        <tr>
          <td height="30" valign="top" class="TipWords"><p>1、“广告类型”的选择决定你的广告的形式。比如如果你选择了“图片广告”，而“上传图片”处为空，即使你的“Flash文件地址”或“具体代码”里有数据，该广告都不会显示。</p>
              <p>2、对现有除“广告联盟代码”外的广告类型进行修改操作都不需要生成相应网页，即可查看到修改后的效果。</p>
              <p>3、“广告联盟代码”广告是指你注册的广告联盟提供给你的广告代码，一般以JS代码居多。常见的广告联盟有<a href="http://www.google.com/adsense/" target="_blank">谷歌广告联盟</a>、<a href="http://union.baidu.com/" target="_blank">百度广告联盟</a>等。</p>
              <p>4 、除“广告联盟代码”广告外，其它广告均以生成JS文件形式出现在网页中。 </p></td>
        </tr>
        <tr>
          <td height="10">&nbsp;</td>
        </tr>
      </table></td>
	  </tr>
	<tr>
	<td width="15%" height=23 class='forumRowHighLight'>广告标题 (必填) </td>
	<td class='forumRowHighLight'><input name='name' type='text' id='name' size='70'>
	  &nbsp;</td>
	</tr>
	  <tr>
	    <td class='forumRow' height=23>广告位置<span class="forumRowHighLight"> (必选) </span></td>
	    <td class='forumRow'><label>
	      <select name="position" id="position">
		  <option value="">请选择</option>
	       <% set rsp=server.createobject("adodb.recordset")
		   sql="select id,name from en_web_ads_position "
		   rsp.open(sql),cn,1,1
		   if not rsp.eof and not rsp.bof then
		   do while not rsp.eof 
		   %> <option value="<%=rsp("id")%>"><%=rsp("name")%></option>
            <%
			rsp.movenext
			loop
			end if
			rsp.close
			set rsp=nothing%></select>
	    </label></td>
      </tr>
	  <tr>
	    <td height=23 class='forumRowHighLight' style="font-weight: bold; color: #FF0000">广告类型(必填)</td>
	    <td class='forumRowHighLight'><label>
		<select name="ADtype" id="ADtype">
		<option value="1" >文字链 </option>
		<option value="2" selected >图片广告</option>
		<option value="3" >Flash广告</option>
		<option value="4" >广告联盟代码</option>					
		</select></label></td>
      </tr>
	  <tr>
	    <td class='forumRow' height=23><span style="color: #FF0000">文字链/图片广告</span>：链接</td>
	    <td class='forumRow'><input name='url' type='text' id='url' size='70'>
        &nbsp;建议以http://开头</td>
      </tr>

	  <tr>
	    <td class='forumRowHighLight' height=23><span style="color: #FF0000">图片广告</span>：上传图片 </td>
	    <td width="85%" class='forumRowHighLight'><table width="100%" border="0" cellspacing="0" cellpadding="0">
         <tr>
           <td width="22%" ><input name="web_image" type="text" id="web_image"  size="30"></td>
           <td width="78%"  ><iframe width="500" name="ad" frameborder=0 height=30 scrolling=no src=upload.asp></iframe></td>
         </tr>
       </table></td>
      </tr>
	  <tr>
	    <td class='forumRow' height=23><span style="color: #FF0000">Flash广告</span>：Flash文件地址</td>
	    <td class='forumRow'><span class="forumRowHighLight">
	      <input name='FlashUrl' type='text' id='FlashUrl' size='70' maxlength="200" />
	    </span></td>
      </tr>
	  <tr>
	    <td class='forumRowHighLight' height=23><span style="color: #FF0000">广告联盟代码</span>：具体代码</td>
	    <td class='forumRowHighLight'><textarea name='ADcode'  cols="100" rows="6" id="ADcode" ></textarea></td>
      </tr>
	  <tr>
	    <td class='forumRow' height=23>图片/Flash大小<span style="color: #FF0000"> (必填)</span></td>
	    <td class='forumRow'>宽度(Width)
          <input name='ADWidth' type='text' id='ADWidth' size='10' maxlength="4" /> 
        &nbsp;高度(Height)
        <input name='ADHeight' type='text' id='ADHeight' size='10' maxlength="4" />
        <span style="color: #FF0000">仅针对图片广告和FLASH广告，必填，否则广告无法显示。</span></td>
      </tr>
	  <tr>
	    <td class='forumRowHighLight' height=23>排序</td>
	    <td class='forumRowHighLight'><span class="forumRow">
	      <input name='order' type='text' id='order' value="1" size='20'>
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
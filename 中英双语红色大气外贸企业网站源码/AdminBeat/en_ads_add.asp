<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="../inc/en_x_to_html/ADs_to_js.asp" -->
<!-- #include file="../inc/en_x_to_html/index_to_html.asp" -->
<% '������ݵ����ݱ�
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
<% '���ɹ��JS
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
response.Write "<script language='javascript'>alert('��ӳɹ���');location.href='en_ads_list.asp';</script>"
end if 
 %>
 

	<%
Call header()

%>

  <form id="form1" name="form1" method="post" action="?act=save">
         <script language='javascript'>
function checksignup1() {
if ( document.form1.name.value == '' ) {
window.alert('�����������^_^');
document.form1.name.focus();
return false;}

if ( document.form1.position.value == '' ) {
window.alert('��ѡ����λ��^_^');
document.form1.position.focus();
return false;}

if ( document.form1.ADtype.value == '' ) {
window.alert('��ѡ��������^_^');
document.form1.ADtype.focus();
return false;}

if(document.form1.ADWidth.value.search(/^([0-9]*)([.]?)([0-9]*)$/)   ==   -1)   
      {   
  window.alert("����ֻ��������^_^");   
document.form1.ADWidth.focus();
return false;}

if(document.form1.ADHeight.value.search(/^([0-9]*)([.]?)([0-9]*)$/)   ==   -1)   
      {   
  window.alert("���ֻ��������^_^");   
document.form1.ADHeight.focus();
return false;}

if(document.form1.order.value.search(/^([0-9]*)([.]?)([0-9]*)$/)   ==   -1)   
      {   
  window.alert("����ֻ��������^_^");   
document.form1.order.focus();
return false;}

return true;}
</script>
	<table cellpadding='3' cellspacing='1' border='0' class='tableBorder' align=center>
	<tr>
	  <th class='tableHeaderText' colspan=2 height=25>��ӹ��</th>
	<tr>
	  <td height=23 colspan="2" class='forumRow'><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td height="20" class='TipTitle'>&nbsp;�� ������ʾ</td>
        </tr>
        <tr>
          <td height="30" valign="top" class="TipWords"><p>1����������͡���ѡ�������Ĺ�����ʽ�����������ѡ���ˡ�ͼƬ��桱�������ϴ�ͼƬ����Ϊ�գ���ʹ��ġ�Flash�ļ���ַ���򡰾�����롱�������ݣ��ù�涼������ʾ��</p>
              <p>2�������г���������˴��롱��Ĺ�����ͽ����޸Ĳ���������Ҫ������Ӧ��ҳ�����ɲ鿴���޸ĺ��Ч����</p>
              <p>3����������˴��롱�����ָ��ע��Ĺ�������ṩ����Ĺ����룬һ����JS����Ӷࡣ�����Ĺ��������<a href="http://www.google.com/adsense/" target="_blank">�ȸ�������</a>��<a href="http://union.baidu.com/" target="_blank">�ٶȹ������</a>�ȡ�</p>
              <p>4 ������������˴��롱����⣬��������������JS�ļ���ʽ��������ҳ�С� </p></td>
        </tr>
        <tr>
          <td height="10">&nbsp;</td>
        </tr>
      </table></td>
	  </tr>
	<tr>
	<td width="15%" height=23 class='forumRowHighLight'>������ (����) </td>
	<td class='forumRowHighLight'><input name='name' type='text' id='name' size='70'>
	  &nbsp;</td>
	</tr>
	  <tr>
	    <td class='forumRow' height=23>���λ��<span class="forumRowHighLight"> (��ѡ) </span></td>
	    <td class='forumRow'><label>
	      <select name="position" id="position">
		  <option value="">��ѡ��</option>
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
	    <td height=23 class='forumRowHighLight' style="font-weight: bold; color: #FF0000">�������(����)</td>
	    <td class='forumRowHighLight'><label>
		<select name="ADtype" id="ADtype">
		<option value="1" >������ </option>
		<option value="2" selected >ͼƬ���</option>
		<option value="3" >Flash���</option>
		<option value="4" >������˴���</option>					
		</select></label></td>
      </tr>
	  <tr>
	    <td class='forumRow' height=23><span style="color: #FF0000">������/ͼƬ���</span>������</td>
	    <td class='forumRow'><input name='url' type='text' id='url' size='70'>
        &nbsp;������http://��ͷ</td>
      </tr>

	  <tr>
	    <td class='forumRowHighLight' height=23><span style="color: #FF0000">ͼƬ���</span>���ϴ�ͼƬ </td>
	    <td width="85%" class='forumRowHighLight'><table width="100%" border="0" cellspacing="0" cellpadding="0">
         <tr>
           <td width="22%" ><input name="web_image" type="text" id="web_image"  size="30"></td>
           <td width="78%"  ><iframe width="500" name="ad" frameborder=0 height=30 scrolling=no src=upload.asp></iframe></td>
         </tr>
       </table></td>
      </tr>
	  <tr>
	    <td class='forumRow' height=23><span style="color: #FF0000">Flash���</span>��Flash�ļ���ַ</td>
	    <td class='forumRow'><span class="forumRowHighLight">
	      <input name='FlashUrl' type='text' id='FlashUrl' size='70' maxlength="200" />
	    </span></td>
      </tr>
	  <tr>
	    <td class='forumRowHighLight' height=23><span style="color: #FF0000">������˴���</span>���������</td>
	    <td class='forumRowHighLight'><textarea name='ADcode'  cols="100" rows="6" id="ADcode" ></textarea></td>
      </tr>
	  <tr>
	    <td class='forumRow' height=23>ͼƬ/Flash��С<span style="color: #FF0000"> (����)</span></td>
	    <td class='forumRow'>���(Width)
          <input name='ADWidth' type='text' id='ADWidth' size='10' maxlength="4" /> 
        &nbsp;�߶�(Height)
        <input name='ADHeight' type='text' id='ADHeight' size='10' maxlength="4" />
        <span style="color: #FF0000">�����ͼƬ����FLASH��棬����������޷���ʾ��</span></td>
      </tr>
	  <tr>
	    <td class='forumRowHighLight' height=23>����</td>
	    <td class='forumRowHighLight'><span class="forumRow">
	      <input name='order' type='text' id='order' value="1" size='20'>
	    &nbsp;ֻ�������֣�����ԽС����Խ��ǰ</span></td>
      </tr>
	  
	  <tr>
	  <td class='forumRow' height=23>�Ƿ���ʾ</td>
	  <td class='forumRow'><label>
	    <input type="radio" name="view_yes" value="1" checked>
      ��
      &nbsp;
      <input name="view_yes" type="radio" value="0" >
      ��</label></td>
	</tr>
	<tr><td height="50" colspan=2  class='forumRow'><div align="center">
	  <INPUT type=submit value='�ύ' onClick='javascript:return checksignup1()' name=Submit>
	  </div></td></tr>
	</table>
</form>

<%
Call DbconnEnd()
 %>
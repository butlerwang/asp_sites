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
window.alert('������qq��^_^');
document.form1.web_name.focus();
return false;}

return true;}
</script>

	<table cellpadding='3' cellspacing='1' border='0' class='tableBorder' align=center>
	<tr>
	  <th class='tableHeaderText' colspan=2 height=31>���߿ͷ�����</th>
	<tr>
	  <td height=23 colspan="2" class='forumRow'><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td height="20" class='TipTitle'>&nbsp;�� ������ʾ</td>
        </tr>
        <tr>
          <td height="30" valign="top" class="TipWords"><p>1�����ݱ���Ϊ�ս������ڿ����������ʾ��</p>
          <p>2���������Ĭ��ֻ�������ҳ���뵽 ��̬���� �� ��������ҳ�档</p>
</td>
        </tr>
        <tr>
          <td height="10">&nbsp;</td>
        </tr>
      </table></td>
	  </tr>
	<tr>
	  <td class='forumRowHighLight' height=23>�Ƿ���ʾ�ͷ�ϵͳ</td>
	  <td class='forumRowHighLight'><label>
	       <input type="radio" name="web_view_yes" value="1"<%
		if rs("view_yes")=1 then
		response.write "checked"
		end if%>>
      ��ʾ
      &nbsp;
      <input name="web_view_yes" type="radio" value="0" <%if rs("view_yes")=0 then
		response.write "checked"
		end if%>>
      ����ʾ</label></td>
	  </tr>
	<tr>
	  <td class='forumRowHighLight' height=23>�ͷ�����</td>
	  <td class='forumRowHighLight'><label>
	       <input type="radio" name="web_stype" value="1"<%
		if rs("stype")=1 then
		response.write "checked"
		end if%>>
      ��ɫ
      &nbsp;
      <input name="web_stype" type="radio" value="2" <%if rs("stype")=2 then
		response.write "checked"
		end if%>>
      ��ɫ</label></td>
	  </tr>
  	<tr>
	<td width="15%" height=23 class='forumRow'>QQ����</td>
	<td class='forumRow'><input name='web_name' type='text' id='web_name' value="<%=rs("name")%>" size='80'> ע�����QQ����ʹ�� | ������</td>
	</tr>
  	<tr>
	<td width="15%" height=23 class='forumRowHighLight'>��������</td>
	<td class='forumRowHighLight'><input name='web_wangwang' type='text' id='web_wangwang' value="<%=rs("wangwang")%>" size='80'> ע�������������޸Ĵ˴���uid=hitux,hitux�������˺š� </td>
	</tr>
  	<tr>
	<td width="15%" height=23 class='forumRow'>Skype����</td>
	<td class='forumRow'><input name='web_Skype' type='text' id='web_Skype' value="<%=rs("Skype")%>" size='80'> </td>
	</tr>    
  	<tr>
	<td width="15%" height=23 class='forumRowHighLight'>MSN����</td>
	<td class='forumRowHighLight'><input name='web_MSN' type='text' id='web_MSN' value="<%=rs("MSN")%>" size='80'> </td>
	</tr>    
	<tr>
	  <td class='forumRow' height=23>�ͷ�ͼ������վ��������</td>
	  <td class='forumRow'><input type='text' id='web_top' name='web_top'  value="<%=rs("top")%>" size='40'>px (����)</td>
	</tr>

	<tr><td height="50" colspan=2  class='forumRow'><div align="center">
	  <INPUT type=submit value='�ύ' onClick='javascript:return checksignup1()' name=Submit>
	  </div></td></tr>
	</table>
</form>

<%
Call DbconnEnd()
else
response.write "��ʱ������"
end if %>
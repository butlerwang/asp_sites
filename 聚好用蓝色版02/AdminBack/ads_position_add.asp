<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="../inc/x_to_html/index_to_html.asp" -->
<% '������ݵ����ݱ�
act=Request("act")
If act="save" Then 
l_name=trim(request.form("name"))
l_memo=trim(request.form("memo"))
l_time=now()

set rs=server.createobject("adodb.recordset")
sql="select * from web_ads_position"
rs.open(sql),cn,1,3
rs.addnew
rs("name")=l_name
rs("memo")=l_memo
rs("time")=l_time

rs.update
rs.close
set rs=nothing
response.Write "<script language='javascript'>alert('��ӳɹ���');location.href='ads_position_list.asp';</script>"
end if 
 %>
 

	<%
Call header()

%>

  <form id="form1" name="form1" method="post" action="?act=save">
         <script language='javascript'>
function checksignup1() {
if ( document.form1.name.value == '' ) {
window.alert('���������߿ͷ�����^_^');
document.form1.name.focus();
return false;}



return true;}
</script>
	<table cellpadding='3' cellspacing='1' border='0' class='tableBorder' align=center>
	<tr>
	  <th class='tableHeaderText' colspan=2 height=25>������߿ͷ�</th>
	<tr>
	<td width="15%" height=23 class='forumRowHighLight'>���߿ͷ����� (����) </td>
	<td width="85%" class='forumRowHighLight'><input name='name' type='text' id='name' size='70' maxlength="30">
	  &nbsp;</td>
	</tr><tr>
	  <td class='forumRow' height=11>���߿ͷ�����</td>
	  <td class='forumRow'><textarea name='memo'  cols="100" rows="6" id="memo" ></textarea></td>
	</tr>
	<tr><td height="50" colspan=2  class='forumRowHighLight'><div align="center">
	  <INPUT type=submit value='�ύ' onClick='javascript:return checksignup1()' name=Submit>
	  </div></td></tr>
	</table>
</form>

<%
Call DbconnEnd()
 %>
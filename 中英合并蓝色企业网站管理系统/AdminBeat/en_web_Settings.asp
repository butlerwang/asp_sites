<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="../inc/en_x_to_html/index_to_html.asp" -->
<%
act=Request("act")
If act="save" Then 
web_name=trim(request.form("web_name"))
web_url=trim(request.form("web_url"))
web_image=trim(request.form("web_image"))
web_title=trim(request.form("web_title"))
web_keywords=trim(request.form("web_keywords"))
web_description=trim(request.form("web_description"))
web_TopHTML=trim(request.form("web_TopHTML"))
web_BottomHTML=trim(request.form("web_BottomHTML"))
web_copyright=trim(request.form("content"))
web_contact=trim(request.form("web_contact"))
web_person=trim(request.form("web_person"))
web_birthdate=trim(request.form("web_birthdate"))
web_birthplace=trim(request.form("web_birthplace"))
web_shortintro=trim(request.form("web_shortintro"))
web_email=trim(request.form("web_email"))
web_tel=trim(request.form("web_tel"))
web_ModelEdit=trim(request.form("web_ModelEdit"))
web_time=trim(request.form("web_time"))
if web_time="" then
 web_time=now()
end if 

set rs=server.createobject("adodb.recordset")
sql="select * from en_web_settings"
rs.open(sql),cn,1,3
rs("web_name")=web_name
rs("web_url")=web_url
rs("web_image")=web_image
rs("web_title")=web_title
rs("web_keywords")=web_keywords
rs("web_description")=web_description
rs("web_TopHTML")=web_TopHTML
rs("web_BottomHTML")=web_BottomHTML
rs("web_copyright")=web_copyright
rs("web_contact")=web_contact
'rs("web_person")=web_person
'rs("web_birthdate")=web_birthdate
'rs("web_birthplace")=web_birthplace
'rs("web_shortintro")=web_shortintro
'rs("web_email")=web_email
rs("web_tel")=web_tel
rs("web_ModelEdit")=web_ModelEdit
rs("web_time")=web_time
rs.update
rs.close
set rs=nothing

call index_to_html()
response.Write "<script language='javascript'>alert('�޸ĳɹ���')</script>"

end if
 %>
 
	<%
Call header()

%>
 <%set rs=server.createobject("adodb.recordset")
sql="select * from en_web_settings"
rs.open(sql),cn,1,1
if not rs.eof and not rs.bof then
%>
  <form id="form1" name="form1" method="post" action="?act=save">
         <script language='javascript'>
function checksignup1() {
if ( document.form1.web_name.value == '' ) {
window.alert('��������վ����^_^');
document.form1.web_name.focus();
return false;}

return true;}
</script>

	<table cellpadding='3' cellspacing='1' border='0' class='tableBorder' align=center>
	<tr>
	  <th class='tableHeaderText' colspan=2 height=31>��վ��Ϣ����</th>
	<tr>
	  <td height=23 colspan="2" class='forumRow'><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td height="20" class='TipTitle'>&nbsp;�� ������ʾ</td>
        </tr>
        <tr>
          <td height="30" valign="top" class="TipWords"><p>1����վ��ʲô������ַ�Ƕ��٣���վ��LOGO�ȡ�������һһ���ðɡ�</p>
            <p>2���޸���ĳ����Ϣ��Ĭ��ֻ���Զ�������ҳ�棬����ҳ����Ҫ�ֶ����ɲŻῴ���鿴���Ч����</p></td>
        </tr>
        <tr>
          <td height="10">&nbsp;</td>
        </tr>
      </table></td>
	  </tr>
	<tr>
	<td width="15%" height=23 class='forumRowHighLight'>��վ����</td>
	<td class='forumRowHighLight'><input name='web_name' type='text' id='web_name' value="<%=rs("web_name")%>" size='40'></td>
	</tr>
	<tr>
	<td class='forumRow' height=23>��վ��ַ</td>
<td class='forumRow'><input type='text' id='web_url' name='web_url' value="<%=rs("web_url")%>" size='40'> 
  &nbsp;����http://��ͷ��<span style="color: #FF0000" >���治Ҫ�� / </span>���磺<a href="http://www.huiguer.com/" target="_blank">http://www.huiguer.com</a></td>
	</tr>
	  <tr>
	    <td class='forumRowHighLight' height=23>��վ��־(logo)</td>
	    <td width="85%" class='forumRowHighLight'><table width="100%" border="0" cellspacing="0" cellpadding="0">
         <tr>
           <td width="22%"  class='forumRow'><input name="web_image" type="text" id="web_image"  value="<%=rs("web_image")%>"  size="30"></td>
           <td width="78%"  class='forumRow'><iframe width="500" name="ad" frameborder=0 height=30 scrolling=no src=upload.asp></iframe></td>
         </tr>
       </table></td>
      </tr>
	    <td class='forumRow' height=23>��ҳ����(Title)</td>
	      <td class='forumRow'><input type='text' id='web_title' name='web_title'   value="<%=rs("web_title")%>" size='80'></td>
	</tr>
	    <td class='forumRowHighLight' height=11>��վ�ؼ���(keywords)</td>
	      <td class='forumRowHighLight'><input type='text' id='v3' name='web_keywords'   value="<%=rs("web_keywords")%>" size='80'>
	  &nbsp;���ԣ�����</td>
	</tr><tr>
	  <td class='forumRow' height=11>��վ����(Description)</td>
	  <td class='forumRow'><textarea name='web_description'  cols="100" rows="4" ><%=rs("web_description")%></textarea></td>
	</tr>
	 <tr>
	  <td class='forumRowHighLight' height=11>��ҳ��ϵ��ʽ</td>
	  <td class='forumRowHighLight'><textarea name='web_contact'  cols="100" rows="10" ><%=rs("web_contact")%></textarea></td>
	</tr>
	<tr>
	  <td class='forumRow' height=23>��վ�ײ�HTML����</td>
	  <td class='forumRow'> <textarea name='web_BottomHTML' cols="100" rows="10"  id="web_BottomHTML" ><%=rs("web_BottomHTML")%></textarea></td>
	</tr>	
	<tr>
	  <td class='forumRow' height=23>��ϵ�绰</td>
	  <td class='forumRow'><input type='text' id='v44' name='web_tel'  value="<%=rs("web_tel")%>" size='40'></td>
	</tr>
	<tr>
	  <td class='forumRowHighLight' height=23>��̨ģ�����</td>
	  <td class='forumRowHighLight'><label>
	       <input type="radio" name="web_ModelEdit" value="1"<%
		if rs("web_ModelEdit")=1 then
		response.write "checked"
		end if%>>
      ����
      &nbsp;
      <input name="web_ModelEdit" type="radio" value="0" <%if rs("web_ModelEdit")=0 then
		response.write "checked"
		end if%>>
      �ر�</label></td>
	  </tr>
	<tr>
	  <td class='forumRow' height=23>�޸�ʱ��</td>
	  <td class='forumRow'><input type='text' id='v45' name='web_time'  value="<%=rs("web_time")%>" size='40'> 
	  &nbsp;<a href="#" class="green" onClick="document.form1.web_time.value='<%=now()%>'">ͬ��������ʱ��</a>     </td>
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
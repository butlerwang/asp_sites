<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="page_next.asp" -->




<script language="JavaScript">
<!--
function ask(msg) {
	if( msg=='' ) {
		msg='���棺ɾ���󽫲��ɻָ����Ƿ�ȷ��ɾ����';
	}
	if (confirm(msg)) {
		return true;
	} else {
		return false;
	}
}
//-->
</script>
	<%
Call header()
%>

	<table cellpadding='3' cellspacing='1' border='0' class='tableBorder' align=center>
	<tr>
	  <th width="100%" height=25 class='tableHeaderText'>���߿ͷ��б�</th>
	
	<tr><td height="400" valign="top"  class='forumRow'><br>
	
	
		 <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="25" class='TipTitle'>&nbsp;�� ������ʾ</td>
          </tr>
          <tr>
            <td height="30" valign="top" class="TipWords">
		    <p>1������޸ģ������߿ͷ������滻Ϊ�Լ��ģ����ɡ�</p>
			<p>2�����ߴ����ȡ������</p>
			<p>��1��QQ���ߴ���������ַ��http://wp.qq.com/������̼ҹ�ͨ�����</p>
			<p>��2���������ߴ���������ַ��http://www.taobao.com/wangwang/2011_seller/wangbiantianxia/</p>
			<p>��3���������ߴ��룬�磺MSN��Skype�ȣ��뵽�ٷ����ɴ��롣</p>
			<p>3��<font color="#009900"><b>���ӡ��޸ġ�ɾ�����߿ͷ��󣬱�����������ȫվ��̬��������Ч��</b></font></p>
			<p>4���������Ҫ���߿ͷ�����ɾ�����У�Ȼ�������������о�̬���Ͳ�����ʾ���߿ͷ��������ˡ�</p>
			</td>
          </tr>
          <tr>
            <td height="10" ></td>
          </tr>
        </table><br />
		
	
	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="25" class='forumRowHighLight'>&nbsp;| <a href="ads_position_add.asp">������߿ͷ�</a></td>
          </tr>
      </table><br />
	   
	  
	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="2">
          <tr>
            <td width="5%" height="30" bgcolor="#A5C6FC" class="TitleHighlight"><div align="center" style="font-weight: bold; color: #FFFFFF">���</div></td>
            <td width="25%" height="30" bgcolor="#A5C6FC" class="TitleHighlight"><div align="center" style="font-weight: bold; color: #FFFFFF">���߿ͷ�����</div></td>
            <td width="50%" bgcolor="#A5C6FC" class="TitleHighlight"><div align="center" style="font-weight: bold; color: #FFFFFF">���߿ͷ�������ʾЧ��</div></td>
            <td width="20%" bgcolor="#A5C6FC" class="TitleHighlight"><div align="center" style="font-weight: bold; color: #FFFFFF">����</div></td>
          </tr>
<% '���߿ͷ��б�ģ��
strFileName="ads_position_list.asp" 
pageno=25
set rs = server.CreateObject("adodb.recordset")
s_sql="select * from web_ads_position order by id"
rs.Open (s_sql),cn,1,1
rscount=rs.recordcount
if not rs.eof and not rs.bof then
call showsql(pageno)
rs.move(rsno)
for p_i=1 to loopno
%>
<% if p_i mod 2 =0 then
class_style="forumRow"
else
class_style="forumRowHighLight"
end if%>
            <form name="form1" method="post" action="?action=edit&id=<%=rs("id")%>">
          <tr >
            <td   height="40" class='<%=class_style%>'><div align="center"><%=rs("id")%></div></td>
           <td class='<%=class_style%>' ><div align="center"><%=rs("name")%></div></td>
            <td class='<%=class_style%>' ><div align="center"><%=rs("memo")%></div></td>
            <td class='<%=class_style%>' >
            <div align="center"><a href="ads_position_edit.asp?id=<%=rs("id")%>&page=<%=page%>&act=<%=act%>&keywords=<%=keywords%>">�޸�</a> | <a href="javascript:if(ask('���棺ɾ���󽫲��ɻָ����Ƿ�ȷ��ɾ����')) location.href='ads_position_del.asp?id=<%=rs("id")%>&page=<%=page%>&act=<%=act%>&keywords=<%=keywords%>';">ɾ��</a>            </div></td>
          </tr></form>
		  		  <%
		  rs.movenext
		  next
		  else
response.write "<div align='center'><span style='color: #FF0000'>�������ݣ�</span></div>"
		  end if 
		  rs.close
		  set rs=nothing
		  %>
		    <tr  >
              <td height="35"  colspan="9" ><div align="center">
                <%call showpage(strFileName,rscount,pageno,false,true,"")%>
           </div></td>
		    </tr>
      </table>
	    <br></td>
	</tr>
	</table>

<%
Call DbconnEnd()
 %>
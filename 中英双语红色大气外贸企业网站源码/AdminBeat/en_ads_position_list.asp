<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="page_next.asp" -->




<script language="JavaScript">
<!--
function ask(msg) {
	if( msg=='' ) {
		msg='���棺ɾ���󽫲��ɻָ�������������벻�������';
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
	  <th width="100%" height=25 class='tableHeaderText'>���λ��</th>
	
	<tr><td height="400" valign="top"  class='forumRow'><br>
	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="25" class='forumRowHighLight'>&nbsp;| <a href="en_ads_position_add.asp">����µĹ��λ��</a></td>
          </tr>
          <tr>
            <td height="30"></td>
          </tr>
      </table>
	   
	  
	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="2">
          <tr>
            <td width="9%" height="30" bgcolor="#A5C6FC" class="TitleHighlight"><div align="center" style="font-weight: bold; color: #FFFFFF">���</div></td>
            <td width="29%" height="30" bgcolor="#A5C6FC" class="TitleHighlight"><div align="center" style="font-weight: bold; color: #FFFFFF">���λ��</div></td>
            <td width="20%" bgcolor="#A5C6FC" class="TitleHighlight"><div align="center" style="font-weight: bold; color: #FFFFFF">��ע</div></td>
            <td width="22%" bgcolor="#A5C6FC" class="TitleHighlight"><div align="center" style="font-weight: bold; color: #FFFFFF">���ʱ��</div></td>
            <td width="20%" bgcolor="#A5C6FC" class="TitleHighlight"><div align="center" style="font-weight: bold; color: #FFFFFF">����</div></td>
          </tr>
<% '�����б�ģ��
strFileName="en_ads_position_list.asp" 
pageno=25
set rs = server.CreateObject("adodb.recordset")
s_sql="select * from en_web_ads_position order by  time asc"
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
            <td class='<%=class_style%>' ><div align="center"><%=rs("time")%></div></td>
            <td class='<%=class_style%>' >
            <div align="center"><a href="en_ads_position_edit.asp?id=<%=rs("id")%>&page=<%=page%>&act=<%=act%>&keywords=<%=keywords%>">�޸�</a> | <a href="javascript:if(ask('���棺ɾ���󽫲��ɻָ�������������벻�������')) location.href='en_ads_position_del.asp?id=<%=rs("id")%>&page=<%=page%>&act=<%=act%>&keywords=<%=keywords%>';">ɾ��</a>            </div></td>
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
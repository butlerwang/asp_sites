<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="../inc/x_to_html/index_to_html.asp" -->
<!-- #include file="page_next.asp" -->

<% '����ģ��
act=request.querystring("act")
keywords=trim(request.form("keywords"))
cid=request("cid")


if act="search" then


if cid=""  then
s_sql="select * from web_ads where [name]  like '%"&keywords&"%'  order by [position]"
else
search_sql="and [position]="&cid&""
s_sql="select * from web_ads where [name] like '%"&keywords&"%'"&search_sql&"  order by [order]"
end if

else
s_sql="select * from web_ads order by [position]"

end if

%>

<% '�޸�˳��ģ��
action1=request.querystring("action")
id1=request.querystring("id")
order1=trim(request.form("order"))
if action1="edit" then
if isnumeric(order1)=false then
response.Write "<script language='javascript'>alert('������Ĳ������֣�');location.href='?page="&page&"&act="&act&"&keywords="&keywords&"';</script>"
else
set rs1=server.createobject("adodb.recordset")
sql="select id,order from web_ads where id="&id1&""
rs1.open(sql),cn,1,3
rs1("order")=cint(order1)
rs1.update
rs1.close
set rs1=nothing
call index_to_html()
end if
end if

%>
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
	  <th width="100%" height=25 class='tableHeaderText'>����б�</th>
	
	<tr><td height="400" valign="top"  class='forumRow'><br>
	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="25" class='TipTitle'>&nbsp;�� ������ʾ</td>
          </tr>
          <tr>
            <td height="30" valign="top" class="TipWords"><p>1���������еĹ����ʽ�������ں�̨���и��¡�Ŀǰ�������ֹ�桢ͼƬ��桢Flash���͹�����˴����档</p>
                <p>2����������˴������⣬���������ʽ��������JS�ļ���ʽ��������ҳ�С�</p>
                <p>3����������˴������⣬�������ĸ��²���Ҫ���������ҳ���ɿ����޸ĵ�Ч����</p>
                <p>4������ļ�������ڸ�Ŀ¼�µ�ADs�ļ����С�</p></td>
          </tr>
          <tr>
            <td height="10" ></td>
          </tr>
        </table>
	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="25" class='forumRowHighLight'>&nbsp;| <a href="ads_add.asp">����µĹ��</a></td>
          </tr>
          <tr>
            <td height="30"></td>
          </tr>
      </table>
	   
	  
	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="2">
          <tr>
            <td width="4%" height="30" bgcolor="#A5C6FC" class="TitleHighlight"><div align="center" style="font-weight: bold; color: #FFFFFF">���</div></td>
            <td width="24%" height="30" bgcolor="#A5C6FC" class="TitleHighlight"><div align="center" style="font-weight: bold; color: #FFFFFF">������</div></td>
            <td width="14%" bgcolor="#A5C6FC" class="TitleHighlight"><div align="center" style="font-weight: bold; color: #FFFFFF">�������</div></td>
            <td width="14%" bgcolor="#A5C6FC" class="TitleHighlight"><div align="center" style="font-weight: bold; color: #FFFFFF">���λ��</div></td>
            <td width="10%" bgcolor="#A5C6FC" class="TitleHighlight"><div align="center" style="font-weight: bold; color: #FFFFFF">��ʾ����</div></td>
            <td width="7%" bgcolor="#A5C6FC" class="TitleHighlight"><div align="center" style="font-weight: bold; color: #FFFFFF">���</div></td>
            <td width="18%" bgcolor="#A5C6FC" class="TitleHighlight"><div align="center" style="font-weight: bold; color: #FFFFFF">���ʱ��</div></td>
            <td width="9%" bgcolor="#A5C6FC" class="TitleHighlight"><div align="center" style="font-weight: bold; color: #FFFFFF">����</div></td>
          </tr>
<% '�����б�ģ��
strFileName="ads_list.asp" 
pageno=20
set rs = server.CreateObject("adodb.recordset")
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
           <td class='<%=class_style%>' ><%=rs("name")%></td>
            <td class='<%=class_style%>' ><div align="center">
              <%
			select case rs("ADtype")
			case 1
			response.write "���ֹ��"
			case 2
			response.write "ͼƬ���"
			case 3
			response.write "Flash���"
			case 4
			response.write "������"
			end select%>
            </div></td>

            <td class='<%=class_style%>' ><div align="center"><%
			set rst=server.createobject("adodb.recordset")
			sql="select name from web_ads_position where [id]="&rs("position")&""
			rst.open(sql),cn,1,1
			if not rst.eof and not rst.bof then
			response.write rst("name")
			end if
			rst.close
			set rst=nothing
			%></div></td>
            <td class='<%=class_style%>' > <div align="center"><label>
            <input name="order" type="text" value="<%=rs("order")%>" size="5">
            <input type="submit" name="Submit" value="�޸�"  >
            </label>
           </div></td>
            <td class='<%=class_style%>' ><div align="center"><a href="ads_view_yes.asp?id=<%=rs("id")%>&page=<%=page%>&act=<%=act%>&keywords=<%=keywords%>"><%if rs("view_yes")=1 then%>�����<%else%><span style="color: #FF0000">δ���</span><% end if%></a></div></td>
            <td class='<%=class_style%>' ><div align="center"><%=rs("time")%></div></td>
            <td class='<%=class_style%>' >
            <div align="center"><a href="ads_edit.asp?id=<%=rs("id")%>&page=<%=page%>&act=<%=act%>&keywords=<%=keywords%>">�޸�</a> | <a href="javascript:if(ask('���棺ɾ���󽫲��ɻָ�������������벻�������')) location.href='ads_del.asp?id=<%=rs("id")%>&page=<%=page%>&act=<%=act%>&keywords=<%=keywords%>';">ɾ��</a>            </div></td>
          </tr></form>
		  		  <%
		  rs.movenext
		  next
		  else
response.write "<div align='center'><span style='color: #FF0000'>�������ӣ�</span></div>"
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
	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="20" class='forumRow'>&nbsp;</td>
          </tr>
          <tr>
            <td height="25" class='forumRowHighLight'>&nbsp;| �������</td>
          </tr>
          <tr>
            <td height="70"><form name="form1" method="post" action="?act=search"><div align="center">
<%
Dim count1,rsClass1,sqlClass1
set rsClass1=server.createobject("adodb.recordset")
sqlClass1="select id,name from web_ads_position  order by id" 
rsClass1.open sqlClass1,cn,1,1
%>
            <select name="cid" id="cid" onChange="changeselect1(this.value)">
              <option value="">ѡ�����</option>
              <%
count1 = 0
do while not rsClass1.eof
response.write"<option value="&rsClass1("ID")&">"&rsClass1("Name")&"</option>"
count1 = count1 + 1
rsClass1.movenext
loop
rsClass1.close
%>
            </select>
            <label>
<input name="keywords" type="text"  size="35" maxlength="40">
              </label>
                <label>
                       &nbsp;
                       <input type="submit" name="Submit" value="�� ��">
                </label>
              </div>
            </form>
            </td>
          </tr>
      </table>
	    <br></td>
	</tr>
	</table>

<%
Call DbconnEnd()
 %>
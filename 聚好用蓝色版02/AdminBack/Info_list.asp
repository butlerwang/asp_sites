<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="page_next.asp" -->

<%
'��Ƹְλ�ļ��л�ȡ
set rs_1=server.createobject("adodb.recordset")
sql="select FileName,FolderName from web_Models_type where [id]=26"
rs_1.open(sql),cn,1,1
if not rs_1.eof then
Model_FileName=rs_1("FileName")
if rs_1("FolderName")<>"" then
Model_FolderName="/"&rs_1("FolderName")
end if
end if
rs_1.close
set rs_1=nothing%>

<% '����ģ��
act=request.querystring("act")
keywords=trim(request.form("keywords"))
if act="search" then
cid=request("cid")
pid=request("pid")
ppid=request("ppid")

if cid="" and pid="" and  ppid="" then
s_sql="select id,title,cid,pid,ppid,file_path,image,index_push,view_yes,hit,ip,time,AuthorID from web_info where [title] like '%"&keywords&"%'  order by time desc"
elseif pid="" and ppid="" then
search_sql="and cid='"&cid&"'"
s_sql="select id,title,cid,pid,ppid,file_path,image,index_push,view_yes,hit,ip,time,AuthorID from web_info where [title] like '%"&keywords&"%'"&search_sql&"   order by time desc"
elseif ppid="" then
search_sql="and pid='"&pid&"'"
s_sql="select id,title,cid,pid,ppid,file_path,image,index_push,view_yes,hit,ip,time,AuthorID from web_info where [title] like '%"&keywords&"%'"&search_sql&"   order by time desc"
else
search_sql="and ppid='"&ppid&"'"
s_sql="select id,title,cid,pid,ppid,file_path,image,index_push,view_yes,hit,ip,time,AuthorID from web_info where [title] like '%"&keywords&"%'"&search_sql&"   order by time desc"
end if
else
s_sql="select id,title,cid,pid,ppid,file_path,image,index_push,view_yes,hit,ip,time,AuthorID from web_info   order by time desc"

end if

%>
<script language="JavaScript">
<!--
function ask(msg) {
	if( msg=='' ) {
		msg='���棺ɾ���󽫲��ɻָ���ȷ��Ҫɾ����';
	}
	if (confirm(msg)) {
		return true;
	} else {
		return false;
	}
}
//-->
</script>
<script language="javascript">

//ȫѡJS
function unselectall(){
if(document.form2.chkAll.checked){
document.form2.chkAll.checked = document.form2.chkAll.checked&0;
}
}
function CheckAll(form){
for (var i=0;i<form.elements.length;i++){
var e = form.elements[i];
if (e.Name != 'chkAll'&&e.disabled==false)
e.checked = form.chkAll.checked;
}
}
</script>	<%
Call header()
%>
<%
if cid>1 then
juhaoyong_cid=cid
juhaoyong_pid=-1
juhaoyong_ppid=-1
end if

if pid>1 then
juhaoyong_cid=juhaoyongGetCategoryParentId(pid)
juhaoyong_pid=pid
juhaoyong_ppid=-1
end if

if ppid>1 then
juhaoyong_pid=juhaoyongGetCategoryParentId(ppid)
juhaoyong_cid=juhaoyongGetCategoryParentId(juhaoyong_pid)
juhaoyong_ppid=ppid
end if

Function juhaoyongGetCategoryParentId(id)
set juhaoyongRs=server.createobject("adodb.recordset")
juhaoyongSql="select id,pid,ppid,name from category where id="&id 
juhaoyongRs.open juhaoyongSql,cn,1,1

juhaoyongGetCategoryParentId=juhaoyongRs("pid")

juhaoyongRs.close
set juhaoyongRs=nothing
End Function
%>
	<table cellpadding='3' cellspacing='1' border='0' class='tableBorder' align=center>
	<tr>
	  <th width="100%" height=25 class='tableHeaderText'>ְλ�б�</th>
	
	<tr><td height="400" valign="top"  class='forumRow'><br>
	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="25" class='TipTitle'>&nbsp;�� ������ʾ</td>
          </tr>
          <tr>
            <td height="30" valign="top" class="TipWords"><p>1��ְλ�б���ʾ������ӵ�����ְλ����ʾ��δ��ˡ���ְλ����������վ����ʾ��</p>
                <p>2��ɾ��ְλ����ͬ��ɾ�����ݿ��еļ�¼��ְλ�ľ����ַ�����ء�</p>
            </td>
          </tr>
          <tr>
            <td height="10" ></td>
          </tr>
        </table>
	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="25" class='forumRowHighLight'>&nbsp;| <a href="info_add.asp?juhaoyong_cid=<%=juhaoyong_cid%>&juhaoyong_pid=<%=juhaoyong_pid%>&juhaoyong_ppid=<%=juhaoyong_ppid%>">����µ�ְλ</a></td>
          </tr>
          
      </table>
 <form name="form2" method="post" action="info_Del.asp?action=AllDel&juhaoyong_cid=<%=juhaoyong_cid%>&juhaoyong_pid=<%=juhaoyong_pid%>&juhaoyong_ppid=<%=juhaoyong_ppid%>&page=<%=page%>&act=<%=act%>&keywords=<%=keywords%>">
 	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="2">
          <tr>
            <td width="2%" height="30" class="TitleHighlight">&nbsp;</td>
            <td width="4%" class="TitleHighlight"><div align="center" style="font-weight: bold;color:#ffffff">���</div></td>
            <td width="31%" height="30" class="TitleHighlight"><div align="center" style="font-weight: bold;color:#ffffff">ְλ����</div></td>
			<td width="31%" height="30" class="TitleHighlight"><div align="center" style="font-weight: bold;color:#ffffff">��Ƹ����</div></td>
            <td width="6%" class="TitleHighlight"><div align="center" style="font-weight: bold;color:#ffffff">���</div></td>
            <td width="17%" class="TitleHighlight"><div align="center" style="font-weight: bold;color:#ffffff">���ʱ��</div></td>
            <td width="8%" class="TitleHighlight"><div align="center" style="font-weight: bold;color:#ffffff">ְλ����</div></td>
          </tr>
<% 'ְλ�б�ģ��
strFileName="info_list.asp" 
pageno=25
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
<%

%>
          <tr >
            <td   height="30" class='<%=class_style%>'><div align="center"><input type="checkbox" name="Selectitem" id="Selectitem" value="<%=rs("id")%>"></div></td>
            <td   height="30" class='<%=class_style%>'><div align="center"><%=rs("id")%></div></td>
            <td class='<%=class_style%>' >&nbsp;<a href="info_edit.asp?id=<%=rs("id")%>&page=<%=page%>&act=<%=act%>&keywords=<%=keywords%>" ><%=left(rs("title"),16)%></a></td>
            
			<td class='<%=class_style%>' >&nbsp;
			<% '������ʾ
			cid=cint(rs("cid"))

			set rs1=server.createobject("adodb.recordset")
			sql="select name from category where id="&cid&""
			rs1.open(sql),cn,1,1
			if not rs1.eof and not rs1.bof then
			response.write rs1("name")
			response.write "&nbsp;>&nbsp;"
			end if
			rs1.close
			set rs1=nothing
			
			if rs("pid")<>"" then
            pid=cint(rs("pid"))
						set rs1=server.createobject("adodb.recordset")
			sql="select name from category where id="&pid&""
			rs1.open(sql),cn,1,1
			if not rs1.eof and not rs1.bof then
			response.write rs1("name")
			response.write "&nbsp;>&nbsp;"
			end if
			rs1.close
			set rs1=nothing
			end if
			
			if rs("ppid")<>"" then
            ppid=cint(rs("ppid"))
						set rs1=server.createobject("adodb.recordset")
			sql="select name from category where id="&ppid&""
			rs1.open(sql),cn,1,1
			if not rs1.eof and not rs1.bof then
			response.write rs1("name")
			end if
			rs1.close
			set rs1=nothing
			end if
			%>            </td>
			
			<td class='<%=class_style%>' ><div align="center"><a href="info_view_yes.asp?id=<%=rs("id")%>&juhaoyong_cid=<%=juhaoyong_cid%>&juhaoyong_pid=<%=juhaoyong_pid%>&juhaoyong_ppid=<%=juhaoyong_ppid%>&page=<%=page%>&act=<%=act%>&keywords=<%=keywords%>"><%if rs("view_yes")=1 then%>�����<%else%><span style="color: #FF0000">δ���</span><% end if%></a></div></td>
            <td class='<%=class_style%>' ><div align="center"><%=rs("time")%></div></td>
            <td class='<%=class_style%>' >
            <div align="center"><a href="info_edit.asp?id=<%=rs("id")%>&juhaoyong_cid=<%=juhaoyong_cid%>&juhaoyong_pid=<%=juhaoyong_pid%>&juhaoyong_ppid=<%=juhaoyong_ppid%>&page=<%=page%>&act=<%=act%>&keywords=<%=keywords%>">�޸�</a> | <a href="javascript:if(ask('���棺ɾ���󽫲��ɻָ���ȷ��Ҫɾ����')) location.href='info_del.asp?id=<%=rs("id")%>&juhaoyong_cid=<%=juhaoyong_cid%>&juhaoyong_pid=<%=juhaoyong_pid%>&juhaoyong_ppid=<%=juhaoyong_ppid%>&page=<%=page%>&act=<%=act%>&keywords=<%=keywords%>';">ɾ��</a>            </div></td>
          </tr>
		  		  <%
		  rs.movenext
		  next
		  else
response.write "<div align='center'><span style='color: #FF0000'>����ְλ��</span></div>"
		  end if 
		  rs.close
		  set rs=nothing
		  %>
		          <tr  >
		            <td height="35"  colspan="9" >&nbsp;<input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'>
                    ȫѡ/ȫ��ѡ&nbsp;<input type="submit" name="Submit" value="ɾ��ѡ��"></td>
          </tr>
            <tr  >
              <td height="35"  colspan="9" ><div align="center">
                <%call showpage(strFileName,rscount,pageno,false,true,"")%>
           </div></td>
		    </tr>
      </table> 
 </form>  
	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="20" class='forumRow'>&nbsp;</td>
          </tr>
          <tr>
            <td height="25" class='forumRowHighLight'>&nbsp;| ְλ����</td>
          </tr>
          <tr>
            <td height="70"><form name="form1" method="post" action="?act=search">
              <div align="center">
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
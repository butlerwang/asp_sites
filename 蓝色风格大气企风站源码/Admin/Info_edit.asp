<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="../inc/rand.asp" -->
<!-- #include file="../inc/x_to_html/index_to_html.asp" -->
<!-- #include file="../inc/x_to_html/Recruit_list_to_html.asp" -->

<% '更新数据到数据表
page=request.querystring("page")
act=request.querystring("act")
keywords=request.querystring("keywords")
article_id=cint(request.querystring("id"))

act1=Request("act1")
If act1="save" Then 
a_id=cint(request.form("a_id"))
a_title=request.form("a_title")
a_author=request.form("a_author")
a_cid=trim(request.form("cid"))
a_pid=trim(request.form("pid"))
a_ppid=trim(request.form("ppid"))
a_image=trim(request.form("web_image"))
a_keywords=trim(request.form("a_keywords"))
a_description=trim(request.form("a_description"))
a_content=request.form("a_content")
a_person=request.form("a_person")
a_address=request.form("a_address")
a_tel=request.form("a_tel")
a_email=request.form("a_email")
a_qq=request.form("a_qq")
a_hit=trim(request.form("a_hit"))
a_index_push=trim(request.form("a_index_push"))


set rs=server.createobject("adodb.recordset")
sql="select * from [web_info] where id="&a_id&""
rs.open(sql),cn,1,3
rs("title")=a_title
rs("AuthorID")=a_author
rs("cid")=a_cid
rs("pid")=a_pid
rs("ppid")=a_ppid
rs("image")=a_image
rs("keywords")=a_keywords
rs("description")=a_description
rs("content")=a_content
rs("person")=a_person
rs("address")=a_address
rs("tel")=a_tel
rs("email")=a_email
rs("qq")=a_qq
'rs("hit")=a_hit
'rs("index_push")=a_index_push
rs("edit_time")=now()
rs.update
rs.close
set rs=nothing
%>
<% 
ClassID=a_cid
call Recruit_list_to_html(ClassID)
%>
<%
response.Write "<script language='javascript'>alert('修改成功！');location.href='info_list.asp?page="&page&"&act="&act&"&keywords="&keywords&"';</script>"
end if 
%>
 <script type="text/javascript" charset="utf-8" src="../KKKKKKKEditor/kindeditor.js"></script>
<script type="text/javascript" src="../KKKKKKKEditor/editor.js"></script>
 <!-- 三级联动菜单 开始 -->
<script language="JavaScript">
<!--
<%
'二级数据保存到数组
Dim count2,rsClass2,sqlClass2
set rsClass2=server.createobject("adodb.recordset")
sqlClass2="select id,pid,ppid,name from [category] where ppid=2 order by id " 
rsClass2.open sqlClass2,cn,1,1
%>
var subval2 = new Array();
//数组结构：一级根值,二级根值,二级显示值
<%
count2 = 0
do while not rsClass2.eof
%>
subval2[<%=count2%>] = new Array('<%=rsClass2("pID")%>','<%=rsClass2("ID")%>','<%=rsClass2("Name")%>')
<%
count2 = count2 + 1
rsClass2.movenext
loop
rsClass2.close
%>

<%
'三级数据保存到数组
Dim count3,rsClass3,sqlClass3
set rsClass3=server.createobject("adodb.recordset")
sqlClass3="select id,pid,ppid,name from [category] where ppid=3 order by id" 
rsClass3.open sqlClass3,cn,1,1
%>
var subval3 = new Array();
//数组结构：二级根值,三级根值,三级显示值
<%
count3 = 0
do while not rsClass3.eof
%>
subval3[<%=count3%>] = new Array('<%=rsClass3("pID")%>','<%=rsClass3("ID")%>','<%=rsClass3("Name")%>')
<%
count3 = count3 + 1
rsClass3.movenext
loop
rsClass3.close
%>

function changeselect1(locationid)
{
    document.form1.pid.length = 0;
    document.form1.pid.options[0] = new Option('选择二级分类','');
    document.form1.ppid.length = 0;
    document.form1.ppid.options[0] = new Option('选择三级分类','');
    for (i=0; i<subval2.length; i++)
    {
        if (subval2[i][0] == locationid)
        {document.form1.pid.options[document.form1.pid.length] = new Option(subval2[i][2],subval2[i][1]);}
    }
}

function changeselect2(locationid)
{
    document.form1.ppid.length = 0;
    document.form1.ppid.options[0] = new Option('选择三级分类','');
    for (i=0; i<subval3.length; i++)
    {
        if (subval3[i][0] == locationid)
        {document.form1.ppid.options[document.form1.ppid.length] = new Option(subval3[i][2],subval3[i][1]);}
    }
}
//-->
</script><!-- 三级联动菜单 结束 -->
	<%
Call header()

%>

<%
			set rs=server.createobject("adodb.recordset")
sql="select * from [web_info] where id="&article_id&""
rs.open(sql),cn,1,1
if not rs.eof and not rs.bof then
%>  <form id="form1" name="form1" method="post" action="?act1=save&page=<%=page%>&act=<%=act%>&keywords=<%=keywords%>">
         <script language='javascript'>
function checksignup1() {
if ( document.form1.a_title.value == '' ) {
window.alert('请输入标题^_^');
document.form1.a_title.focus();
return false;}

if ( document.form1.cid.value == '' ) {
window.alert('请选择分类^_^');
document.form1.cid.focus();
return false;}

return true;}
</script>
	<table cellpadding='3' cellspacing='1' border='0' class='tableBorder' align=center>
	<tr>
	  <th class='tableHeaderText' colspan=2 height=25>修改招聘职位</th>
	<tr>
	<td class='forumRowHighLight' height=23>分类<span class="forumRow"> (必选) </span></td>
    <td class='forumRowHighLight'><%
set rsc1=server.createobject("adodb.recordset")
sqlClass1="select id,pid,ppid,name from category where ppid=1 and ClassType=4 order by id" 
rsc1.open sqlClass1,cn,1,1
%>
            <select name="cid" id="cid" onChange="changeselect1(this.value)">
              <% '输出一级分类，并选定当前分类
count1 = 0
do while not rsc1.eof
%><option value="<%=rsc1("ID")%>"  <%if cint(rs("cid"))=rsc1("id") then
response.write "selected"
end if%>><%=rsc1("Name")%></option>
<%count1 = count1 + 1
rsc1.movenext
loop
rsc1.close
%>
            </select>
            &nbsp;&nbsp;
	
            <select name="pid" id="pid"  onchange="changeselect2(this.value)">
              <option value="">选择二级分类</option>
			 		<%'输出二级分类，并选定当前分类
			
set rsc2=server.createobject("adodb.recordset")
sqlClass2="select id,pid,ppid,name from category where ppid=2 and ClassType=4 and pid="&cint(rs("cid"))&" order by id" 
rsc2.open sqlClass2,cn,1,1

count1 = 0
do while not rsc2.eof
%><option value="<%=rsc2("ID")%>"  <%
if rs("pid")<>"" then
if cint(rs("pid"))=rsc2("id") then
response.write "selected"
end if
end if%>><%=rsc2("Name")%></option>
<%count1 = count1 + 1
rsc2.movenext
loop
rsc2.close

%>
            </select>
            &nbsp;&nbsp;
				
            <select name="ppid" id="ppid">
              <option value="">选择三级分类</option>
			  			  		<% '输出三级分类，并选定当前分类
			if rs("pid")<>"" then				
set rsc3=server.createobject("adodb.recordset")
sqlClass3="select id,pid,ppid,name from category where ppid=3 and ClassType=4 and pid="&cint(rs("pid"))&" order by id" 
rsc3.open sqlClass3,cn,1,1

count1 = 0
do while not rsc3.eof
%><option value="<%=rsc3("ID")%>"  <%
	if rs("ppid")<>"" then
if cint(rs("ppid"))=rsc3("id") then
response.write "selected"
end if
end if
%>><%=rsc3("Name")%></option>
<%count1 = count1 + 1
rsc3.movenext
loop
rsc3.close
end if
%>
            </select>&nbsp;</td>
	</tr>	
	<tr>
	<td width="15%" height=23 class='forumRow'>职位名称 (必填) </td>
	<td class='forumRow'><input name='a_title' type='text' id='a_title' value="<%=rs("title")%>" size='70'>
	<input name='a_id' type='hidden' id='a_id' value="<%=rs("id")%>" size='70'>
	  &nbsp;</td>
	</tr>

	<tr>
	<td width="15%" height=23 class='forumRowHighLight'>工作地点</td>
	<td class='forumRowHighLight'><input name='a_address' type='text' id='a_address'  value="<%=rs("address")%>" size='40'>
	  &nbsp;</td>
	</tr>
	<tr>
	<td width="15%" height=23 class='forumRow'>工资待遇</td>
	<td class='forumRow'><input name='a_person' type='text' id='a_person' value="<%=rs("person")%>"  size='40'>
	  &nbsp;</td>
	</tr>
	<tr>
	<td width="15%" height=23 class='forumRowHighLight'>招聘人数</td>
	<td class='forumRowHighLight'><input name='a_tel' type='text' id='a_tel' value="<%=rs("tel")%>"  size='40'>
	  &nbsp;</td>
	</tr>
	<tr>
	<td width="15%" height=23 class='forumRow'>性别要求</td>
	<td class='forumRow'><input name='a_email' type='text' id='a_email'  value="<%=rs("email")%>" size='40'>
	  &nbsp;</td>
	</tr>
	<tr>
	<td width="15%" height=23 class='forumRowHighLight'>年龄要求</td>
	<td class='forumRowHighLight'><input name='a_qq' type='text' id='a_qq' value="<%=rs("qq")%>" size='40'>
	  &nbsp;</td>
	</tr>					
<tr>
	  <td class='forumRow' height=11>条件要求 </td>
	  <td class='forumRow'><textarea name='a_description'  cols="100" rows="4" id="a_description" ><%=rs("description")%></textarea></td>
	</tr>
	<tr>
	  <td class='forumRowHighLight' height=23>职位描述 </td>
	  <td class='forumRowHighLight'> <textarea name="a_content" id="a_content" style=" width:100%; height:400px; visibility:hidden;"><%=rs("content")%></textarea></td>
	</tr>

	  
	<tr><td height="50" colspan=2  class='forumRow'><div align="center">
	  <INPUT type=submit value='提交' onClick='javascript:return checksignup1()' name=Submit>
	  </div></td></tr>
	</table>
</form>
<%
else
response.write"未找到数据"
end if%>
<%
Call DbconnEnd()
 %>
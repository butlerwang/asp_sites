
<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="../inc/rand.asp" -->
<!-- #include file="../inc/x_to_html/index_to_html.asp" -->
<!-- #include file="../inc/x_to_html/cases_to_html.asp" -->
<!-- #include file="../inc/x_to_html/cases_list_to_html.asp" -->

<% '添加数据到数据表
act=Request("act")
If act="save" and request.form("cid")<>"选择一级分类" Then 
a_title=request.form("a_title")
a_cid=trim(request.form("cid"))
a_pid=trim(request.form("pid"))
a_ppid=trim(request.form("ppid"))
a_url=trim(request.form("a_url"))
a_image=trim(request.form("web_image"))
a_Files=trim(request.form("Files"))
a_keywords=trim(request.form("a_keywords"))
a_description=trim(request.form("a_description"))
a_content=request.form("a_content")
a_from_name=trim(request.form("a_from_name"))
a_from_url=trim(request.form("a_from_url"))
a_author=trim(request.form("a_author"))
a_hit=trim(request.form("a_hit"))
a_index_push=trim(request.form("a_index_push"))
a_keywords_yes=trim(request.form("a_keywords_yes"))
a_slide_yes=cint(request.form("slide_yes"))
a_special_yes=cint(request.form("special_yes"))
a_OrderCount=trim(request.form("a_OrderCount"))
a_time=now()

'替换关键字
if a_keywords_yes=1 then
set rs=server.createobject("adodb.recordset")
sql="select [name],[url] from web_keywords order by [time] desc"
rs.open(sql),cn,1,1
if not rs.eof  then
do while not rs.eof 
a_content=replace(a_content,rs("name"),"<a href='"&rs("url")&"' target='_blank'>"&rs("name")&"</a>")
rs.movenext
loop
end if
rs.close
set rs=nothing
end if

set rs=server.createobject("adodb.recordset")
sql="select * from article"
rs.open(sql),cn,1,3
rs.addnew
rs("title")=a_title
rs("ArticleType")=3
rs("cid")=a_cid
rs("pid")=a_pid
rs("ppid")=a_ppid
rs("url")=a_url
rs("image")=a_image
rs("Files")=a_Files
rs("keywords")=a_keywords
rs("description")=a_description
rs("content")=a_content
rs("from_name")=a_from_name
rs("from_url")=a_from_url
rs("author")=a_author
rs("hit")=a_hit
rs("index_push")=a_index_push
rs("OrderCount")=a_OrderCount
rs("Pics")=trim(request.form("Pics"))
'rs("slide_yes")=a_slide_yes
'rs("special_yes")=a_special_yes
rs("time")=a_time
rs("edit_time")=a_time
rs("File_Path")=a7&minute(now)&second(now)&".html"
rs.update
rs.close
set rs=nothing
%>

<% '生成首页
call index_to_html()
%>
<% '生成案例静态页,列表页
set rs2=server.createobject("adodb.recordset")
sql="select top 1 [cid],[id] from [article] where [title]='"&a_title&"' and ArticleType=3 order by [time] desc"
rs2.open(sql),cn,1,1
if not rs2.eof  then
a_id=rs2("id")
ClassID=rs2("cid")
call cases_to_html(a_id)
call cases_list_to_html(ClassID)
end if
%>
<%
response.Write "<script language='javascript'>alert('添加成功');location.href='case_list.asp';</script>"
end if 
 %>
 <script type="text/javascript" src="PicUpload2/init.js"></script>
	<script charset="utf-8" src="Keditor/kindeditor.js"></script>
	<script charset="utf-8" src="Keditor/lang/zh_CN.js"></script>
	<script charset="utf-8" src="Keditor/editor.js"></script>
<script language="JavaScript" type="text/javascript">
  function show(){

var obj = document.getElementById("position");

var index = obj.selectedIndex;

var text =  obj.options[index].text;

var value = obj.options[index].value;

document.form1.a_from_name.value=text;
document.form1.a_from_url.value=document.form1.position.value;
  }
  </script>


 <!-- 三级联动菜单 开始 -->
<script language="JavaScript">
<!--
<%
'二级数据保存到数组
Dim count2,rsClass2,sqlClass2
set rsClass2=server.createobject("adodb.recordset")
sqlClass2="select id,pid,ppid,name from category where ppid=2 and ClassType=3 order by id " 
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
sqlClass3="select id,pid,ppid,name from category where ppid=3  and ClassType=3 order by id" 
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

  <form id="form1" name="form1" method="post" action="?act=save">
         <script language='javascript'>
function checksignup1() {
if ( document.form1.a_title.value == '' ) {
window.alert('请输入案例标题^_^');
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
	  <th class='tableHeaderText' colspan=2 height=25>添加案例</th>
	<tr>
	  <td height=23 colspan="2" class='forumRow'><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td height="20" class='TipTitle'>&nbsp;√ 操作提示</td>
        </tr>
        <tr>
          <td height="30" valign="top" class="TipWords"><p>1、本系统采用手动分页方法，如果想要对案例进行分页，只需在编辑器按钮栏点击<img src="images/inserthorizontalrule.gif" width="20" height="20" />图标，就会在编辑器内自动插入一条线做为分页标志，你不需要删除它，它将不会在案例中显示。</p></td>
        </tr>
        <tr>
          <td height="10">&nbsp;</td>
        </tr>
      </table></td>
	  </tr>
	<tr>
	<td width="15%" height=23 class='forumRow'>案例标题 (必填) </td>
	<td class='forumRow'><input name='a_title' type='text' id='a_title' size='70'>
	  &nbsp;</td>
	</tr>
	<tr>
	<td class='forumRowHighLight' height=23>案例分类<span class="forumRow"> (必选) </span></td>
    <td class='forumRowHighLight'><%
Dim count1,rsClass1,sqlClass1
set rsClass1=server.createobject("adodb.recordset")
sqlClass1="select id,pid,ppid,name from category where ppid=1 and ClassType=3 order by id" 
rsClass1.open sqlClass1,cn,1,1
%>
            <select name="cid" id="cid" onChange="changeselect1(this.value)">
              <option>选择一级分类</option>
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
            &nbsp;&nbsp;
            <select name="pid" id="pid"  onchange="changeselect2(this.value)">
              <option value="">选择二级分类</option>
            </select>
            &nbsp;&nbsp;
            <select name="ppid" id="ppid">
              <option value="">选择三级分类</option>
            </select>&nbsp;</td>
	</tr>
 
	  <tr>
	    <td class='forumRowHighLight' height=23>封面图片<br>大小：175 x 131</td>
	    <td width="85%" class='forumRowHighLight'><table width="100%" border="0" cellspacing="0" cellpadding="0">
         <tr>
           <td width="22%" ><input name="web_image" type="text" id="web_image"  size="30"></td>
           <td width="78%"  ><iframe width="500" name="ad" frameborder=0 height=30 scrolling=no src=upload.asp></iframe></td>
         </tr>
       </table></td>
      </tr>
	  <tr>
	    <td class='forumRowHighLight' height=23>批量上传图片<br>大小：800 x 600</td>
	    <td width="85%" class='forumRowHighLight'><table width="100%" border="0" cellspacing="0" cellpadding="0">
         <tr>
           <td ><input id="Pics" name="Pics" type="text" size="80" /> <br><input type="button" value="上传图片" onClick="showUpload2(null,'Pics','',100,null);" />
		   </td>
         </tr>
       </table></td>
      </tr>
 

<tr>
        <td  class='forumRow' height=23>案例关键字</td>
	      <td class='forumRow'><input type='text' id='a_keywords' name='a_keywords' size='60'> <select name="keywords_list" id="keywords_list" onclick="document.form1.a_keywords.value=document.form1.keywords_list.value">
	      <option value="">请选择</option>
		   <% set rsp=server.createobject("adodb.recordset")
		   sql="select name from web_keywords order by [id] "
		   rsp.open(sql),cn,1,1
		   if not rsp.eof and not rsp.bof then
		   do while not rsp.eof 
		   %> <option value="<%=rsp("name")%>"  ><%=rsp("name")%></option>
            <%
			rsp.movenext
			loop
			end if
			rsp.close
			set rsp=nothing%></select>
	  &nbsp;请以，隔开(中文逗号)</td>
	</tr><tr>
	  <td class='forumRowHighLight' height=11>案例描述 / 案例摘要 </td>
	  <td class='forumRowHighLight'><textarea name='a_description'  cols="100" rows="4" id="a_description" ></textarea></td>
	</tr>
	<tr>
	  <td class='forumRow' height=23>案例内容 (必填) </td>
	  <td class='forumRow'><textarea name="a_content" id="a_content" style=" width:100%; height:400px; visibility:hidden;"></textarea>
</td>
	</tr>
	<tr>
	  <td class='forumRowHighLight' height=23>案例来源</td>
	  <td class='forumRowHighLight'>
	    <input name='a_from_name' type='text' id='a_from_name' size='30'>
	 	      <select name="position" id="position" onChange="show()">
	      <option value="">请选择</option>
		   <% set rsp=server.createobject("adodb.recordset")
		   sql="select name,[url] from web_article_author order by [order] "
		   rsp.open(sql),cn,1,1
		   if not rsp.eof and not rsp.bof then
		   do while not rsp.eof 
		   %> <option value="<%=rsp("url")%>"  ><%=rsp("name")%></option>
            <%
			rsp.movenext
			loop
			end if
			rsp.close
			set rsp=nothing%></select>	 </td>
	  </tr>
	<tr>
	  <td class='forumRow' height=23>来源网址</td>
	  <td class='forumRow'><span class="forumRow">
	    <input name='a_from_url' type='text' id='a_from_url' size='40'>
      &nbsp;以http://开头</span></td>
	</tr>
	<tr>
	  <td class='forumRowHighLight' height=23>案例作者</td>
	  <td class='forumRowHighLight'><span class="forumRow">
	    <input name='a_author' type='text' id='a_author' value="<%=Session("log_name")%>" size='40'>
	  </span></td>
	</tr>
	<tr>
	  <td class='forumRow' height=23>案例浏览次数</td>
	  <td class='forumRow'><input name='a_hit' type='text' id='a_hit' value="0" size='40'>
      &nbsp;只能是数字</td>
	  </tr>
	<tr>
	  <td class='forumRowHighLight' height=23>排序</td>
	  <td class='forumRowHighLight'><input name='a_OrderCount' type='text' id='a_OrderCount' value="10000" size='40'>
      &nbsp;只能是数字,数值越小排名越靠前</td>
	  </tr>      
	<tr>
	  <td class='forumRowHighLight' height=23>案例推荐</td>
	  <td class='forumRowHighLight'><input type="radio" name="a_index_push" value="1">
是
      &nbsp;
      <input name="a_index_push" type="radio" value="0" checked>
否</td>
	  </tr>	  

	<tr>
	  <td class='forumRow' height=23>替换关键词链接</td>
	  <td class='forumRow'><input name="a_keywords_yes" type="radio" value="1" >
是
      &nbsp;
      <input name="a_keywords_yes" type="radio" value="0" checked="checked">
否</td>
	  </tr>

	<tr><td height="50" colspan=2  class='forumRow'><div align="center">
	  <INPUT type=submit value='提交' onClick='javascript:return checksignup1()' name=Submit>
	  </div></td></tr>
	</table>
</form>

<%
Call DbconnEnd()
 %>
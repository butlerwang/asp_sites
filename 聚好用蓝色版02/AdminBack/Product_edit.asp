<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="../inc/rand.asp" -->
<!-- #include file="../inc/x_to_html/index_to_html.asp" -->
<!-- #include file="../inc/x_to_html/Product_to_html.asp" -->
<!-- #include file="../inc/x_to_html/Case_List_to_html.asp" -->
<% '更新数据到数据表
page=request.querystring("page")
act=request.querystring("act")
keywords=request.querystring("keywords")
article_id=cint(request.querystring("id"))



act1=Request("act1")
If act1="save" Then 
a_id=cint(request.form("a_id"))
a_title=request.form("a_title")
a_cid=trim(request.form("cid"))
a_pid=trim(request.form("pid"))
a_ppid=trim(request.form("ppid"))
a_url=trim(request.form("a_url"))
a_SaleCount=trim(request.form("SaleCount"))
a_SalePrice=trim(request.form("SalePrice"))
a_image=trim(request.form("web_image"))
a_keywords=trim(request.form("a_keywords"))
a_description=trim(request.form("a_description"))
a_content=request.form("a_content")
a_from_name=trim(request.form("a_from_name"))
a_from_url=trim(request.form("a_from_url"))
a_from_rank=trim(request.form("a_from_rank"))
a_author=trim(request.form("a_author"))
a_hit=trim(request.form("a_hit"))
a_index_push=trim(request.form("a_index_push"))
a_keywords_yes=trim(request.form("a_keywords_yes"))
a_slide_yes=cint(request.form("slide_yes"))
a_special_yes=cint(request.form("special_yes"))
a_time=now()
product_order_show=trim(request.form("product_order_show"))
product_tbbuy_url=trim(request.form("product_tbbuy_url"))


set rs=server.createobject("adodb.recordset")
sql="select * from [Article] where id="&a_id&""
rs.open(sql),cn,1,3
juhaoyong_a_cid=rs("cid")
juhaoyong_a_pid=rs("pid")
juhaoyong_a_ppid=rs("ppid")

rs("title")=a_title
rs("ArticleType")=2
rs("cid")=a_cid
rs("pid")=a_pid
rs("ppid")=a_ppid
'rs("url")=a_url
rs("SaleCount")=a_SaleCount
rs("SalePrice")=a_SalePrice
rs("wine")=trim(request.form("a_wine"))
rs("net")=trim(request.form("a_net"))
rs("image")=a_image
rs("keywords")=a_keywords
rs("description")=a_description
rs("content")=a_content
rs("from_name")=a_from_name
rs("from_url")=a_from_url
rs("author")=a_author
rs("hit")=a_hit
rs("index_push")=a_index_push
'rs("slide_yes")=a_slide_yes
'rs("special_yes")=a_special_yes
rs("headline")=a_headline
rs("edit_time")=a_time
rs("product_order_show")=product_order_show
rs("product_tbbuy_url")=product_tbbuy_url
rs.update
rs.close
set rs=nothing
%>

<%
'生成产品静态页和首页
call Product_to_html(a_id)
call index_to_html()

'根据分类变化情况生成各级别产品列表页
if juhaoyong_a_cid<>a_cid then
	if a_cid<>"" then call Case_List_to_html(a_cid)
	if juhaoyong_a_cid<>"" then call Case_List_to_html(juhaoyong_a_cid)
else
	if a_cid<>"" then call Case_List_to_html(a_cid)
end if

if juhaoyong_a_pid<>a_pid then
	if a_pid<>"" then call Case_List_to_html(a_pid)
	if juhaoyong_a_pid<>"" then call Case_List_to_html(juhaoyong_a_pid)
else
	if a_pid<>"" then call Case_List_to_html(a_pid)
end if

if juhaoyong_a_ppid<>a_ppid then
	if a_ppid<>"" then call Case_List_to_html(a_ppid)
	if juhaoyong_a_ppid<>"" then call Case_List_to_html(juhaoyong_a_ppid)
else
	if a_ppid<>"" then call Case_List_to_html(a_ppid)
end if

%>

<%
juhaoyong_cid=request.QueryString("juhaoyong_cid")
juhaoyong_pid=request.QueryString("juhaoyong_pid")
juhaoyong_ppid=request.QueryString("juhaoyong_ppid")

if juhaoyong_ppid>0 then
response.Write "<script language='javascript'>alert('修改成功！');location.href='Product_list.asp?ppid="&juhaoyong_ppid&"&page="&page&"&act="&act&"&keywords="&keywords&"';</script>"
elseif juhaoyong_pid>0 then
response.Write "<script language='javascript'>alert('修改成功！');location.href='Product_list.asp?pid="&juhaoyong_pid&"&page="&page&"&act="&act&"&keywords="&keywords&"';</script>"
elseif juhaoyong_cid>0 then
response.Write "<script language='javascript'>alert('修改成功！');location.href='Product_list.asp?cid="&juhaoyong_cid&"&page="&page&"&act="&act&"&keywords="&keywords&"';</script>"
end if

end if 
%>
  	<script charset="utf-8" src="Keditor/kindeditor.js"></script>
	<script charset="utf-8" src="Keditor/lang/zh_CN.js"></script>
	<script charset="utf-8" src="Keditor/editor.js"></script>
 
 <!-- 三级联动菜单 开始 -->
<script language="JavaScript">
<!--
<%
'二级数据保存到数组
Dim count2,rsc2,sqlClass2
set rsc2=server.createobject("adodb.recordset")
sqlClass2="select id,pid,ppid,name from category where ppid=2 and ClassType=2 order by id " 
rsc2.open sqlClass2,cn,1,1
%>
var subval2 = new Array();
//数组结构：一级根值,二级根值,二级显示值
<%
count2 = 0
do while not rsc2.eof
%>
subval2[<%=count2%>] = new Array('<%=rsc2("pID")%>','<%=rsc2("ID")%>','<%=rsc2("Name")%>')
<%
count2 = count2 + 1
rsc2.movenext
loop
rsc2.close
%>

<%
'三级数据保存到数组
Dim count3,rsc3,sqlClass3
set rsc3=server.createobject("adodb.recordset")
sqlClass3="select id,pid,ppid,name from category where ppid=3 and ClassType=2 order by id" 
rsc3.open sqlClass3,cn,1,1
%>
var subval3 = new Array();
//数组结构：二级根值,三级根值,三级显示值
<%
count3 = 0
do while not rsc3.eof
%>
subval3[<%=count3%>] = new Array('<%=rsc3("pID")%>','<%=rsc3("ID")%>','<%=rsc3("Name")%>')
<%
count3 = count3 + 1
rsc3.movenext
loop
rsc3.close
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

         <script language='javascript'>
function checksignup1() {
if ( document.form1.a_title.value == '' ) {
window.alert('请输入产品标题^_^');
document.form1.a_title.focus();
return false;}

if ( document.form1.cid.value == '' ) {
window.alert('请选择分类^_^');
document.form1.cid.focus();
return false;}


return true;}
</script>
<%
juhaoyong_cid=request.QueryString("juhaoyong_cid")
juhaoyong_pid=request.QueryString("juhaoyong_pid")
juhaoyong_ppid=request.QueryString("juhaoyong_ppid")
%>
<%
      
			set rs=server.createobject("adodb.recordset")
sql="select * from [Article] where id="&article_id&""
rs.open(sql),cn,1,1
if not rs.eof and not rs.bof then
%>  <form id="form1" name="form1" method="post" action="?act1=save&juhaoyong_cid=<%=juhaoyong_cid%>&juhaoyong_pid=<%=juhaoyong_pid%>&juhaoyong_ppid=<%=juhaoyong_ppid%>&page=<%=page%>&act=<%=act%>&keywords=<%=keywords%>">

	<table cellpadding='3' cellspacing='1' border='0' class='tableBorder' align=center>
	<tr>
	  <th class='tableHeaderText' colspan=2 height=25>修改产品</th>
	<tr>
	<td width="15%" height=23 class='forumRow'>标题 (必填) </td>
	<td class='forumRow'><input name='a_title' type='text' id='a_title' value="<%=rs("title")%>" size='70'>
	<input name='a_id' type='hidden' id='a_id' value="<%=rs("id")%>" size='70'>
	  &nbsp;</td>
	</tr>
	<tr>
	<td class='forumRowHighLight' height=23>分类<span class="forumRow"> (必选) </span></td>
    <td class='forumRowHighLight'><%
set rsc1=server.createobject("adodb.recordset")
sqlClass1="select id,pid,ppid,name from category where ppid=1 and ClassType=2 order by id" 
rsc1.open sqlClass1,cn,1,1
%>
            <select name="cid" id="cid" onChange="changeselect1(this.value)">
              <option value="">选择一级分类</option>
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
sqlClass2="select id,pid,ppid,name from category where ppid=2 and ClassType=2 and pid="&cint(rs("cid"))&" order by id" 
rsc2.open sqlClass2,cn,1,1

count1 = 0
do while not rsc2.eof
%><option value="<%=rsc2("ID")%>"  <%if rs("pid")<>"" then
if cint(rs("pid"))=rsc2("id") then
response.write "selected"
end if
end if%> ><%=rsc2("Name")%></option>
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
								if rs("ppid")<>"" then
set rsc3=server.createobject("adodb.recordset")
sqlClass3="select id,pid,ppid,name from category where ppid=3 and ClassType=2 and pid="&cint(rs("pid"))&" order by id" 
rsc3.open sqlClass3,cn,1,1

count1 = 0
do while not rsc3.eof
%><option value="<%=rsc3("ID")%>"  <%if cint(rs("ppid"))=rsc3("id") then
response.write "selected"
end if%>><%=rsc3("Name")%></option>
<%count1 = count1 + 1
rsc3.movenext
loop
rsc3.close
end if
%>
            </select>&nbsp;</td>
	</tr>
<tr>
	    <td class='forumRowHighLight' height=23>品牌 </td>
	    <td class='forumRowHighLight'><input name='SalePrice' type='text' id='SalePrice' size='30' value="<%=rs("SalePrice")%>"/><font color="#FF0000">（若为空，则前台详情页不显示该项）</font>
    </td>
      </tr>	
      
<tr>
	    <td class='forumRow' height=23>型号</td>
        <td class='forumRow'><input name='SaleCount' type='text' id='SaleCount' size='30' value="<%=rs("SaleCount")%>"/><font color="#FF0000">（若为空，则前台详情页不显示该项）</font>
        </td>
      </tr>	  
	  <tr>
	    <td class='forumRowHighLight' height=23>市场价</td>
	    <td class='forumRowHighLight'><input name='a_wine' type='text' id='a_wine' size='30'  value="<%=rs("wine")%>"><font color="#FF0000">（若为空，则前台详情页不显示该项）</font></td>
      </tr>
	  <tr>
	    <td class='forumRow' height=23>优惠价</td>
	    <td class='forumRow'><input name='a_net' type='text' id='a_net' size='30'  value="<%=rs("net")%>"><font color="#FF0000">（若为空，则前台详情页不显示该项）</font></td>
      </tr>      
      
	  <tr>
	    <td class='forumRowHighLight' height=23>产品图片</td>
	    <td width="85%" class='forumRowHighLight'><table width="100%" border="0" cellspacing="0" cellpadding="0">
         <tr>
           <td width="22%"  ><input name="web_image" type="text" id="web_image"  value="<%=rs("image")%>" size="30"></td>
           <td width="78%" ><iframe width="500" name="ad" frameborder=0 height=30 scrolling=no src=upload.asp?juhaoyongUploadFileName=<%=juhaoyongGetUploadFileName(rs("image"))%>></iframe><font color="#FF0000">（注：上传的图片名称必须是英文或数字）</font></td>
         </tr>
       </table></td>
      </tr>

        <td  class='forumRow' height=23>关键字</td>
	      <td class='forumRow'><input type='text' id='v3' name='a_keywords'  value="<%=rs("keywords")%>" size='60'> <select name="keywords_list" id="keywords_list" onclick="document.form1.a_keywords.value=document.form1.keywords_list.value">
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
	  <td class='forumRowHighLight' height=11>描述 </td>
	  <td class='forumRowHighLight'><textarea name='a_description'  cols="100" rows="4" id="a_description" ><%=rs("description")%></textarea></td>
	</tr>
	<tr>
	  <td class='forumRow' height=23>介绍 (必填) </td>
	  <td class='forumRow'> <textarea name="a_content" id="a_content" style=" width:100%; height:400px; visibility:hidden;"><%=rs("content")%></textarea>
  </td>
	</tr>
	<tr>
	  <td class='forumRowHighLight' height=23>作者</td>
	  <td class='forumRowHighLight'>
	  <span class="forumRow">
	    <input name='a_author' type='text' id='c_name32' value="<%=rs("author")%>" size='40'>
	  </span></td>
	</tr>
	
	<tr>
	  <td class='forumRow' height=23>是否显示“提交订单”按钮</td>
	  <td class='forumRow'><label>
	    <input type="radio" name="product_order_show" value="1" 
		<%if rs("product_order_show")=1 then
		response.write "checked"
		end if%>>
      是
      &nbsp;
        <input type="radio" name="product_order_show" value="0" 
		<%if rs("product_order_show")=0 then
		response.write "checked"
		end if%>>
      否</label></td>
	</tr>
	
	<tr>
	  <td class='forumRowHighLight' height=23>淘宝宝贝网址</td>
	  <td class='forumRowHighLight'>
	  <span class="forumRow">
	    <input name='product_tbbuy_url' type='text' id='c_name302' value="<%=rs("product_tbbuy_url")%>" size='120'>
	  </span>（为空则不显示“去淘宝拍”按钮）</td>
	</tr>
	
	<tr>
	  <td class='forumRow' height=23>浏览次数</td>
	  <td class='forumRow'><input name='a_hit' type='text' id='a_hit' value="<%=rs("hit")%>" size='40'>
      &nbsp;只能是数字</td>
	</tr>
	
	<tr>
	  <td class='forumRowHighLight' height=23>推荐</td>
	  <td class='forumRowHighLight'><label>
	    <input type="radio" name="a_index_push" value="1" 
		<%if rs("index_push")=1 then
		response.write "checked"
		end if%>>
      是
      &nbsp;
        <input type="radio" name="a_index_push" value="0" 
		<%if rs("index_push")=0 then
		response.write "checked"
		end if%>>
      否</label></td>
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
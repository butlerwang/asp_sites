﻿<%@ LANGUAGE=VBScript CodePage=65001%>
<% response.charset="utf-8" %>
<% session.codepage=65001 %>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="../inc/x_to_html/Article_to_html.asp" -->
<!-- #include file="../inc/x_to_html/Product_to_html.asp" -->
<link rel="stylesheet" type="text/css" href="/css/default/common.css"  />
<script  language="javascript" src="/js/date_input_1.js"></script>

<script language="javascript">

//全选JS
function unselectall(){
if(document.form1.chkAll.checked){
document.form1.chkAll.checked = document.form1.chkAll.checked&0;
}
}
function CheckAll(form){
for (var i=0;i<form.elements.length;i++){
var e = form.elements[i];
if (e.Name != 'chkAll'&&e.disabled==false)
e.checked = form.chkAll.checked;
}
}
</script>

<!-- 资讯文章三级联动菜单 开始 -->
<script language="JavaScript">
<!--
<%
'二级数据保存到数组
Dim count2,rsClass2,sqlClass2
set rsClass2=server.createobject("adodb.recordset")
sqlClass2="select id,pid,ppid,name from category where ppid=2  order by id " 
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
sqlClass3="select id,pid,ppid,name from category where ppid=3  order by id" 
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
</script><!-- 资讯文章三级联动菜单 结束 -->

	<%
Call header()
%>

<%'生成
if request.querystring("action")="create" then

'生成文章 start
if request.form("news_article")=1 then
cid=request.form("cid")
pid=request.form("pid")
ppid=request.form("ppid")
a_start=request.form("a_start")
a_end=request.form("a_end")

if cid<>"" or pid<>"" or ppid<>"" or  a_start<>"" or a_end<>"" then
'------------------
if ppid<>"" then
a_c_sql=" ppid='"&ppid&"'"
end if
if pid<>"" and ppid="" then
a_c_sql=" pid='"&pid&"'"
end if
if cid<>"" and pid="" and ppid="" then
a_c_sql=" cid='"&cid&"'"
end if
'------------------
if a_start<>"" and a_end<>"" then
a_t_sql=" [time]>=#"&a_start&"# and [time]<=#"&a_end&"# "
end if
if  a_end<>"" and a_start="" then
a_t_sql=" [time]<=#"&a_end&"# "
end if
if a_start<>"" and a_end="" then
a_t_sql=" [time]>=#"&a_start&"# "
end if
'------------------
if a_c_sql<>"" and a_t_sql<>"" then
sql="select [id],[ArticleType] from [article] where "&a_c_sql&" and "&a_t_sql&" and view_yes=1 order by [time] desc"
end if
if a_c_sql<>"" and a_t_sql="" then
sql="select [id],[ArticleType] from [article] where "&a_c_sql&"  and view_yes=1 order by [time] desc"
end if
if a_c_sql="" and a_t_sql<>"" then
sql="select [id],[ArticleType] from [article] where "&a_t_sql&"  and view_yes=1 order by [time] desc"
end if
else
sql="select [id],[ArticleType] from [article]  where view_yes=1 order by [time] desc"
end if
set rs_create=server.createobject("adodb.recordset")
rs_create.open(sql),cn,1,1
do while not rs_create.eof 
a_id=rs_create("id")
select case rs_create("ArticleType")
case 1
call article_to_html(a_id)
case 2
call Product_to_html(a_id)
end select
rs_create.movenext
loop
rs_create.close
set rs_create=nothing

end if
'生成资讯文章 end

response.Write "<script language='javascript'>alert('更新成功！');history.go(-1);</script>"


end if
%>

	<table cellpadding='3' cellspacing='1' border='0' class='tableBorder' align=center>
	<tr>
	  <th width="100%" height=25 class='tableHeaderText'>生成文章</th>
	
	<tr><td height="400" valign="top"  class='forumRow'><br>
	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="25" class='TipTitle'>&nbsp;√ 操作提示</td>
          </tr>
          <tr>
            <td height="30" valign="top" class="TipWords"><p>1、内容生成主要包括的是所有类型文章页面的生成。</p>
                <p>2、不建议全选后再生成，栏目过多，生成时间可以会过慢或导致超时。</p>
                <p>3、文章如果过多，可能会导致超时，请逐个选择分类后进行生成，减轻程序压力。</p></td>
          </tr>
          <tr>
            <td height="10" ></td>
          </tr>
        </table>
	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="100">
              <form name="form1" method="post" action="?action=create">
                <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
				      <tr > <td height="32" colspan="2"  class="TitleHighlight3"> &nbsp;
                      <label>
                      <input type="checkbox" name="news_article" value="1">
                      <span style="font-weight: bold">文章内容生成</span></label></td>
                  </tr>
				      <tr>
                    <td height="35" colspan="2" class="forumRowt" style="line-height:200%;"><table width="100%" border="0" cellspacing="3" cellpadding="0">
                      <tr class="forumRowHighLight">
                        <td width="8%" height="30" class="forumRowHighLight">&nbsp;选择分类</td>
                        <td width="92%" class="forumRowHighLight"> &nbsp;
                          <%
Dim count1,rsClass1,sqlClass1
set rsClass1=server.createobject("adodb.recordset")
sqlClass1="select id,pid,ppid,name from category where ppid=1  order by id" 
rsClass1.open sqlClass1,cn,1,1
%>
                          <select name="cid" id="cid" onChange="changeselect1(this.value)">
                            <option value="">选择一级分类</option>
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
</select></td>
                      </tr>
                      <tr>
                        <td height="30" class="forumRowHighLight">&nbsp;选择时间段</td>
                        <td height="30" class="forumRowHighLight">
&nbsp;从(大于)
  <input name="a_start"  type="text" value=""  readonly="readonly" onClick="showcalendar(event, this);" onFocus="showcalendar(event, this);if(this.value=='0000-00-00')this.value=''" />
&nbsp;&nbsp;&nbsp;到(小于)
<input name="a_end"  type="text" value=""  readonly="readonly" onClick="showcalendar(event, this);" onFocus="showcalendar(event, this);if(this.value=='0000-00-00')this.value=''" /></td>
                      </tr>
                    </table></td>
                  </tr>	
	
						  
                  <tr>
                    <td width="12%" height="50"><label>
                      &nbsp;
                      <input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'>
                    全选/全不选</label></td>
                    <td width="88%"><div align="center">
                      <input type="submit" name="Submit" value="提交">
                    </div></td>
                  </tr>
                </table>
              </form>            </td>
          </tr>
      </table>
	    </td>
	</tr>
	</table>

</body>
<script  language="javascript" src="/js/date_input_2.js"></script>
</html>
<%
Call DbconnEnd()
 %>
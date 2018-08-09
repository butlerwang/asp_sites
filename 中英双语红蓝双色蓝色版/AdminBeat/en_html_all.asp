<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="../inc/en_x_to_html/index_to_html.asp" -->
<!-- #include file="../inc/en_x_to_html/Article_list_to_html.asp" -->
<!-- #include file="../inc/en_x_to_html/post_index_to_html.asp" -->
<!-- #include file="../inc/en_x_to_html/Blank_Content_to_html.asp" -->
<!-- #include file="../inc/en_x_to_html/Recruit_list_to_html.asp" -->
<!-- #include file="../inc/en_x_to_html/Case_List_to_html.asp" -->
<!-- #include file="../inc/en_x_to_html/Search_index_to_html.asp" -->
<!-- #include file="../inc/en_x_to_html/SiteMap_index_to_html.asp" -->
<!-- #include file="../inc/en_x_to_html/Article_to_html.asp" -->
<!-- #include file="../inc/en_x_to_html/Product_to_html.asp" -->
<!-- #include file="../inc/en_x_to_html/order_index_to_html.asp" -->

	<%
Call header()
%>

<%'生成
'生成首页
call index_to_html()

'生成栏目
sql="select [id],ppid,ClassType,Html_Yes,index_push from [en_category]  order by [time] desc"
set rs_create=server.createobject("adodb.recordset")
rs_create.open(sql),cn,1,1
if not rs_create.eof then
do while not rs_create.eof
ClassID=rs_create("id")

'文章
if rs_create("ClassType")=1 then
call Article_list_to_html(ClassID)
end if

'产品
if rs_create("ClassType")=2 then
call Case_List_to_html(ClassID)
end if

'招聘
if rs_create("ClassType")=4 then
call Recruit_list_to_html(ClassID)
end if

'单页
if rs_create("ClassType")=5  then
call Blank_Content_to_html(ClassID)
end if

rs_create.movenext
loop
end if
rs_create.close
set rs_create=nothing

'生成留言首页及列表
call post_index_to_html()

'生成搜索页
call search_index_to_html()

'生成网站地图
call SiteMap_to_html()

'生成订单页面
call order_index_to_html()

'生成资讯文章
sql="select [id],[ArticleType] from [en_article]  where view_yes=1 order by [time] desc"
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

response.Write "<script language='javascript'>alert('更新成功！');history.go(-1);</script>"
%>



<%
Call DbconnEnd()
 %>
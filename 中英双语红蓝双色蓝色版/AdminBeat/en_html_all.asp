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

<%'����
'������ҳ
call index_to_html()

'������Ŀ
sql="select [id],ppid,ClassType,Html_Yes,index_push from [en_category]  order by [time] desc"
set rs_create=server.createobject("adodb.recordset")
rs_create.open(sql),cn,1,1
if not rs_create.eof then
do while not rs_create.eof
ClassID=rs_create("id")

'����
if rs_create("ClassType")=1 then
call Article_list_to_html(ClassID)
end if

'��Ʒ
if rs_create("ClassType")=2 then
call Case_List_to_html(ClassID)
end if

'��Ƹ
if rs_create("ClassType")=4 then
call Recruit_list_to_html(ClassID)
end if

'��ҳ
if rs_create("ClassType")=5  then
call Blank_Content_to_html(ClassID)
end if

rs_create.movenext
loop
end if
rs_create.close
set rs_create=nothing

'����������ҳ���б�
call post_index_to_html()

'��������ҳ
call search_index_to_html()

'������վ��ͼ
call SiteMap_to_html()

'���ɶ���ҳ��
call order_index_to_html()

'������Ѷ����
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

response.Write "<script language='javascript'>alert('���³ɹ���');history.go(-1);</script>"
%>



<%
Call DbconnEnd()
 %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="X-UA-Compatible" content="IE=7">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<!-- #include file="../inc/AntiAttack.asp" -->
<!-- #include file="../inc/conn.asp" -->
<!-- #include file="../inc/web_config.asp" -->
<!-- #include file="../inc/html_clear.asp" -->
<%
search_q=request.querystring("q")
%>
<title>搜索：<%=search_q%>_企业网站管理系统</title>
<meta name="keywords" content="$Class_Keywords$" />
<meta name="description" content="$Class_Description$" />
<link href="/css/HituxCMSBlue/inner.css" rel="stylesheet" type="text/css" />
<link href="/css/HituxCMSBlue/common.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="/js/functions.js"></script>
<script type="text/javascript" src="/images/iepng/iepngfix_tilebg.js"></script>
<script type="text/javascript">
window.onerror=function(){return true;}
</script></head>

<body>
<%
keywords=split(search_q," ")
c=ubound(keywords)
for i=0 to c
if i=0 then
search_sql1=search_sql1&"where  ( [title] like '%"&keywords(i)&"%'"
keywords_all=keywords(i)
else
search_sql1=search_sql1&" or   [title] like '%"&keywords(i)&"%'"
keywords_all=keywords_all&"+"&keywords(i)
end if
next

s_sql="select [title],[content],[file_path],[time],ArticleType from [article] "&search_sql1&" )  and view_yes=1 order by [time] desc"
%>
<div id="wrapper">

<!--head start-->
<div id="head">

<!--top start -->
<div class="top">
<div class="TopLogo">
<div class="logo"><a href="/"><img src="/images/up_images/20131117142552.png" alt="企业网站管理系统"></a></div>
<div class="tel">
<p class="telW">客服热线</p>
<p class="telN">400-888-888</p></div>

</div>

</div>
<!--top end-->
<!--nav start-->
<div id="NavLink">
<div class="NavBG">
<!--Head Menu Start-->
<ul id='sddm'><li class='CurrentLi'><a href='/'>网站首页</a></li> <li><a href='/About/' onmouseover=mopen('m2') onmouseout='mclosetime()'>关于公司</a> <div id='m2' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'><a href='/About/'>公司介绍</a> <a href='/Honor/'>公司荣誉</a> <a href='/About/Group'>组织机构</a> <a href='/About/Culture'>企业文化</a> <a href='/About/Enviro'>公司环境</a> <a href='/About/Business'>业务介绍</a> </div></li> <li><a href='/Honor/' onmouseover=mopen('m3') onmouseout='mclosetime()'>公司荣誉</a> <div id='m3' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'><a href='/Honor/2011/'>2011年</a> <a href='/Honor/2012/'>2012年</a> <a href='/Honor/2013/'>2013年</a> </div></li> <li><a href='/news/' onmouseover=mopen('m4') onmouseout='mclosetime()'>新闻动态</a> <div id='m4' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'><a href='/news/CompanyNews'>公司新闻</a> <a href='/news/IndustryNews'>行业新闻</a> </div></li> <li><a href='/Product/' onmouseover=mopen('m5') onmouseout='mclosetime()'>公司产品</a> <div id='m5' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'><a href='/Product/DigitalPlayer'>数码播放器</a> <a href='/Product/Pad'>平板电脑</a> <a href='/Product/GPS'>GPS导航</a> <a href='/Product/NoteBook'>笔记本电脑</a> <a href='/Product/Mobile'>智能手机</a> </div></li> <li><a href='/Recruit/' onmouseover=mopen('m6') onmouseout='mclosetime()'>人才招聘</a> <div id='m6' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'><a href='/recruit/peiyang'>人才培养</a> <a href='/recruit/fuli'>福利待遇</a> <a href='/recruit/jobs'>招聘职位</a> </div></li> <li><a href='/Case/' onmouseover=mopen('m7') onmouseout='mclosetime()'>案例展示</a> <div id='m7' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'><a href='/Case/Case1/'>案例分类一</a> <a href='/Case/Case2/'>案例分类二</a> </div></li> <li><a href='/contact/'>联系我们</a></li> <li><a href='/Feedback/'>访客留言</a></li> </ul>
<!--Head Menu End-->
</div>
<div class="clearfix"></div>
</div>
<!--nav end-->

<!--focus start-->
<div id="InnerFocus">
<div id="FocusBG">
</div>
</div>
<!--foncus end-->
<div class="HeightTab clearfix"></div>
</div>
<!--head end-->
<!--body start-->
<div id="body">
<!--focus start-->
<div id="InnerBanner">
<script src='/ADs/106.js' type='text/javascript'></script>
</div>
<!--foncus end-->
<div class="HeightTab clearfix"></div>
<!--inner start -->
<div class="inner">
<!--left start-->
<div class="left">

<div class="Sbox">
<div class="topic">搜索 Search</div>
<div class="SearchBar">
<form method="get" action="/Search/index.asp">
				<input type="text" name="q" id="search-text" size="15" onBlur="if(this.value=='') this.value='请输入关键词';" 
onfocus="if(this.value=='请输入关键词') this.value='';" value="请输入关键词" /><input type="submit" id="search-submit" value="搜索" />
			</form>
</div>
</div>

</div>
<!--left end-->
<!--right start-->
<div class="right">
<div class="Position"><span>你的位置：<a href="/">首页</a> > 搜索</span></div>
<div class="HeightTab clearfix"></div>
<!--main start-->
<div class="main">

<!--search content start-->
<div id="search_content" class="clearfix">

<%
if search_q<>"" then 

set rs=server.createobject("adodb.recordset")
rs.open(s_sql),cn,1,1
%>

<%'=============分页定义开始，要放在数据库打开之后
if err.number<>0 then '错误处理
response.write "数据库操作失败：" & err.description
err.clear
else
if not (rs.eof and rs.bof) then '检测记录集是否为空
r=cint(rs.RecordCount) '记录总数
rowcount = 10 '设置每一页的数据记录数，可根据实际自定义
rs.pagesize = rowcount '分页记录集每页显示记录数
maxpagecount=rs.pagecount '分页页数
page=request.querystring("page")
  if page="" then
  page=1
  end if
rs.absolutepage = page 
rcount1=0
pagestart=page-5
pageend=page+5
if pagestart<1 then
pagestart=1
end if
if pageend>maxpagecount then
pageend=maxpagecount
end if
rcount=rs.RecordCount
'=============分页定义结束%>

<!--position start-->
<div class="searchtip">您正在搜索“<span class="FontRed"><%=search_q%></span>”,找到相关信息 <span class="font_brown"><%=rcount%></span> 条</div>
<!--position end-->
<!--list start-->
<div class="result_list">
<div class="gray">提示：用空格隔开多个搜寻关键词可获取更理想结果，如“最新 产品”。</div>
<dl>

<%'===========循环体开始
do while not rs.eof and rowcount%>
<%
select case rs("ArticleType")
case 1
Content_FolderName=Article_FolderName
case 2
Content_FolderName=Product_FolderName
case 3
Content_FolderName=Case_FolderName
end select

title1=left(rs("title"),30)
for i=0 to c
title1=Replace(title1, keywords(i), "<span class='FontRed'>" & keywords(i)& "</span>")
next

content1=left(nohtml(rs("content")),110)
for i=0 to c
content1=Replace(content1,keywords(i), "<span class='FontRed'>" & keywords(i)& "</span>")
next
%>
<dt ><a href='<%="/"&Content_FolderName&"/"&rs("file_path")%>' target='_blank' title='<%=rs("title")%>'><%=title1%></a></dt>
<dd><%=content1%>...</dd>
<dd class="font12 arial font_green line"><a href='<%="/"&Content_FolderName&"/"&rs("file_path")%>' target='_blank'><span class="font_green"><%=web_url&"/"&Content_FolderName&"/"&rs("file_path")%></span></a><%=year(rs("time"))%>-<%=month(rs("time"))%>-<%=day(rs("time"))%></dd>
<%
rowcount=rowcount-1 
rs.movenext
loop
 '===========循环体结束%>

</dl>
</div>
<!--list end-->

<!--page start-->
<div class="result_page clearfix">
<!--#include file="../inc/page_list.asp"-->
</div>
<!--page end-->

<%
else
response.write "<div class='search_welcome'>很抱歉,没有找到与 <span class='FontRed'>"&search_q&"</span> 相关的信息！<p >提示：用空格隔开多个搜寻关键词可获取更理想结果，如“最新 产品”。</p></div>"
end if
end if
end if%>
</div>
<!--search content end-->	

</div>
<!--main end-->
</div>
<!--right end-->
</div>
<!--inner end-->
</div>
<!--body end-->
<div class="HeightTab clearfix"></div>
<!--footer start-->
<div id="footer">
<div class="inner">
<p><a href="/">网站首页</a> | <a href="/About/">关于我们</a> | <a href="/Contact">联系方式</a> | <a href="/Sitemap">网站地图</a></p>

<p>Copyright   2013 企业网站管理系统 boot007.taobao.com 版权所有 All rights reserved</p>

<p>沪ICP备0000000号 <a href="http://boot007.taobao.com/" target="_blank">启网网络</a> 技术支持 <a href="http://boot007.taobao.com/" target="_blank">boot007</a> <a href="/rss" target="_blank"><img src="/images/rss_icon.gif"></a> <a href="/rss/feed.xml" target="_blank"><img src="/images/xml_icon.gif"></a>

</p>
</div>
</div>
<!--footer end -->


</div>
<script type="text/javascript" src="/js/ServiceCenter.js"></script>

</body>
</html>
<!--
Powered By HituxCMS ASP V2.1   
-->




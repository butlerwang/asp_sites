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
<title>搜索：<%=search_q%>_企业网站管理系统 JuhaoyongCMS</title>
<meta name="keywords" content="$Class_Keywords$" />
<meta name="description" content="$Class_Description$" />
<link href="/css/juhaoyongfgstyle/inner.css" rel="stylesheet" type="text/css" />
<link href="/css/juhaoyongfgstyle/common.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="/css/juhaoyongfgstyle/functions.js"></script>

</head>

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
<div class="logo"><a href="/"><img src="/css/juhaoyongfgstyle/logo.gif" alt="企业网站管理系统 JuhaoyongCMS"></a></div>
</div>

<div class="TopInfo">
<div class="link"><a href="/">网站首页</a> | <a href="/Sitemap">网站导航</a> | <a href="/Feedback">留言反馈</a></div>
<div class="clearfix"></div>
<div class="tel"><p class="telW">24小时客服热线</p><p class="telN">400-800-888</p></div>
<div class="juhaoyongTopSearchClass">
	<form method="get" action="/Search/index.asp">
	<span class="SearchBar">
	<input type="text" name="q" id="search-text" size="15" onBlur="if(this.value=='') this.value='请输入关键词';" onfocus="if(this.value=='请输入关键词') this.value='';" value="请输入关键词" />
	<input type="submit" id="search-submit" value="搜索" />
	</span>
	</form>
</div>
</div>

</div>
<!--top end-->

<!--nav start-->
<div id="NavLink">
<div class="NavBG">
<!--Head Menu Start-->
<ul id='sddm'><li class='CurrentLi'><a href='/'>网站首页</a></li> <li id='aaabbb2' onmouseover=juhaoyongNavBgaColor('aaabbb2') onmouseout=style.background=''><a href='/About/' onmouseover=mopen('m2') onmouseout='mclosetime()'>关于公司</a> <div id='m2' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'><a href='/About/intro'>公司介绍</a> <a href='/About/Group'>组织机构</a> <a href='/About/Culture'>企业文化</a> <a href='/About/Enviro'>公司环境</a> <a href='/About/Business'>业务介绍</a> </div></li> <li id='aaabbb3' onmouseover=juhaoyongNavBgaColor('aaabbb3') onmouseout=style.background=''><a href='/news/' onmouseover=mopen('m3') onmouseout='mclosetime()'>公告动态</a> <div id='m3' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'><a href='/About/introgrf'>tgerfeqdwq</a> <a href='/News/announcement'>公司公告</a> <a href='/news/CompanyNews'>公司新闻</a> <a href='/news/IndustryNews'>行业新闻</a> </div></li> <li><a href='/Product/'>公司产品</a></li> <li id='aaabbb5' onmouseover=juhaoyongNavBgaColor('aaabbb5') onmouseout=style.background=''><a href='/BusinessScope/' onmouseover=mopen('m5') onmouseout='mclosetime()'>业务范围</a> <div id='m5' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'><a href='/BusinessScope/content'>服务内容</a> <a href='/BusinessScope/area'>服务地区</a> </div></li> <li><a href='/DisplayCase'>案例展示</a></li> <li id='aaabbb7' onmouseover=juhaoyongNavBgaColor('aaabbb7') onmouseout=style.background=''><a href='/Recruit' onmouseover=mopen('m7') onmouseout='mclosetime()'>公司人才</a> <div id='m7' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'><a href='/recruit/peiyang'>人才培养</a> <a href='/recruit/fuli'>福利待遇</a> <a href='/recruit/jobs'>招聘职位</a> </div></li> <li id='aaabbb8' onmouseover=juhaoyongNavBgaColor('aaabbb8') onmouseout=style.background=''><a href='/ServiceCenter/' onmouseover=mopen('m8') onmouseout='mclosetime()'>客服中心</a> <div id='m8' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'><a href='/ServiceCenter/FAQ'>常见问题</a> <a href='/ServiceCenter/Download'>文档下载</a> <a href='/ServiceCenter/Services'>售后服务</a> <a href='/ServiceCenter/contactus'>联系我们</a> </div></li> </ul>
<!--Head Menu End-->
</div>
<div class="clearfix"></div>
</div>
<!--nav end-->

</div>
<!--head end-->
<!--body start-->
<div id="body">
<!--focus start-->
<div id="InnerBanner">

</div>
<!--foncus end-->
<div class="HeightTab clearfix"></div>
<!--inner start -->
<div class="inner">
<!--left start-->
<div class="left">
<div class="Sbox">
<div class="topic">联系我们&nbsp;&nbsp;&nbsp;Contact</div>
<div class="txt ColorLink">
<p>地址：西安市高新区XX路XX号</p>
<p>电话：029-88888888</p>
<p>手机：029-88888888</p>
<p>传真：029-88888888</p>
<p>邮件：admin@company.com</p>
<p>网站：<a href='http://www.juhaoyong.net' target='_blank'>http://www.juhaoyong.net</a> </p>
</div>
</div>
<div class="HeightTab clearfix"></div>

<div class="Sbox">
<div class="topic">搜索&nbsp;&nbsp;&nbsp;Search</div>
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
<div class='BottomNav'><a href="/">网站首页</a> | <a href="/About">关于我们</a> | <a href="/Recruit">人才招聘</a>  | <a href="/Sitemap">网站地图</a> | <a href="/Feedback">留言反馈</a>  | <a href="/RSS" target=_blank>订阅RSS</a></div>
<div class='HeightTab'></div>
<p>Copyright 2012 <a href='http://www.juhaoyong.net' target='_blank'>www.juhaoyong.net</a> 某某某某有限公司 版权所有 All Rights Reserved </p>
<p>公司地址：上海徐汇区XX路XX号 联系电话：021-00000000 电子邮件：admin@juhaoyong.net</p>
<p>Built By <a href="http://www.juhaoyong.net/" target="_blank">JuhaoyongCMS V2013</a> <a href="http://www.juhaoyong.net/" target="_blank">聚好用</a> 技术支持 <a href="/rss" target="_blank"><img src="/images/rss_icon.gif"></a> <a href="/rss" target="_blank"><img src="/images/xml_icon.gif"></a></p>
<DIV id=juhaoyong_xuanfukefu><DIV id=juhaoyong_xuanfukefuBut onmouseover='ShowJhyXuanfu()'><table class=juhaoyong_xuanfukefuBut_table border=0 cellspacing=0 cellpadding=0><tr><td> </td></tr></table></DIV><DIV id=juhaoyong_xuanfukefuContent><table width=143 border=0 cellspacing=0 cellpadding=0><tr><td class=juhaoyong_xuanfukefuContent01 valign=top> </td></tr><tr><td class=juhaoyong_xuanfukefuContent02 align=center>	<table border=0 cellspacing=0 cellpadding=0 align=center>    <tr><td class=jhykefu_box1>售前QQ客服</td></tr>    <tr><td class=jhykefu_box2><a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=1000000000&site=qq&menu=yes"><img border="0" src="http://wpa.qq.com/pa?p=2:1000000000:41 &r=0.11310566705699291" alt="点击这里给我发消息" title="点击这里给我发消息"></a></td></tr>    <tr><td class=jhykefu_box1>售后QQ客服</td></tr>    <tr><td class=jhykefu_box2><a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=1000000000&site=qq&menu=yes"><img border="0" src="http://wpa.qq.com/pa?p=2:1000000000:41 &r=0.11310566705699291" alt="点击这里给我发消息" title="点击这里给我发消息"></a></td></tr>    <tr><td class=jhykefu_box1>售前旺旺客服</td></tr>    <tr><td class=jhykefu_box2><a target="_blank" href="http://www.taobao.com/webww/ww.php?ver=3&touid=%E6%917%AA%918A%A&siteid=cntaobao&status=1&charset=utf-8"><img border="0" src="http://amos.alicdn.com/online.aw?v=2&uid=%E6%917%AA%918A%A&site=cntaobao&s=1&charset=utf-8" alt="点这里给我发消息" /></a></td></tr>    <tr><td class=jhykefu_box1>售后旺旺客服</td></tr>    <tr><td class=jhykefu_box2><a target="_blank" href="http://www.taobao.com/webww/ww.php?ver=3&touid=%E6%917%AA%918A%A&siteid=cntaobao&status=1&charset=utf-8"><img border="0" src="http://amos.alicdn.com/online.aw?v=2&uid=%E6%917%AA%918A%A&site=cntaobao&s=1&charset=utf-8" alt="点这里给我发消息" /></a></td></tr>	</table></td></tr>	<tr><td class=juhaoyong_xuanfukefuContent03 onclick=window.location.href='/Contact/'> </td></tr></table></DIV></DIV>
</div>
</div>
<!--footer end -->


</div>
<script type="text/javascript">
window.onerror=function(){return true;}
</script>

<script type="text/javascript" src="/juhaoyong-kfimgs/jquery.js"></script>
<script type="text/javascript" src="/juhaoyong-kfimgs/kf.js"></script>
</body>
</html>




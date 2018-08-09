<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="X-UA-Compatible" content="IE=7">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!-- #include file="../inc/AntiAttack.asp" -->
<!-- #include file="../inc/conn.asp" -->
<!-- #include file="../inc/web_config.asp" -->
<!-- #include file="../inc/html_clear.asp" -->
<%
search_q=request.querystring("q")
%>
<title>Search_<%=search_q%>_English Templates</title>
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
<div class="TopInfo"><div class="link"> <a href="/">Home</a> | <a href="/Contact">Contact</a> | <a href="/Sitemap">Sitemap</a></div>
</div>
<div class="clearfix"></div>
<div class="TopLogo">
<div class="logo"><a href="/"><img src="/images/up_images/20131117143754.png" alt="English Templates"></a></div>
<div class="tel">
<p class="telW">Hotline</p>
<p class="telN">400-888-888</p></div>

</div>
<div class="clearfix"></div>

</div>
<!--top end-->
<!--nav start-->
<div id="NavLink">
<div class="NavBG">
<!--Head Menu Start-->
<ul id='sddm'><li class='CurrentLi'><a href='/'>HOME</a></li> <li><a href='/About' onmouseover=mopen('m2') onmouseout='mclosetime()'>ABOUT</a> <div id='m2' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'><a href='/About'>Introduction</a> <a href='/About/Group'>Team</a> <a href='/About/Culture'>Culture</a> <a href='/About/Enviro'>Environment</a> <a href='/About/Business'>Business</a> </div></li> <li><a href='/Honor/' onmouseover=mopen('m3') onmouseout='mclosetime()'>HONOR</a> <div id='m3' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'><a href='/Honor/2011/'>2011 Year</a> <a href='/Honor/2012/'>2012 Year</a> <a href='/Honor/2013/'>2013 Year</a> </div></li> <li><a href='/news/' onmouseover=mopen('m4') onmouseout='mclosetime()'>NEWS</a> <div id='m4' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'><a href='/news/CompanyNews'>Company News</a> <a href='/news/IndustryNews'>Industry News</a> </div></li> <li><a href='/Product/' onmouseover=mopen('m5') onmouseout='mclosetime()'>PRODUCT</a> <div id='m5' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'><a href='/Product/DigitalPlayer'>Digital Player</a> <a href='/Product/Pad'>Tablet</a> <a href='/Product/GPS'>GPS</a> <a href='/Product/NoteBook'>NoteBook</a> <a href='/Product/Mobile'>Mobile</a> </div></li> <li><a href='/Case/' onmouseover=mopen('m6') onmouseout='mclosetime()'>CASE</a> <div id='m6' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'><a href='/Case/Case1/'>Case 1</a> <a href='/Case/Case2/'>Case 2</a> </div></li> <li><a href='/Recruit/' onmouseover=mopen('m7') onmouseout='mclosetime()'>RECRUIT</a> <div id='m7' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'><a href='/recruit/peiyang'>Talented</a> <a href='/recruit/fuli'>Fuli</a> <a href='/recruit/jobs'>Jobs</a> </div></li> <li><a href='/contact/'>CONTACT</a></li> <li><a href='/Feedback/'>FEEDBACK</a></li> </ul>
<!--Head Menu End-->
</div>
<div class="clearfix"></div>
</div>
<!--nav end-->

<div class="clearfix"></div>
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
<div class="topic">Search</div>
<div class="SearchBar">
<form method="get" action="/Search/index.asp">
				<input type="text" name="q" id="search-text" size="15" onBlur="if(this.value=='') this.value='Input Keywords...';" 
onfocus="if(this.value=='Input Keywords...') this.value='';" value="Input Keywords..." /><input type="submit" id="search-submit" value=" " />
			</form>
</div>
</div>

<div class="HeightTab clearfix"></div>
<div class="Sbox">
<div class="topic"><div class='TopicTitle'>Contact Us</div></div>
<div class="ContactInfo">
<p>CompanyName Group Co.,ltd</p>
<p>ADD：Zhongshan Road No.311 Zikawei District Shanghai City China</p>
<p>Tel：000-40324190</p>
<p>Fax：000-40324190</p>
<p>Web：<a href='http://boot007.taobao.com' target='_blank'>boot007.taobao.com</a></p>
</div>

</div>
<div class="HeightTab clearfix"></div>

</div>
<!--left end-->
<!--right start-->
<div class="right">
<div class="Position"><span>Position:<a href='/'>Home</a> > Search</span></div>
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
'=============page end%>

<!--position start-->
<div class="searchtip">You are searching"<span class="FontRed"><%=search_q%></span>",We found Results <span class="font_brown"><%=rcount%></span> </div>
<!--position end-->
<!--list start-->
<div class="result_list">
<div class="gray">Tip：Insert Tab into your keywords for more results</div>
<dl>

<%'==========round start
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
 '===========round end%>

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
response.write "<div class='search_welcome'>Sorry,No results for <span class='FontRed'>"&search_q&"</span><p >Tip:insert tab into your keywords for more results.</p></div>"
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
<div class="clearfix"></div>
</div>
<!--body end-->
<div class="clearfix"></div>
<!--footer start-->
<div id="footer">
<div class="inner">
<div class='BottomNav'><a href="/">HOME</a> | <a href="/About">INSTRUCTION</a> | <a href="/Contact">CONTACT US</a> | <a href="/Sitemap">SITEMAP</a></div>

<p>Copyright   2013 boot007.taobao.com  All rights reserved</p>

<p> <a href="http://boot007.taobao.com/" target="_blank">boot007.taobao.com</a> Technology Support <a href="http://boot007.taobao.com/" target="_blank">HituxCMS V2.1 English</a> <a href="/rss" target="_blank"><img src="/images/rss_icon.gif"></a> <a href="/rss/feed.xml" target="_blank"><img src="/images/xml_icon.gif"></a>

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


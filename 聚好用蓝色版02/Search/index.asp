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
<title>������<%=search_q%>_��ҵ��վ����ϵͳ JuhaoyongCMS</title>
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
<div class="logo"><a href="/"><img src="/css/juhaoyongfgstyle/logo.gif" alt="��ҵ��վ����ϵͳ JuhaoyongCMS"></a></div>
</div>

<div class="TopInfo">
<div class="link"><a href="/">��վ��ҳ</a> | <a href="/Sitemap">��վ����</a> | <a href="/Feedback">���Է���</a></div>
<div class="clearfix"></div>
<div class="tel"><p class="telW">24Сʱ�ͷ�����</p><p class="telN">400-800-888</p></div>
<div class="juhaoyongTopSearchClass">
	<form method="get" action="/Search/index.asp">
	<span class="SearchBar">
	<input type="text" name="q" id="search-text" size="15" onBlur="if(this.value=='') this.value='������ؼ���';" onfocus="if(this.value=='������ؼ���') this.value='';" value="������ؼ���" />
	<input type="submit" id="search-submit" value="����" />
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
<ul id='sddm'><li class='CurrentLi'><a href='/'>��վ��ҳ</a></li> <li id='aaabbb2' onmouseover=juhaoyongNavBgaColor('aaabbb2') onmouseout=style.background=''><a href='/About/' onmouseover=mopen('m2') onmouseout='mclosetime()'>���ڹ�˾</a> <div id='m2' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'><a href='/About/intro'>��˾����</a> <a href='/About/Group'>��֯����</a> <a href='/About/Culture'>��ҵ�Ļ�</a> <a href='/About/Enviro'>��˾����</a> <a href='/About/Business'>ҵ�����</a> </div></li> <li id='aaabbb3' onmouseover=juhaoyongNavBgaColor('aaabbb3') onmouseout=style.background=''><a href='/news/' onmouseover=mopen('m3') onmouseout='mclosetime()'>���涯̬</a> <div id='m3' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'><a href='/About/introgrf'>tgerfeqdwq</a> <a href='/News/announcement'>��˾����</a> <a href='/news/CompanyNews'>��˾����</a> <a href='/news/IndustryNews'>��ҵ����</a> </div></li> <li><a href='/Product/'>��˾��Ʒ</a></li> <li id='aaabbb5' onmouseover=juhaoyongNavBgaColor('aaabbb5') onmouseout=style.background=''><a href='/BusinessScope/' onmouseover=mopen('m5') onmouseout='mclosetime()'>ҵ��Χ</a> <div id='m5' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'><a href='/BusinessScope/content'>��������</a> <a href='/BusinessScope/area'>�������</a> </div></li> <li><a href='/DisplayCase'>����չʾ</a></li> <li id='aaabbb7' onmouseover=juhaoyongNavBgaColor('aaabbb7') onmouseout=style.background=''><a href='/Recruit' onmouseover=mopen('m7') onmouseout='mclosetime()'>��˾�˲�</a> <div id='m7' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'><a href='/recruit/peiyang'>�˲�����</a> <a href='/recruit/fuli'>��������</a> <a href='/recruit/jobs'>��Ƹְλ</a> </div></li> <li id='aaabbb8' onmouseover=juhaoyongNavBgaColor('aaabbb8') onmouseout=style.background=''><a href='/ServiceCenter/' onmouseover=mopen('m8') onmouseout='mclosetime()'>�ͷ�����</a> <div id='m8' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'><a href='/ServiceCenter/FAQ'>��������</a> <a href='/ServiceCenter/Download'>�ĵ�����</a> <a href='/ServiceCenter/Services'>�ۺ����</a> <a href='/ServiceCenter/contactus'>��ϵ����</a> </div></li> </ul>
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
<div class="topic">��ϵ����&nbsp;&nbsp;&nbsp;Contact</div>
<div class="txt ColorLink">
<p>��ַ�������и�����XX·XX��</p>
<p>�绰��029-88888888</p>
<p>�ֻ���029-88888888</p>
<p>���棺029-88888888</p>
<p>�ʼ���admin@company.com</p>
<p>��վ��<a href='http://www.juhaoyong.net' target='_blank'>http://www.juhaoyong.net</a> </p>
</div>
</div>
<div class="HeightTab clearfix"></div>

<div class="Sbox">
<div class="topic">����&nbsp;&nbsp;&nbsp;Search</div>
<div class="SearchBar">
<form method="get" action="/Search/index.asp">
				<input type="text" name="q" id="search-text" size="15" onBlur="if(this.value=='') this.value='������ؼ���';" 
onfocus="if(this.value=='������ؼ���') this.value='';" value="������ؼ���" /><input type="submit" id="search-submit" value="����" />
			</form>
</div>
</div>

</div>
<!--left end-->
<!--right start-->
<div class="right">
<div class="Position"><span>���λ�ã�<a href="/">��ҳ</a> > ����</span></div>
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

<%'=============��ҳ���忪ʼ��Ҫ�������ݿ��֮��
if err.number<>0 then '������
response.write "���ݿ����ʧ�ܣ�" & err.description
err.clear
else
if not (rs.eof and rs.bof) then '����¼���Ƿ�Ϊ��
r=cint(rs.RecordCount) '��¼����
rowcount = 10 '����ÿһҳ�����ݼ�¼�����ɸ���ʵ���Զ���
rs.pagesize = rowcount '��ҳ��¼��ÿҳ��ʾ��¼��
maxpagecount=rs.pagecount '��ҳҳ��
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
'=============��ҳ�������%>

<!--position start-->
<div class="searchtip">������������<span class="FontRed"><%=search_q%></span>��,�ҵ������Ϣ <span class="font_brown"><%=rcount%></span> ��</div>
<!--position end-->
<!--list start-->
<div class="result_list">
<div class="gray">��ʾ���ÿո���������Ѱ�ؼ��ʿɻ�ȡ�����������硰���� ��Ʒ����</div>
<dl>

<%'===========ѭ���忪ʼ
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
 '===========ѭ�������%>

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
response.write "<div class='search_welcome'>�ܱ�Ǹ,û���ҵ��� <span class='FontRed'>"&search_q&"</span> ��ص���Ϣ��<p >��ʾ���ÿո���������Ѱ�ؼ��ʿɻ�ȡ�����������硰���� ��Ʒ����</p></div>"
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
<div class='BottomNav'><a href="/">��վ��ҳ</a> | <a href="/About">��������</a> | <a href="/Recruit">�˲���Ƹ</a>  | <a href="/Sitemap">��վ��ͼ</a> | <a href="/Feedback">���Է���</a>  | <a href="/RSS" target=_blank>����RSS</a></div>
<div class='HeightTab'></div>
<p>Copyright 2012 <a href='http://www.juhaoyong.net' target='_blank'>www.juhaoyong.net</a> ĳĳĳĳ���޹�˾ ��Ȩ���� All Rights Reserved </p>
<p>��˾��ַ���Ϻ������XX·XX�� ��ϵ�绰��021-00000000 �����ʼ���admin@juhaoyong.net</p>
<p>Built By <a href="http://www.juhaoyong.net/" target="_blank">JuhaoyongCMS V2013</a> <a href="http://www.juhaoyong.net/" target="_blank">�ۺ���</a> ����֧�� <a href="/rss" target="_blank"><img src="/images/rss_icon.gif"></a> <a href="/rss" target="_blank"><img src="/images/xml_icon.gif"></a></p>
<DIV id=juhaoyong_xuanfukefu><DIV id=juhaoyong_xuanfukefuBut onmouseover='ShowJhyXuanfu()'><table class=juhaoyong_xuanfukefuBut_table border=0 cellspacing=0 cellpadding=0><tr><td> </td></tr></table></DIV><DIV id=juhaoyong_xuanfukefuContent><table width=143 border=0 cellspacing=0 cellpadding=0><tr><td class=juhaoyong_xuanfukefuContent01 valign=top> </td></tr><tr><td class=juhaoyong_xuanfukefuContent02 align=center>	<table border=0 cellspacing=0 cellpadding=0 align=center>    <tr><td class=jhykefu_box1>��ǰQQ�ͷ�</td></tr>    <tr><td class=jhykefu_box2><a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=1000000000&site=qq&menu=yes"><img border="0" src="http://wpa.qq.com/pa?p=2:1000000000:41 &r=0.11310566705699291" alt="���������ҷ���Ϣ" title="���������ҷ���Ϣ"></a></td></tr>    <tr><td class=jhykefu_box1>�ۺ�QQ�ͷ�</td></tr>    <tr><td class=jhykefu_box2><a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=1000000000&site=qq&menu=yes"><img border="0" src="http://wpa.qq.com/pa?p=2:1000000000:41 &r=0.11310566705699291" alt="���������ҷ���Ϣ" title="���������ҷ���Ϣ"></a></td></tr>    <tr><td class=jhykefu_box1>��ǰ�����ͷ�</td></tr>    <tr><td class=jhykefu_box2><a target="_blank" href="http://www.taobao.com/webww/ww.php?ver=3&touid=%E6%917%AA%918A%A&siteid=cntaobao&status=1&charset=utf-8"><img border="0" src="http://amos.alicdn.com/online.aw?v=2&uid=%E6%917%AA%918A%A&site=cntaobao&s=1&charset=utf-8" alt="��������ҷ���Ϣ" /></a></td></tr>    <tr><td class=jhykefu_box1>�ۺ������ͷ�</td></tr>    <tr><td class=jhykefu_box2><a target="_blank" href="http://www.taobao.com/webww/ww.php?ver=3&touid=%E6%917%AA%918A%A&siteid=cntaobao&status=1&charset=utf-8"><img border="0" src="http://amos.alicdn.com/online.aw?v=2&uid=%E6%917%AA%918A%A&site=cntaobao&s=1&charset=utf-8" alt="��������ҷ���Ϣ" /></a></td></tr>	</table></td></tr>	<tr><td class=juhaoyong_xuanfukefuContent03 onclick=window.location.href='/Contact/'> </td></tr></table></DIV></DIV>
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




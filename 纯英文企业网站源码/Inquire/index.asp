<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="X-UA-Compatible" content="IE=7">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Inquire_HuiguerCMS  English Templates</title>
<meta name="keywords" content="Sitemap" />
<link href="/css/HituxUnicode/inner.css" rel="stylesheet" type="text/css" />
<link href="/css/HituxUnicode/common.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="/js/functions.js"></script>
<script type="text/javascript" src="/images/iepng/iepngfix_tilebg.js"></script>
</head>

<body>
<!-- #include file="../inc/AntiAttack.asp" -->
<!-- #include file="../inc/conn.asp" -->
<!-- #include file="../inc/web_config.asp" -->
<!-- #include file="../inc/html_clear.asp" -->
<%
a_id=request.querystring("id")
%>
<%
set rs=server.createobject("adodb.recordset")
sql="select [title] from [article] where [id]="&a_id&" and view_yes=1"
rs.open(sql),cn,1,1
if not rs.eof then
ProductName=rs("title")
end if
rs.close 
set rs=nothing
%>
<div id="wrapper">

<!--head start-->
<div id="head">

<!--top start -->
<div class="top">
<div class="clearfix"></div>
<div class="TopLogo">
<div class="logo"><a href="/"><img src="/images/up_images/logo.png" alt="HuiguerCMS  English Templates"></a></div>
<div class="tel">
<div class="SearchBar">
<form method="get" action="/Search/index.asp">
				<input type="text" name="q" id="search-text" size="15" onBlur="if(this.value=='') this.value='Keywords';" 
onfocus="if(this.value=='Keywords') this.value='';" value="Keywords" /><input type="submit" id="search-submit" value=" " />
			</form>
</div>
</div>
</div>

</div>
<!--top end-->
<!--nav start-->
<div id="NavLink">
<div class="NavBG">
<!--Head Menu Start-->
<ul id='sddm'><li class='CurrentLi'><a href='/'>HOME</a></li> <li><a href='/About' onmouseover=mopen('m2') onmouseout='mclosetime()'>ABOUT</a> <div id='m2' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'><a href='/About'>Introduction</a> <a href='/About/Group'>Team</a> <a href='/About/Culture'>Culture</a> <a href='/About/Enviro'>Environment</a> <a href='/About/Business'>Business</a> </div></li> <li><a href='/news/' onmouseover=mopen('m3') onmouseout='mclosetime()'>NEWS</a> <div id='m3' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'><a href='/news/CompanyNews'>Company News</a> <a href='/news/IndustryNews'>Industry News</a> </div></li> <li><a href='/Product/' onmouseover=mopen('m4') onmouseout='mclosetime()'>PRODUCT</a> <div id='m4' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'><a href='/Product/DigitalPlayer'>Digital Player</a> <a href='/Product/Pad'>Tablet</a> <a href='/Product/GPS'>GPS</a> <a href='/Product/NoteBook'>NoteBook</a> <a href='/Product/Mobile'>Mobile</a> </div></li> <li><a href='/Support/' onmouseover=mopen('m5') onmouseout='mclosetime()'>SUPPORT</a> <div id='m5' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'><a href='/Support/Services'>Services</a> <a href='/Support/Download'>Download</a> <a href='/Support/FAQ'>Faq</a> </div></li> <li><a href='/Recruit/' onmouseover=mopen('m6') onmouseout='mclosetime()'>RECRUIT</a> <div id='m6' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'><a href='/recruit/peiyang'>Talented</a> <a href='/recruit/fuli'>Fuli</a> <a href='/recruit/jobs'>Jobs</a> </div></li> <li><a href='/contact/'>CONTACT</a></li> <li><a href='/Feedback/'>FEEDBACK</a></li> </ul>
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
<div class="topic"><div class='TopicTitle'>Contact Us</div></div>
<div class="txt ColorLink">
<p>CompanyName Group Co.,ltd</p>
<p>ADD：kawei District Guangzhou China</p>
<p>Tel：000-40324888</p>
<p>Fax：000-40324888</p>
<p>Web：<a href='http://www.huiguer.com' target='_blank'>www.huiguer.com</a></p>
<p>Email：admin@company.com</p>
<p align='center'><a href="http://wpa.qq.com/msgrd?v=3&uin=995226433&site=qq&menu=yes"><img src="/images/pa6.gif" alt='QQ'/></a>   <a href="http://wpa.qq.com/msgrd?v=3&uin=995226433&site=qq&menu=yes"><img src="/images/pa6.gif" alt='QQ'/></a></p>
</div>

</div>
<div class="HeightTab clearfix"></div>

</div>
<!--left end-->
<!--right start-->
<div class="right">
<div class="Position"><span>Position�<a href="/">Home</a> > Inquire</span></div>
<div class="HeightTab clearfix"></div>
<!--main start-->
<div class="main">
<!--FeedBack start-->
<div class="FeedBack">


<div class="commentbox">
<form id="form1" name="form1" method="post" action="/inc/order.asp?act=add&id=<%=a_id%>">
  <table id="commentform" width="600" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td>Product</td>
      <td><span class='OrderName'><%=ProductName%></span></td>
    </tr>  
    <tr>
      <td>Your name</td>
      <td><input name='name' type='text' id='name' size='30' maxlength="30"><span class="FontRed">*</span> </td>
    </tr>
    <tr>
      <td>Your address</td>
      <td><input name='address' type='text' id='address' size='30' maxlength="30"> </td>
    </tr>
    <tr>
      <td>Your tel</td>
      <td><input name='tel' type='text' id='tel' size='30' maxlength="30"> </td>
    </tr>    
    <tr>
      <td>Your Email</td>
      <td><input name='email' type='text' id='email' size='30' maxlength="80"><span class="FontRed">*</span></td>
    </tr>
    <tr>
      <td>Memo</td>
      <td>
        <textarea name="content" cols="60" rows="7"  value="" ></textarea>
           </td>    </tr>
    <tr>
      <td>Verify code</td>
      <td><input name="verycode"  maxLength=5 size=10 > <span class="FontRed">*</span><img src="/inc/getcode.asp" width="55"  onclick="this.src=this.src+'?'" alt="click for new code" style="cursor:hand;"></td>
    </tr>	
    <tr>
      <td> </td>
      <td><input class="Cbutton" type="submit" value=" SENT " onClick='javascript:return order_check()'></td>
    </tr>
  </table>
</form>
</div>

</div>
<!--FeedBack end-->




</div>
<!--main end-->
</div>
<!--right end-->
</div>
<!--inner end-->
<div class="clearfix"></div>
</div>
<!--body end-->
<div class="HeightTab clearfix"></div>
<!--footer start-->
<div id="footer">
<div class="inner">
<div class='BottomNav'><a href="/">HOME</a> | <a href="/About">INSTRUCTION</a> | <a href="/Contact">CONTACT US</a> | <a href="/Sitemap">SITEMAP</a></div>

<p>Copyright   2009-2013 Huiguer Network Studio  All rights reserved</p>

<p> <a href="http://www.huiguer.com/" target="_blank">Huiguer.com</a> Technology Support <a href="http://www.huiguer.com/" target="_blank">HuiguerCMS V3.6 UTF-8</a> <a href="/rss" target="_blank"><img src="/images/rss_icon.gif"></a> <a href="/rss/feed.xml" target="_blank"><img src="/images/xml_icon.gif"></a>

</p>

</div>
</div>
<!--footer end -->


</div>
<script type="text/javascript">
window.onerror=function(){return true;}
</script>
</body>
</html>
<!--
Powered By HituxCMS ASP V2.O   
-->






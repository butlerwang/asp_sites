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
<title>Require_<%=ProductName%></title>
<link href="/css/HituxCMSBlue/inner.css" rel="stylesheet" type="text/css" />
<link href="/css/HituxCMSBlue/style.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="/js/functions.js"></script>
<script type="text/javascript" src="/images/iepng/iepngfix_tilebg.js"></script>
</head>

<body>
<div class='OrderArea'>
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
      <td>Count</td>
      <td><input name='ordercount' type='text' id='ordercount' size='10' maxlength="10" value='1'><span class="FontRed">*</span></td>
    </tr>    
    <tr>
      <td>Contact Person</td>
      <td><input name='name' type='text' id='name' size='30' maxlength="30"><span class="FontRed">*</span></td>
    </tr>
    <tr>
      <td>Address</td>
      <td><input name='address' type='text' id='address' size='30' maxlength="30"><span class="FontRed">*</span></td>
    </tr>
    <tr>
      <td>Tel</td>
      <td><input name='tel' type='text' id='tel' size='30' maxlength="30"><span class="FontRed">*</span></td>
    </tr>    
    <tr>
      <td>Email</td>
      <td><input name='email' type='text' id='email' size='30' maxlength="80"></td>
    </tr>
    <tr>
      <td>QQ</td>
      <td><input name='qq' type='text' id='qq' size='30' maxlength="30"></td>
    </tr>	
    <tr>
      <td>Memo</td>
      <td>
        <textarea name="content" cols="60" rows="3"  value="" ></textarea>
           </td>    </tr>
    <tr>
      <td>Verify Code</td>
      <td><input name="verycode"  maxLength=5 size=10 > <span class="FontRed">*</span><img src="/inc/getcode.asp" width="55"  onclick="this.src=this.src+'?'" alt="click for new code" style="cursor:hand;"></td>
    </tr>	
    <tr>
      <td></td>
      <td><input class="Cbutton" type="submit" value=" Submit " onClick='javascript:return order_check()'></td>
    </tr>
  </table>
</form>
</div>

</div>
<!--FeedBack end-->

</div>
</body>
</html>
<!--
Powered By HituxCMS ASP V2.1   
-->
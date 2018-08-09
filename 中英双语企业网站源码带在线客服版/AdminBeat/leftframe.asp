<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!-- #include file="../inc/access.asp" -->
<!-- #include file="inc/functions.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="css/common.css" type="text/css" />
<title>左侧导航栏</title>
<script type="text/javascript">
window.onerror=function(){return true;}
</script></head>
<script language="JavaScript">
<!--
function key(){ 
if(event.shiftKey){
window.close();}
//禁止Shift
if(event.altKey){
window.close();}
//禁止Alt
if(event.ctrlKey){
window.close();}
//禁止Ctrl
return false;}
document.onkeydown=key;
if (window.Event)
document.captureEvents(Event.MOUSEUP);
function nocontextmenu(){
event.cancelBubble = true
event.returnValue = false;
return false;}
function norightclick(e){
if (window.Event){
if (e.which == 2 || e.which == 3)
return false;}
else
if (event.button == 2 || event.button == 3){
event.cancelBubble = true
event.returnValue = false;
return false;}
}
//禁右键
document.oncontextmenu = nocontextmenu;   // for IE5+
document.onmousedown = norightclick;   // for all others
//-->
</script>
<script language=JavaScript>
function logout(){
	if (confirm("您确定要退出后台管理系统吗？"))
	top.location = "logout.asp";
	return false;
}

</script>
<script  type="text/javascript" src="js/nav.js"></script>
<body onload="initinav('管理首页')">
<div id="left_content">
     <div id="user_info">欢迎您，<strong><%=session("log_name")%></strong><br /><a href="#"><%If logr() Then
		 response.write " 超级管理员 "
		else
		response.write " 普通管理员 "
		end if%></a> - <a href="#" target="_self" onClick="logout();">退出</a></div>
	 <div id="main_nav">
	     <div id="left_main_nav"></div>
		 <div id="right_main_nav"></div>
	 </div>
</div>
</body>

</html>

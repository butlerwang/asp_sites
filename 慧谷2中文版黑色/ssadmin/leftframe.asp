<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!-- #include file="../inc/access.asp" -->
<!-- #include file="inc/functions.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" href="css/common.css" type="text/css" />
<title>��ർ����</title>
</head>
<script language="JavaScript">
<!--
function key(){ 
if(event.shiftKey){
window.close();}
//��ֹShift
if(event.altKey){
window.close();}
//��ֹAlt
if(event.ctrlKey){
window.close();}
//��ֹCtrl
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
//���Ҽ�
document.oncontextmenu = nocontextmenu;   // for IE5+
document.onmousedown = norightclick;   // for all others
//-->
</script>
<script language=JavaScript>
function logout(){
	if (confirm("��ȷ��Ҫ�˳���̨����ϵͳ��"))
	top.location = "logout.asp";
	return false;
}

</script>
<script  type="text/javascript" src="js/nav.js"></script>
<body onload="initinav('������ҳ')">
<div id="left_content">
     <div id="user_info">��ӭ����<strong><%=session("log_name")%></strong><br /><a href="#"><%If logr() Then
		 response.write " ��������Ա "
		else
		response.write " ��ͨ����Ա "
		end if%></a> - <a href="#" target="_self" onClick="logout();">�˳�</a></div>
	 <div id="main_nav">
	     <div id="left_main_nav"></div>
		 <div id="right_main_nav"></div>
	 </div>
</div>
</body>
</html>

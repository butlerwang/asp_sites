<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!-- #include file="../inc/access.asp" -->
<!-- #include file="inc/functions.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" href="css/common.css" type="text/css" />
<title>企业网站管理系统</title>
</head>
<script language=JavaScript>
function logout(){
	if (confirm("您确定要退出后台管理系统吗？"))
	top.location = "logout.asp";
	return false;
}
</script>
<script  type="text/javascript">
var preClassName = "man_nav_1";
function list_sub_nav(Id,sortname){
   if(preClassName != ""){
      getObject(preClassName).className="bg_image";
   }
   if(getObject(Id).className == "bg_image"){
      getObject(Id).className="bg_image_onclick";
      preClassName = Id;
	  showInnerText(Id);
	  window.top.frames['leftFrame'].outlookbar.getbytitle(sortname);
	  window.top.frames['leftFrame'].outlookbar.getdefaultnav(sortname);
   }
}

function showInnerText(Id){
    var switchId = parseInt(Id.substring(8));
	var showText = "对不起没有信息！";
	switch(switchId){
	    case 1:
		   showText =  "欢迎进入企业网站管理系统!";
		   break;
	    case 2:
		   showText =  "进入系统后，首先在这里对网站进行配置，包含网站各称、网址、首页轮播、在线客服、友情链接等信息。";
		   break;
	    case 3:
		   showText =  "管理前台导航，添加、删除、修改导航，设置导航和栏目的对应关系，或者设置导航的外部链接，完成操作后请重新生成所有静态。";
		   break;		   
	    case 4:
		   showText =  "管理栏目、内容、订单、留言，如新闻、产品、招聘等内容。";
		   break;	
	    case 5:
		   showText =  "高级设置可以更换主题及对模板进行修改。";
		   break;		   		   
	    case 6:
		   showText =  "这里可将全站生成静态页面，请选择相应或内容进行生成。";
		   break;
	    case 7:
		   showText =  "进入系统后，首先在这里对网站进行配置，包含网站各称、网址、首页轮播、在线客服、友情链接等信息。";
		   break;
	    case 8:
		   showText =  "管理前台导航，添加、删除、修改导航，设置导航和栏目的对应关系，或者设置导航的外部链接，完成操作后请重新生成所有静态。";
		   break;		   
	    case 9:
		   showText =  "在这里添加栏目下的文章，如新闻、产品、招聘等内容。";		   		}
	getObject('show_text').innerHTML = showText;
}
 //获取对象属性兼容方法
 function getObject(objectId) {
    if(document.getElementById && document.getElementById(objectId)) {
	// W3C DOM
	return document.getElementById(objectId);
    } else if (document.all && document.all(objectId)) {
	// MSIE 4 DOM
	return document.all(objectId);
    } else if (document.layers && document.layers[objectId]) {
	// NN 4 DOM.. note: this won't find nested layers
	return document.layers[objectId];
    } else {
	return false;
    }
}
</script>
<body>
<div id="nav">
    <ul>
    <li id="man_nav_1" onclick="list_sub_nav(id,'管理首页')"  class="bg_image_onclick">管理首页</li>
    
<%If logr() Then %>    
    <li id="man_nav_2" onclick="list_sub_nav(id,'系统设置')"  class="bg_image">系统设置</li>
<%End If %>

	<li id="man_nav_3" onclick="list_sub_nav(id,'导航管理')"  class="bg_image">导航管理</li>
    <li id="man_nav_4" onclick="list_sub_nav(id,'内容管理')"  class="bg_image">内容管理</li>

    <li id="man_nav_6" onclick="list_sub_nav(id,'静态管理')"  class="bg_image">静态管理</li>
    </ul>
</div>
<div id="sub_info">&nbsp;&nbsp;<img src="images/hi.gif" />&nbsp;<span id="show_text">欢迎进入 <strong><%=gdb("select web_name from web_settings ")%></strong> 网站后台管理系统 !</span></div>
</body>
</html>

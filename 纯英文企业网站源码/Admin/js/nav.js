var preClassName = ""; 
function list_sub_detail(Id, item) 
{ 
if(preClassName != "") 
{ 
getObject(preClassName).className = "left_back" 
} 
if(getObject(Id).className == "left_back") 
{ 
getObject(Id).className = "left_back_onclick"; 
outlookbar.getbyitem(item); 
preClassName = Id 
} 
} 
function getObject(objectId) 
{ 
if(document.getElementById && document.getElementById(objectId)) 
{ 
return document.getElementById(objectId) 
} 
else if(document.all && document.all(objectId)) 
{ 
return document.all(objectId) 
} 
else if(document.layers && document.layers[objectId]) 
{ 
return document.layers[objectId] 
} 
else 
{ 
return false 
} 
} 
function outlook() 
{ 
this.titlelist = new Array(); 
this.itemlist = new Array(); 
this.addtitle = addtitle; 
this.additem = additem; 
this.getbytitle = getbytitle; 
this.getbyitem = getbyitem; 
this.getdefaultnav = getdefaultnav 
} 
function theitem(intitle, insort, inkey, inisdefault) 
{ 
this.sortname = insort; 
this.key = inkey; 
this.title = intitle; 
this.isdefault = inisdefault 
} 
function addtitle(intitle, sortname, inisdefault) 
{ 
outlookbar.itemlist[outlookbar.titlelist.length] = new Array(); 
outlookbar.titlelist[outlookbar.titlelist.length] = new theitem(intitle, sortname, 0, inisdefault); 
return(outlookbar.titlelist.length - 1) 
} 
function additem(intitle, parentid, inkey) 
{ 
if(parentid >= 0 && parentid <= outlookbar.titlelist.length) 
{ 
insort = "item_" + parentid; 
outlookbar.itemlist[parentid][outlookbar.itemlist[parentid].length] = new theitem(intitle, insort, inkey, 0); 
return(outlookbar.itemlist[parentid].length - 1) 
} 
else additem = - 1 
} 
function getdefaultnav(sortname) 
{ 
var output = ""; 
for(i = 0; i < outlookbar.titlelist.length; i ++ ) 
{ 
if(outlookbar.titlelist[i].isdefault == 1 && outlookbar.titlelist[i].sortname == sortname) 
{ 
output += "<div class=list_tilte id=sub_sort_" + i + " onclick=\"hideorshow('sub_detail_"+i+"')\">"; 
output += "<span>" + outlookbar.titlelist[i].title + "</span>"; 
output += "</div>"; 
output += "<div class=list_detail id=sub_detail_" + i + "><ul>"; 
for(j = 0; j < outlookbar.itemlist[i].length; j ++ ) 
{ 
output += "<li id=" + outlookbar.itemlist[i][j].sortname + j + " onclick=\"changeframe('"+outlookbar.itemlist[i][j].title+"', '"+outlookbar.titlelist[i].title+"', '"+outlookbar.itemlist[i][j].key+"')\"><a href=#>" + outlookbar.itemlist[i][j].title + "</a></li>" 
} 
output += "</ul></div>" 
} 
} 
getObject('right_main_nav').innerHTML = output 
} 
function getbytitle(sortname) 
{ 
var output = "<ul>"; 
for(i = 0; i < outlookbar.titlelist.length; i ++ ) 
{ 
if(outlookbar.titlelist[i].sortname == sortname) 
{ 
output += "<li id=left_nav_" + i + " onclick=\"list_sub_detail(id, '"+outlookbar.titlelist[i].title+"')\" class=left_back>" + outlookbar.titlelist[i].title + "</li>" 
} 
} 
output += "</ul>"; 
getObject('left_main_nav').innerHTML = output 
} 
function getbyitem(item) 
{ 
var output = ""; 
for(i = 0; i < outlookbar.titlelist.length; i ++ ) 
{ 
if(outlookbar.titlelist[i].title == item) 
{ 
output = "<div class=list_tilte id=sub_sort_" + i + " onclick=\"hideorshow('sub_detail_"+i+"')\">"; 
output += "<span>" + outlookbar.titlelist[i].title + "</span>"; 
output += "</div>"; 
output += "<div class=list_detail id=sub_detail_" + i + " style='display:block;'><ul>"; 
for(j = 0; j < outlookbar.itemlist[i].length; j ++ ) 
{ 
output += "<li id=" + outlookbar.itemlist[i][j].sortname + "_" + j + " onclick=\"changeframe('"+outlookbar.itemlist[i][j].title+"', '"+outlookbar.titlelist[i].title+"', '"+outlookbar.itemlist[i][j].key+"')\"><a href=#>" + outlookbar.itemlist[i][j].title + "</a></li>" 
} 
output += "</ul></div>" 
} 
} 
getObject('right_main_nav').innerHTML = output 
} 
function changeframe(item, sortname, src) 
{ 
if(item != "" && sortname != "") 
{ 
window.top.frames['mainFrame'].getObject('show_text').innerHTML = sortname + "  <img src=images/slide.gif broder=0 />  " + item 
} 
if(src != "") 
{ 
window.top.frames['manFrame'].location = src 
} 
} 
function hideorshow(divid) 
{ 
subsortid = "sub_sort_" + divid.substring(11); 
if(getObject(divid).style.display == "none") 
{ 
getObject(divid).style.display = "block"; 
getObject(subsortid).className = "list_tilte" 
} 
else 
{ 
getObject(divid).style.display = "none"; 
getObject(subsortid).className = "list_tilte_onclick" 
} 
} 
function initinav(sortname) 
{ 
outlookbar.getdefaultnav(sortname); 
outlookbar.getbytitle(sortname); 
//window.top.frames['manFrame'].location = "manFrame.html" 
}

// 导航栏配置文件
var outlookbar=new outlook();
var t;
t=outlookbar.addtitle('系统信息','管理首页',1)
outlookbar.additem('系统检测',t,'start.asp')
outlookbar.additem('版权信息',t,'version.asp')

t=outlookbar.addtitle('网站设置','系统设置',1)
outlookbar.additem('网站信息设置',t,'web_settings.asp')

t=outlookbar.addtitle('导航管理','系统设置',2)
outlookbar.additem('添加一级导航',t,'menu_type_add.asp')
outlookbar.additem('一级导航列表',t,'menu_type_list.asp')
outlookbar.additem('添加二级导航',t,'menu_add.asp')
outlookbar.additem('二级导航列表',t,'menu_list.asp')

t=outlookbar.addtitle('广告管理','系统设置',3)
outlookbar.additem('广告位置管理',t,'ads_position_list.asp')
outlookbar.additem('添加广告',t,'ads_add.asp')
outlookbar.additem('广告列表',t,'ads_list.asp')

t=outlookbar.addtitle('友情链接','系统设置',4)
outlookbar.additem('添加链接',t,'link_add.asp')
outlookbar.additem('链接列表',t,'link_list.asp')

t=outlookbar.addtitle('后台用户','系统设置',5)
outlookbar.additem('添加用户',t,'admin_add.asp')
outlookbar.additem('用户列表',t,'admin_list.asp')


t=outlookbar.addtitle('栏目管理','栏目管理',1)
outlookbar.additem('添加一级栏目',t,'category_add.asp?ppid=1')
outlookbar.additem('栏目列表',t,'category_list.asp')

t=outlookbar.addtitle('留言管理','栏目管理',2)
outlookbar.additem('留言列表',t,'message_list.asp')


t=outlookbar.addtitle('文章管理','内容管理',1)
outlookbar.additem('添加文章',t,'article_add.asp')
outlookbar.additem('文章列表',t,'article_list.asp')
outlookbar.additem('文章关键词',t,'keywords_list.asp')
outlookbar.additem('文章来源',t,'author_list.asp')

t=outlookbar.addtitle('产品管理','内容管理',2)
outlookbar.additem('添加产品',t,'product_add.asp')
outlookbar.additem('产品列表',t,'product_list.asp')
outlookbar.additem('订单列表',t,'order_list.asp')

t=outlookbar.addtitle('招聘管理','内容管理',3)
outlookbar.additem('添加招聘职位',t,'info_add.asp')
outlookbar.additem('招聘职位列表',t,'info_list.asp')


t=outlookbar.addtitle('数据管理','内容管理',4)
outlookbar.additem('备份数据库',t,'data_backup.asp')
outlookbar.additem('还原数据库',t,'data_restore.asp')
outlookbar.additem('备份数据列表',t,'data_list.asp')


t=outlookbar.addtitle('主题管理','高级设置',1)
outlookbar.additem('添加新主题',t,'theme_add.asp')
outlookbar.additem('主题列表',t,'themesetting.asp')

t=outlookbar.addtitle('模板分类','高级设置',2)
outlookbar.additem('添加新模板分类',t,'models_type_add.asp')
outlookbar.additem('模板分类列表',t,'models_type_list.asp')

t=outlookbar.addtitle('模板管理','高级设置',3)
outlookbar.additem('添加新模板',t,'web_models_add.asp')
outlookbar.additem('模板列表',t,'web_models.asp')


t=outlookbar.addtitle('生成首页','静态管理',1)
outlookbar.additem('生成首页',t,'html_index.asp')
outlookbar.additem('生成所有页面',t,'html_all.asp')

t=outlookbar.addtitle('生成栏目','静态管理',2)
outlookbar.additem('生成栏目',t,'html_items.asp')

t=outlookbar.addtitle('生成内容','静态管理',3)
outlookbar.additem('生成内容',t,'html_article.asp')

t=outlookbar.addtitle('生成所有','静态管理',4)
outlookbar.additem('生成所有',t,'html_all.asp')










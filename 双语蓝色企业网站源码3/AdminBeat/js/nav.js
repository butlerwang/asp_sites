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

// �����������ļ�
var outlookbar=new outlook();
var t;
t=outlookbar.addtitle('ϵͳ��Ϣ','������ҳ',1)
outlookbar.additem('ϵͳ���',t,'start.asp')
outlookbar.additem('��Ȩ��Ϣ',t,'version.asp')

t=outlookbar.addtitle('�������','������ҳ',2)
outlookbar.additem('���������',t,'theme_add.asp')
outlookbar.additem('�����б�',t,'themesetting.asp')

t=outlookbar.addtitle('��̨�û�','������ҳ',3)
outlookbar.additem('����û�',t,'admin_add.asp')
outlookbar.additem('�û��б�',t,'admin_list.asp')

t=outlookbar.addtitle('���ݹ���','������ҳ',4)
outlookbar.additem('�������ݿ�',t,'data_backup.asp')
outlookbar.additem('��ԭ���ݿ�',t,'data_restore.asp')
outlookbar.additem('���������б�',t,'data_list.asp')

t=outlookbar.addtitle('����һ��վ����','����һϵͳ����',1)
outlookbar.additem('��վ��Ϣ����',t,'web_settings.asp')

t=outlookbar.addtitle('����һ��������','����һϵͳ����',2)
outlookbar.additem('���һ������',t,'menu_type_add.asp')
outlookbar.additem('һ�������б�',t,'menu_type_list.asp')
outlookbar.additem('��Ӷ�������',t,'menu_add.asp')
outlookbar.additem('���������б�',t,'menu_list.asp')

t=outlookbar.addtitle('����һ������','����һϵͳ����',3)
outlookbar.additem('���λ�ù���',t,'ads_position_list.asp')
outlookbar.additem('��ӹ��',t,'ads_add.asp')
outlookbar.additem('����б�',t,'ads_list.asp')

t=outlookbar.addtitle('����һ��������','����һϵͳ����',4)
outlookbar.additem('�������',t,'link_add.asp')
outlookbar.additem('�����б�',t,'link_list.asp')


t=outlookbar.addtitle('����һ��Ŀ����','��Ŀ����',1)
outlookbar.additem('���һ����Ŀ',t,'category_add.asp?ppid=1')
outlookbar.additem('��Ŀ�б�',t,'category_list.asp')

t=outlookbar.addtitle('����һ���Թ���','��Ŀ����',2)
outlookbar.additem('�����б�',t,'message_list.asp')

t=outlookbar.addtitle('Ӣ��һ��Ŀ����','��Ŀ����',3)
outlookbar.additem('���һ����Ŀ',t,'en_category_add.asp?ppid=1')
outlookbar.additem('��Ŀ�б�',t,'en_category_list.asp')

t=outlookbar.addtitle('Ӣ��һ���Թ���','��Ŀ����',4)
outlookbar.additem('�����б�',t,'en_message_list.asp')

t=outlookbar.addtitle('����һ���¹���','����һ���ݹ���',1)
outlookbar.additem('�������',t,'article_add.asp')
outlookbar.additem('�����б�',t,'article_list.asp')
outlookbar.additem('���¹ؼ���',t,'keywords_list.asp')
outlookbar.additem('������Դ',t,'author_list.asp')

t=outlookbar.addtitle('����һ��Ʒ����','����һ���ݹ���',2)
outlookbar.additem('��Ӳ�Ʒ',t,'product_add.asp')
outlookbar.additem('��Ʒ�б�',t,'product_list.asp')
outlookbar.additem('��������',t,'order_list.asp')

t=outlookbar.addtitle('����һ��Ƹ����','����һ���ݹ���',3)
outlookbar.additem('�����Ƹְλ',t,'info_add.asp')
outlookbar.additem('��Ƹְλ�б�',t,'info_list.asp')




t=outlookbar.addtitle('Ӣ��һ��վ����','Ӣ��һϵͳ����',1)
outlookbar.additem('��վ��Ϣ����',t,'en_web_settings.asp')

t=outlookbar.addtitle('Ӣ��һ��������','Ӣ��һϵͳ����',2)
outlookbar.additem('���һ������',t,'en_menu_type_add.asp')
outlookbar.additem('һ�������б�',t,'en_menu_type_list.asp')
outlookbar.additem('��Ӷ�������',t,'en_menu_add.asp')
outlookbar.additem('���������б�',t,'en_menu_list.asp')

t=outlookbar.addtitle('Ӣ��һ������','Ӣ��һϵͳ����',3)
outlookbar.additem('���λ�ù���',t,'en_ads_position_list.asp')
outlookbar.additem('��ӹ��',t,'en_ads_add.asp')
outlookbar.additem('����б�',t,'en_ads_list.asp')

t=outlookbar.addtitle('Ӣ��һ��������','Ӣ��һϵͳ����',4)
outlookbar.additem('�������',t,'en_link_add.asp')
outlookbar.additem('�����б�',t,'en_link_list.asp')


t=outlookbar.addtitle('Ӣ��һ���¹���','Ӣ��һ���ݹ���',1)
outlookbar.additem('�������',t,'en_article_add.asp')
outlookbar.additem('�����б�',t,'en_article_list.asp')
outlookbar.additem('���¹ؼ���',t,'en_keywords_list.asp')
outlookbar.additem('������Դ',t,'en_author_list.asp')

t=outlookbar.addtitle('Ӣ��һ��Ʒ����','Ӣ��һ���ݹ���',2)
outlookbar.additem('��Ӳ�Ʒ',t,'en_product_add.asp')
outlookbar.additem('��Ʒ�б�',t,'en_product_list.asp')
outlookbar.additem('��������',t,'en_order_list.asp')

t=outlookbar.addtitle('Ӣ��һ��Ƹ����','Ӣ��һ���ݹ���',3)
outlookbar.additem('�����Ƹְλ',t,'en_info_add.asp')
outlookbar.additem('��Ƹְλ�б�',t,'en_info_list.asp')


t=outlookbar.addtitle('�������','�߼�����',1)
outlookbar.additem('���������',t,'theme_add.asp')
outlookbar.additem('�����б�',t,'themesetting.asp')

t=outlookbar.addtitle('ģ�����','�߼�����',2)
outlookbar.additem('�����ģ�����',t,'models_type_add.asp')
outlookbar.additem('ģ������б�',t,'models_type_list.asp')

t=outlookbar.addtitle('ģ�����','�߼�����',3)
outlookbar.additem('�����ģ��',t,'web_models_add.asp')
outlookbar.additem('ģ���б�',t,'web_models.asp')


t=outlookbar.addtitle('������ҳ','��̬����',1)
outlookbar.additem('����������ҳ',t,'html_index.asp')
outlookbar.additem('����Ӣ����ҳ',t,'en_html_index.asp')

t=outlookbar.addtitle('������Ŀ','��̬����',2)
outlookbar.additem('����������Ŀ',t,'html_items.asp')
outlookbar.additem('����Ӣ����Ŀ',t,'en_html_items.asp')

t=outlookbar.addtitle('��������','��̬����',3)
outlookbar.additem('������������',t,'html_article.asp')
outlookbar.additem('����Ӣ������',t,'en_html_article.asp')

t=outlookbar.addtitle('��������','��̬����',4)
outlookbar.additem('������������',t,'html_all.asp')
outlookbar.additem('����Ӣ������',t,'en_html_all.asp')










<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!-- #include file="../inc/access.asp" -->
<!-- #include file="inc/functions.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" href="css/common.css" type="text/css" />
<title>������ҵ��վ����ϵͳ UTF-8</title>
</head>
<script language=JavaScript>
function logout(){
	if (confirm("��ȷ��Ҫ�˳���̨����ϵͳ��"))
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
	var showText = "�Բ���û����Ϣ��";
	switch(switchId){
	    case 1:
		   showText =  "��ӭ����������ҵ��վ����ϵͳ UTF-8!";
		   break;
	    case 2:
		   showText =  "����ϵͳ���������������վ�������ã�������վ���ơ���ַ����������桢�������ӵ���Ϣ��";
		   break;
	    case 3:
		   showText =  "��Ŀ�������������ӣ��޸ģ�ɾ����Ŀ�������㽨���ʺ��Լ�����Ŀ��";
		   break;		   
	    case 4:
		   showText =  "�����������Ŀ�µ����£������š���Ʒ����Ƹ�����ݡ�";
		   break;	
	    case 5:
		   showText =  "�߼����ÿ��Ը������⼰��ģ������޸ġ�";
		   break;		   		   
	    case 6:
		   showText =  "����ɽ�ȫվ���ɾ�̬ҳ�棬��ѡ����Ӧ�����ݽ������ɡ�";
		   break;		}
	getObject('show_text').innerHTML = showText;
}
 //��ȡ�������Լ��ݷ���
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
    <li id="man_nav_1" onclick="list_sub_nav(id,'������ҳ')"  class="bg_image_onclick">������ҳ</li>
      <%If logr() Then %>    
    <li id="man_nav_2" onclick="list_sub_nav(id,'ϵͳ����')"  class="bg_image">ϵͳ����</li>
<%End If %>
    <li id="man_nav_3" onclick="list_sub_nav(id,'��Ŀ����')"  class="bg_image">��Ŀ����</li>
    <li id="man_nav_4"  onclick="list_sub_nav(id,'���ݹ���')"  class="bg_image">���ݹ���</li>
      <%If logr() Then %>        
    <li id="man_nav_5"  onclick="list_sub_nav(id,'�߼�����')"  class="bg_image">�߼�����</li>
<%End If %>
    
    <li id="man_nav_6"  onclick="list_sub_nav(id,'��̬����')"  class="bg_image">��̬����</li>
    </ul>
</div>
<div id="sub_info">&nbsp;&nbsp;<img src="images/hi.gif" />&nbsp;<span id="show_text">��ӭ���� <strong><%=gdb("select web_name from web_settings ")%></strong> ��վ��̨����ϵͳ !</span></div>
</body>
</html>

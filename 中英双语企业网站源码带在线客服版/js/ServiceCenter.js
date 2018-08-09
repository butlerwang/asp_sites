document.writeln("<div class=qqbox id=divQQbox style=\"TOP: 591px\">");
document.writeln("<div class=qqlv id=meumid onmouseover=show() style=\"DISPLAY: block\"><IMG src=\"\/images\/serviceimg\/qqbg.gif\"><\/div>");
document.writeln("<div class=qqkf id=contentid style=\"DISPLAY: none\" onmouseout=hideMsgBox(event)>");
document.writeln("<div class=qqkfbt id=qq-1 onfocus=this.blur(); onmouseout=\"showandhide(\'qq-\',\'qqkfbt\',\'qqkfbt\',\'K\',1,1);\">客 服 中 心<\/div>");
document.writeln("<div id=K1 style=\"PADDING-RIGHT: 0px; PADDING-LEFT: 0px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px\" align=center margin-left=\"5px\">")
document.writeln("<p>客服1: <a target=\'_blank\' href=\'tencent:\/\/message\/?uin=2216935501&Site=www.hitux.com&Menu=yes\'><img  SRC=\'http:\/\/wpa.qq.com\/pa?p=1:2216935501:1\' alt=\'在线咨询\'><\/a><\/p>"); document.writeln("<p>客服2: <a target=\'_blank\' href=\'tencent:\/\/message\/?uin=2216935501&Site=www.hitux.com&Menu=yes\'><img  SRC=\'http:\/\/wpa.qq.com\/pa?p=1:2216935501:1\' alt=\'在线咨询\'><\/a><\/p>"); document.writeln("<p>客服3: <a target=\'_blank\' href=\'tencent:\/\/message\/?uin=2216935501&Site=www.hitux.com&Menu=yes\'><img  SRC=\'http:\/\/wpa.qq.com\/pa?p=1:2216935501:1\' alt=\'在线咨询\'><\/a><\/p>"); 
document.writeln("<p>旺旺: <a target=\'_blank\' href=\'http://www.taobao.com/webww/ww.php?ver=3&touid=hitux&siteid=cntaobao&status=1&charset=utf-8\'><img  SRC=\'/images/serviceimg/wang_icon.gif\' alt=\'在线咨询\'><\/a><\/p>"); 


document.writeln("<\/div><\/div><\/div>")
function showandhide(h_id,hon_class,hout_class,c_id,totalnumber,activeno) {
var h_id,hon_id,hout_id,c_id,totalnumber,activeno;
for (var i=1;i<=totalnumber;i++) {
document.getElementById(c_id+i).style.display='none';
document.getElementById(h_id+i).className=hout_class;
}
document.getElementById(c_id+activeno).style.display='block';
document.getElementById(h_id+activeno).className=hon_class;
}
var tips; 
var theTop = 170;
var old = theTop;
function initFloatTips() 
{ 
tips = document.getElementById('divQQbox');
moveTips();
}
function moveTips()
{
var tt=50; 
if (window.innerHeight) 
{
pos = window.pageYOffset 
}else if (document.documentElement && document.documentElement.scrollTop) {
pos = document.documentElement.scrollTop 
}else if (document.body) {
pos = document.body.scrollTop; 
}
pos=pos-tips.offsetTop+theTop; 
pos=tips.offsetTop+pos/10; 
if (pos < theTop){
pos = theTop;
}
if (pos != old) { 
tips.style.top = pos+"px";
tt=10; //alert(tips.style.top); 
}
old = pos;
setTimeout(moveTips,tt);
}
initFloatTips();
if(typeof(HTMLElement)!="undefined") //firefox定义contains()方法，ie下不起作用
{ 
HTMLElement.prototype.contains=function (obj) 
{ 
while(obj!=null&&typeof(obj.tagName)!="undefind"){
if(obj==this) return true; 
obj=obj.parentNode;
} 
return false; 
}
}
function show()
{
document.getElementById("meumid").style.display="none"
document.getElementById("contentid").style.display="block"
}
function hideMsgBox(theEvent){
if (theEvent){
var browser=navigator.userAgent;
if (browser.indexOf("Firefox")>0){ //如果是Firefox
if (document.getElementById("contentid").contains(theEvent.relatedTarget)) {
return
}
}
if (browser.indexOf("MSIE")>0 || browser.indexOf("Presto")>=0){
if (document.getElementById('contentid').contains(event.toElement)) {
return; 
}
}
}
document.getElementById("meumid").style.display = "block";
document.getElementById("contentid").style.display = "none";
}

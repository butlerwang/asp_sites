
function Alert(src,msg,act)
{ 
    var p=new KesionPopup();
	p.MsgBorder=1;
	p.ShowBackground=false;
	p.BgColor='#fff';
	p.TitleCss="font-size:14px;background:#1B76B7;color:#fff;height:22px;";
    p.popup("<span style='color:#fff'>提示</span>","<div style='margin:30px;'><table><tr><td><img src='/images/default/"+src+".png' align='left'></td><td>"+msg+"</td></tr></table></div><div style='text-align:center;margin:10px'><input type='button' onclick='"+act+";closeWindow()' value=' 确 定 ' class='btn'></div>",350);
}
var fadeOutTime=0
function KesionPopup(BgColor,MsgBorder,ShowBackground,TitleCss)
{
 this.QuoteByAdmin=false;  
 this.FadeInTime=0;        //淡入效果时间，不想有效果输入0
 this.FadeOutTime=0;       //淡出效果时间，不想有效果输入0
 this.ShowClose=true;     //显示关闭
 this.BgColor=BgColor==null?"#DAEAEE":BgColor;
 this.MsgBorder=MsgBorder==null?9:MsgBorder;
 this.TitleCss=TitleCss==null?'':TitleCss;
 this.ShowBackground=ShowBackground==null?true:ShowBackground;
 this.PopupImgDir='';
  //在鼠标处弹出文字提示
 this.mousepopup=function(title,content,width){  
       var ev=getEvent();
	   var objPos = mousePosition(ev);
	   if (document.all){objPos.y=objPos.y+document.documentElement.scrollTop};
	   this.showMessageBox('<div style="padding:5px">'+title+'</div>','<div style="padding:10px">'+content+'</div>',objPos,width);
  }
  //在鼠标处弹出简单的文字提示,并取消边框,背景色等
  this.mousepop=function(title,content,width){
	   this.BgColor='#fff';
	   this.MsgBorder=2;
	   this.ShowBackground=false;
	   this.FadeInTime=this.FadeOutTime=800;
	   this.mousepopup(title,content,width);  
  }
//在鼠标处弹出iframe提示
this.mousePopupIframe=function(title,url,width,height,scrollType){   
  var ev=getEvent();
  var objPos = mousePosition(ev);
  var text="<iframe id='myframe' name='myframe' marginwidth='0' marginheight='0' width='99%' style='height:"+height+"px'  src='"+url+"' frameborder='0' scrolling='"+scrollType+"'></iframe>";
   this.showMessageBox(title,text,objPos,width);
   setTimeout(function(){myframe.location.href=url;},10);
}
//居中弹出文字提示
 this.popupTips=function(title,content,width,height){  
   this.QuoteByAdmin=true;
   var objPos = middlePosition(width);
   this.showMessageBox(title,content,objPos,width);
 }
//后台居中弹出
 this.PopupCenterIframe=function(title,url,width,height,scrollType){ 
  this.QuoteByAdmin=true;
  onscrolls=false;
  var objPos = Position(width,height);
  var text="<iframe id='myframe' name='myframe' marginwidth='0' marginheight='0' width='99%' style='height:"+height+"px'  src='about:blank' frameborder='0' scrolling='"+scrollType+"'></iframe>";
  this.showMessageBox(title,text,objPos,width);
  setTimeout(function(){myframe.location.href=url;myframe.focus();},10);
  
 }
//居中初始化弹窗 参数:title 标题 content内容 width 宽度
this.popupIframe=function(title,url,width,height,scrollType){ 
  var objPos = middlePosition(width);
  var text="<iframe id='myframe' name='myframe' width='100%' style='height:"+height+"px'  src='about:blank' frameborder='0' scrolling='"+scrollType+"'></iframe>";
  this.showMessageBox(title,text,objPos,width);
  setTimeout(function(){myframe.location.href=url;myframe.focus();},10);
}
//居中弹出普通窗口 参数:title 标题 content内容 width 宽度
this.popup=function(title,content,width){ 
  var objPos = middlePosition(width);
  this.showMessageBox(title,content,objPos,width);
}
 //弹出方法
 this.showMessageBox=function(wTitle,content,pos,wWidth)
	{
	   closeWindow();
       fadeOutTime=this.FadeOutTime;
	   var mesWindowCss="border:#3D5480 "+this.MsgBorder+"px solid;background:"+this.BgColor+";" //弹窗边框
	   if (this.QuoteByAdmin==true){
		   mesWindowCss="border:#000000 1px solid;background:#F1F6F9;" //弹窗边框
	   }
	   var bWidth=parseInt(document.documentElement.scrollWidth);
	   var bHeight=parseInt(document.body.offsetHeight);
	   bWidth=parseInt(document.body.scrollWidth);
	  // if (bHeight<parseInt(document.body.scrollHeight)) bHeight=parseInt(document.body.scrollHeight);
	   
	   if (this.ShowBackground==true){
	   if(isIe){ setSelectState('hidden');}
	   var back=document.createElement("div");
	   back.id="back";
	   var styleStr="top:0px;left:0px;position:absolute;background:#666;width:"+bWidth+"px;height:"+bHeight+"px;";
	   styleStr+=(isIe)?"filter:alpha(opacity=8);":"opacity:0.10;";
	   back.style.cssText=styleStr;
	   document.body.appendChild(back);
	   }
	   var mesW=document.createElement("div");
	   mesW.id="mesWindow";
	   //mesW.className="mesWindow";
	   if (this.QuoteByAdmin==true){
			  mesW.innerHTML="<div style='border-bottom:#eee 1px solid;font-weight:bold;text-align:left;font-size:12px;'><table cellpadding='0' cellspacing='0' bgcolor='#CFE0EA' background='"+this.PopupImgDir+"images/menu_bg.gif' width='100%' height='100%'><tr onmousedown=MDown(mesWindow)><td style='text-align:left;padding-left:10px;' height='28'><b>"+wTitle+"</b></td><td align='center' width='80'><span style='cursor:pointer' id='closeWindow' onclick='closeWindow();' title='关闭窗口' class='close'><img src='"+this.PopupImgDir+"../images/default/close.gif' border='0'> 关闭</span> </td></tr></table></div><div style='_margin:4px;font-size:12px;' id='mesWindowContent'>"+content+"</div><div class='mesWindowBottom'></div>";
	
	   }else{
		   if (this.ShowClose==true){
			mesW.innerHTML="<div style='border-bottom:#eee 1px solid;padding:3px;font-weight:bold;text-align:left;font-size:12px;"+this.TitleCss+"'><table cellpadding='0' cellspacing='0' width='100%' height='100%'><tr><td><b>"+wTitle+"</b></td><td style='width:41px;'><a href='javascript:closeWindow();' id='closeBtn' style='color:#ff6600'>X关闭</a></td></tr></table></div><div style='font-size:12px;' id='mesWindowContent'>"+content+"</div><div class='mesWindowBottom'></div>";
		   }else{
			mesW.innerHTML="<div style='margin:4px;font-size:12px;' id='mesWindowContent'>"+content+"</div>";
		   }
	
	   }
	   styleStr=mesWindowCss+"left:"+(((pos.x-wWidth)>0)?(pos.x-wWidth):pos.x)+"px;top:"+(pos.y)+"px;position:absolute;width:"+wWidth+"px;";
	   mesW.style.cssText=styleStr;
	   document.body.appendChild(mesW);
	   if (parseInt(this.FadeInTime)>0){$('#mesWindow').hide().fadeIn(this.FadeInTime);}
	};
	
	showBackground=function(obj,endInt)
	{
	   obj.filters.alpha.opacity+=1;
	   if(obj.filters.alpha.opacity<endInt)
	   {
		setTimeout(function(){showBackground(obj,endInt)},8);
	   }
	};
}
var isIe=(document.all)?true:false;
function Position(width,height)
{  
	if (isIe){
	 var y=document.body.offsetHeight/2-height/2-20;
	 if (y<0) y=20;
     return {x:document.body.offsetWidth/2-width/2, y:y};
	}
	else{
	  var h=window.screen.height;	
	  if (h==undefined){
		  h=height/2-150;
	  }else{h=h/2-height/2-150;}
     return {x:document.documentElement.scrollWidth/2-width/2, y:h};
	}
}
function getEvent(){ //同时兼容ie和ff的写法 
	if(window.event!=null) return window.event; 
	func=getEvent.caller; 
	while(func!=null){ 
	var arg0=func.arguments[0]; 
	if(arg0){ 
	 if((arg0.constructor==Event || arg0.constructor ==MouseEvent) || (typeof(arg0)=="object" && arg0.preventDefault && arg0.stopPropagation)){return arg0; } 
	} 
	func=func.caller; 
	} 
	return null; 
} 
function closeWindow()
{	  if(document.getElementById('back')!=null){
		document.getElementById('back').parentNode.removeChild(document.getElementById('back'));
	   }
	   if(document.getElementById('mesWindow')!=null){
		if (parseInt(fadeOutTime)>0){
		$('#mesWindow').fadeOut(fadeOutTime);
		setTimeout(function(){document.getElementById('mesWindow').parentNode.removeChild(document.getElementById('mesWindow'));},fadeOutTime);
		}else{
		 document.getElementById('mesWindow').parentNode.removeChild(document.getElementById('mesWindow'));
	   }
	   }
	  if(isIe){ setSelectState('');}
} 
function setSelectState(state){var objl=document.getElementsByTagName('select');for(var i=0;i<objl.length;i++){objl[i].style.visibility=state;}} 
function mousePosition(ev){if(ev.pageX || ev.pageY){return {x:ev.pageX, y:ev.pageY};}return {x:ev.clientX + document.body.scrollLeft - document.body.clientLeft,y:ev.clientY + document.body.scrollTop - document.body.clientTop};}
var popTopHeight=200; //离顶部位移像素
function middlePosition(width)
{ 
    var left = parseInt((document.body.clientWidth/2+width/2));//屏幕居中
    var top = document.body.scrollTop;
	if (top==0) top=document.documentElement.scrollTop;
    top=top+popTopHeight;
    return {x:left, y:top};
}
var onscrolls=true
window.onscroll = function(){
	 if (!onscrolls) return;
	 try{
	 var top=document.body.scrollTop;
	 if (top==0) top=document.documentElement.scrollTop;
	 document.getElementById("mesWindow").style.top=top+200;}
	 catch(e){}
}
var Obj=''
document.onmouseup=MUp
document.onmousemove=MMove

function MDown(Object){Obj=Object.id;document.all(Obj).setCapture();pX=event.x-document.all(Obj).style.pixelLeft;pY=event.y-document.all(Obj).style.pixelTop;}
function MMove(){if(Obj!=''){document.all(Obj).style.left=event.x-pX;document.all(Obj).style.top=event.y-pY;}}
function MUp(){if(Obj!=''){document.all(Obj).releaseCapture();Obj='';	}}
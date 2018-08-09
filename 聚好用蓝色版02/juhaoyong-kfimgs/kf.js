$(document).ready(function(){
  $("#juhaoyong_xuanfukefuBut").mouseenter(function(){
    $("#juhaoyong_xuanfukefuContent").css("display","block");
	$("#juhaoyong_xuanfukefuBut").css("display","none");
  });
  $("#juhaoyong_xuanfukefuContent").mouseleave(function(){
    $("#juhaoyong_xuanfukefuContent").css("display","none");
	$("#juhaoyong_xuanfukefuBut").css("display","block");
  });
});

juhaoyongKefu=function(id,_top,_left){
	var me=id.charAt?document.getElementById(id):id,d1=document.body,d2=document.documentElement;
	me.style.top=_top?_top+'px':0;
	me.style.left=_left+"px";
	me.style.position='absolute';
	setInterval(function(){me.style.top=parseInt(me.style.top)+(Math.max(d1.scrollTop,d2.scrollTop)+_top-parseInt(me.style.top))*0.1+'px'},10+parseInt(Math.random()*20));
	return arguments.callee;
}
juhaoyongKefu('juhaoyong_xuanfukefu',175,0);
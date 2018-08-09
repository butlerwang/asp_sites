<%
response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1

Dim PhotoUrl:PhotoUrl=Request("photourl")
Dim del:del=request("del")
if PhotoUrl="" Then
  response.write "<script>alert('您没有选择图片!');window.close();</script>"
end if
if left(photourl,1)<>"/" and left(lcase(photourl),4)<>"http" then photourl="/" & PhotoUrl
if request("action")="main" then
 call main
else
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>在线图片裁剪</title>
<META HTTP-EQUIV="pragma" CONTENT="no-cache">
<META HTTP-EQUIV="Cache-Control" CONTENT="no-cache, must-revalidate">
<META HTTP-EQUIV="expires" CONTENT="Wed, 26 Feb 1997 08:21:57 GMT">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<body>
<iframe src="imgcut.asp?action=main&del=<%=del%>&photourl=<%=photourl%>" scrolling="yes" width="100%" height="540"></iframe>
</body>
</html>
<%
end if
sub main()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>在线图片裁剪</title>
</head>
<body bgcolor="#000000">
<script type="text/javascript">
var isIE = (document.all) ? true : false;

var isIE6 = isIE && ([/MSIE (\d)\.0/i.exec(navigator.userAgent)][0][1] == 6);

var $$ = function (id) {
	return "string" == typeof id ? document.getElementById(id) : id;
};

var Class = {
	create: function() {
		return function() { this.initialize.apply(this, arguments); }
	}
}

var Extend = function(destination, source) {
	for (var property in source) {
		destination[property] = source[property];
	}
}

var Bind = function(object, fun) {
	return function() {
		return fun.apply(object, arguments);
	}
}

var BindAsEventListener = function(object, fun) {
	var args = Array.prototype.slice.call(arguments).slice(2);
	return function(event) {
		return fun.apply(object, [event || window.event].concat(args));
	}
}

var CurrentStyle = function(element){
	return element.currentStyle || document.defaultView.getComputedStyle(element, null);
}

function addEventHandler(oTarget, sEventType, fnHandler) {
	if (oTarget.addEventListener) {
		oTarget.addEventListener(sEventType, fnHandler, false);
	} else if (oTarget.attachEvent) {
		oTarget.attachEvent("on" + sEventType, fnHandler);
	} else {
		oTarget["on" + sEventType] = fnHandler;
	}
};

function removeEventHandler(oTarget, sEventType, fnHandler) {
    if (oTarget.removeEventListener) {
        oTarget.removeEventListener(sEventType, fnHandler, false);
    } else if (oTarget.detachEvent) {
        oTarget.detachEvent("on" + sEventType, fnHandler);
    } else { 
        oTarget["on" + sEventType] = null;
    }
};
</script>
<script type="text/javascript" src="../ks_inc/imgplus/ImgCropper.js"></script>
<script type="text/javascript" src="../ks_inc/imgplus/Drag.js"></script>
<script type="text/javascript" src="../ks_inc/imgplus/Resize.js"></script>
<script src="../ks_inc/jquery.js"></script>
<style type="text/css">
#rRightDown,#rLeftDown,#rLeftUp,#rRightUp,#rRight,#rLeft,#rUp,#rDown{
	position:absolute;
	background:#FFF;
	border: 1px solid #333;
	width: 6px;
	height: 6px;
	z-index:500;
	font-size:0;
	opacity: 0.5;
	filter:alpha(opacity=50);
}

#rLeftDown,#rRightUp{cursor:ne-resize;}
#rRightDown,#rLeftUp{cursor:nw-resize;}
#rRight,#rLeft{cursor:e-resize;}
#rUp,#rDown{cursor:n-resize;}

#rLeftDown{left:0px;bottom:0px;}
#rRightUp{right:0px;top:0px;}
#rRightDown{right:0px;bottom:0px;background-color:#00F;}
#rLeftUp{left:0px;top:0px;}
#rRight{right:0px;top:50%;margin-top:-4px;}
#rLeft{left:0px;top:50%;margin-top:-4px;}
#rUp{top:0px;left:50%;margin-left:-4px;}
#rDown{bottom:0px;left:50%;margin-left:-4px;}

#bgDiv{ min-height:400px;border:3px solid #000; position:relative;}
#dragDiv{border:1px dashed #fff; width:150px; height:120px; top:50px; left:50px; cursor:move; }
</style>
<table border="0" width="99%" bgcolor="#666666" align="center" cellspacing="0" cellpadding="0">
  <tr>
    <td style="padding:10px">
	 <div id="bgDiv">
        <div id="dragDiv">
          <div id="rRightDown"> </div>
          <div id="rLeftDown"> </div>
          <div id="rRightUp"> </div>
          <div id="rLeftUp"> </div>
          <div id="rRight"> </div>
          <div id="rLeft"> </div>
          <div id="rUp"> </div>
          <div id="rDown"></div>
        </div>
      </div>
	  
	  <div id="tools"> 
  <input value="缩小原图" type="button" id="idSize_small" /> 
  <input value="放大原图" type="button" id="idSize_big" /> 
  <input value="默认大小" type="button" id="idSize_old" /> 
  裁剪宽度：<input value="200" name="drag_w" id="drag_w" type="text" style="width:30px;"/> px 
  裁剪高度：<input value="200" name="drag_h" id="drag_h" type="text" style="width:30px;"/> px 
   </div>

	  </td>
    <td valign="top" align="left">
	 <br/><br/>
	 <table border="0">
	  <tr>
	   <td>
	<div style="text-align:left;font-weight:bold;maring:2px">效果预览:</div>
	   </td>
	  </tr>
	  <tr>
	   <td style="height:120px">
	<div id="viewDiv" style="border:3px solid #000;width:200px; min-height:120px;"> </div>
	   </td>
	  </tr>
	  <tr>
	   <td style="height:40px;color:#ff6600;font-size:12px">
	    <form name="myform" id="myform" action="" method="post">
		 <%if del="1" then%>
		  <label><input type="checkbox" name="del" value="1">删除原图</label>
	  <br/> <br/>
		  如果图片在其它地方有用到,请不要勾选删除原图。
		  <br/>
		 <%end if%>
		  <br/>
	       <input name="" type="button" value="生成图片" onclick="Create()" />
    <input name="" type="button" value="放弃使用原图" onclick="top.close()"/>
        </form>
	   </td>
	  </tr>
	  </table>
	</td>
  </tr>
</table>
<br />
<br />

<Img id="si" src="<%=PhotoUrl%>" style="display:none"/>
<img id="imgCreat" style="display:none;" />

<script>
var h,w,ic;
var o_w,o_h,max_w=620,max_h=600; 
$(document).ready(function(){
 w=$("#si").width();
 h=$("#si").height();
 o_w=w; 
 o_h=h; 
if (w>max_w) {w=max_w;o_w=max_w;} 
if (h>max_h) {h=max_h;o_h=max_h;} 
 //if (h>600) h=600;
	  ic = new ImgCropper("bgDiv", "dragDiv", "<%=PhotoUrl%>", {
		Width:w, Height: h, Color: "#999999",
		Resize: true,
		Right: "rRight", Left: "rLeft", Up:	"rUp", Down: "rDown",
		RightDown: "rRightDown", LeftDown: "rLeftDown", RightUp: "rRightUp", LeftUp: "rLeftUp",
		Preview: "viewDiv", viewWidth: 200, viewHeight: 200
	})
});
$$("drag_w").onchange = function(){ 
v_drag_w=$$("drag_w").value; 
$$("dragDiv").style.width=v_drag_w+"px"; 
v_drag_h=$$("drag_h").value; 
$$("dragDiv").style.height=v_drag_h+"px"; 
ic.Resize=false; 
ic.Init(); 
} 
$$("drag_h").onchange = function(){ 
v_drag_w=$$("drag_w").value; 
$$("dragDiv").style.width=v_drag_w+"px"; 
v_drag_h=$$("drag_h").value; 
$$("dragDiv").style.height=v_drag_h+"px"; 
ic.Resize=false; 
ic.Init(); 
} 
//缩小原图尺寸 
$$("idSize_small").onclick = function(){ 
w=w*0.9; 
h=h*0.9; 
if (w<10) w=10; 
if (h<10) h=10; 
ic.Width = w; 
ic.Height = h; 
ic.Init(); 
} 
//放大原图尺寸 
$$("idSize_big").onclick = function(){ 
w=w*1.1; 
h=h*1.1; 
if (w>max_w) w=max_w; 
if (h>max_h) h=max_h; 
ic.Width = w; 
ic.Height = h; 
ic.Init(); 
} 
//还原原图尺寸 
$$("idSize_old").onclick = function(){ 
w=o_w; 
h=o_h; 
ic.Width = w; 
ic.Height = h; 
ic.Init(); 
} 
function Create(){
	var p = ic.Url, o = ic.GetPos();
	x = o.Left,
	y = o.Top,
	w = o.Width,
	h = o.Height,
	pw = ic._layBase.width,
	ph = ic._layBase.height;
	$("#myform").attr("action","ImgCutSave.asp?p=" + p + "&x=" + x + "&y=" + y + "&w=" + w + "&h=" + h + "&pw=" + pw + "&ph=" + ph + "&" + Math.random());
	$("#myform").submit();
}

$(function(){
 //$("#bgDiv").width($("#bgDiv").find("img").width());
 //$("#bgDiv").height($("#bgDiv").find("img").height());
 $("#bgDiv").width=max_w; 
$("#bgDiv").height=max_h; 

});
</script>

</body>
</html>
<%end sub%>

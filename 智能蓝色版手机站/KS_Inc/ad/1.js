var collection;
var closeB=false;
function floaters() {
this.items	= [];
this.addItem	= function(id,x,y,content)
		  {
			document.write('<DIV id='+id+' style="Z-INDEX: 10; POSITION: absolute;  width:80px; height:60px;left:'+(typeof(x)=='string'?eval(x):x)+';top:'+(typeof(y)=='string'?eval(y):y)+'">'+content+'</DIV>');
			var newItem				= {};
			newItem.object			= document.getElementById(id);
			newItem.x				= x;
			newItem.y				= y;
			this.items[this.items.length]		= newItem;
		  }
this.play	= function()
		  {
			collection				= this.items
			setInterval('play()',30);
		  }
}
function play()
{
	if(screen.width<=800 || closeB)
	{
		for(var i=0;i<collection.length;i++)
		{
			collection[i].object.style.display	= 'none';
		}
		return;
	}
	for(var i=0;i<collection.length;i++)
	{
		var followObj		= collection[i].object;
		var followObj_x		= (typeof(collection[i].x)=='string'?eval(collection[i].x):collection[i].x);
		var followObj_y		= (typeof(collection[i].y)=='string'?eval(collection[i].y):collection[i].y);
		if(followObj.offsetLeft!=(document.body.scrollLeft+followObj_x)) {
			var dx=(document.body.scrollLeft+followObj_x-followObj.offsetLeft)*delta;
			dx=(dx>0?1:-1)*Math.ceil(Math.abs(dx));
			followObj.style.left=followObj.offsetLeft+dx;
			}
		if(followObj.offsetTop!=(document.body.scrollTop+followObj_y)) {
			var dy=(document.body.scrollTop+followObj_y-followObj.offsetTop)*delta;
			dy=(dy>0?1:-1)*Math.ceil(Math.abs(dy));
			followObj.style.top=followObj.offsetTop+dy;
			}
		followObj.style.display	= '';
	}
}
function closeBanner()
{
	closeB=true;
	return;
}
var theFloaters		= new floaters();
var sr='<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="100" height="250"><param name="movie" value="'+rightSrc+'"><param name="quality" value="high"><embed src="'+rightSrc+'" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="100" height="250"></embed></object><br>';
if (closeSrc!='0' && closeSrc!='')
sr+='<br><img src="'+closeSrc+'" onClick="closeBanner();">';

var sl='<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="100" height="250"><param name="movie" value="'+leftSrc+'"><param name="quality" value="high"><embed src="'+leftSrc+'" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="100" height="250"></embed></object><br>';
if (closeSrc!='0' && closeSrc!='')
sl+='<br><img src="'+closeSrc+'" onClick="closeBanner();">';

theFloaters.addItem('followDiv1','document.body.clientWidth-106',80,sr);
theFloaters.addItem('followDiv2',6,80,sl);
theFloaters.play();
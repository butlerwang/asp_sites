var FlashSrc; //定义动漫地址
var total;//定义flash影片总帧数
var frame_number;//定义flash影片当前帧数

//以下是滚动条图片拖动程序
var dragapproved=false;
var z,x,y
//移动函数
function move(){
if (event.button==1&&dragapproved){
y=temp1+event.clientX-x;
//以下是控制移动的范围
if(y<0)
 y=0;
if(y>500)
 y=500;

z.style.pixelLeft=y
movie.GotoFrame(y/500*total);//移动到某一位置，flash影片播放到某个位置
return false
}
}
//获得拖动前初始数据的函数
function drags(){
if (!document.all)
return
if (event.srcElement.className=="drag"){
dragapproved=true
z=event.srcElement
temp1=z.style.pixelLeft
x=event.clientX
document.onmousemove=move
}
}

//动态显示播放影片的当前帧/总帧数
function ShowCount(){
 frame_number=movie.CurrentFrame();
 frame_number++;
 frameCount.innerText=frame_number+"/"+movie.TotalFrames;
 element.style.pixelLeft=480*(frame_number/movie.TotalFrames)-15;//滚动条图片随之到相应的位置
 if(frame_number==movie.TotalFrames)
  clearTimeout(tn_ID);
 else
  var tn_ID=setTimeout('ShowCount();',1000);
}
//使影片返回第一帧 
function Rewind(){
 if(movie.IsPlaying()){
 Pause();
 }
 movie.Rewind();
 element.style.pixelLeft=0;
 frameCount.innerText="1/"+total;
}
//播放影片 
function Play(){
 movie.Play();
 ShowCount();
}
//暂停播放
function Pause(){
 movie.StopPlay();
}

//跳至最末帧
function GoToEnd(){
 if(movie.IsPlaying())
  Pause();
 movie.GotoFrame(total);
 element.style.pixelLeft=500;
 frameCount.innerText=total+"/"+total;
}
//快退影片
function Back()
{
 if(movie.IsPlaying())
  Pause();
 frame_number=frame_number-50;
 movie.GotoFrame(frame_number);
 Play();
}
//快进影片
function Forward()
{
 if(movie.IsPlaying())
  Pause();
 frame_number=frame_number+50;
 movie.GotoFrame(frame_number);
 Play();
}
//重新播放影片
function Replay(){
 if(movie.IsPlaying()){
 Pause();
 movie.Rewind();
 Play();
 }
 else
 {
 movie.Rewind();
 Play(); 
 }
}
//停止播放影片返回到第一帧
function Stop(){
 if(movie.IsPlaying()){
 Pause();
 movie.Rewind();
 }
 else
 {
 movie.Rewind();
 }
}
//全屏观看
function FullScreen()
{
 window.open(FlashSrc);	
}
//显示影片载入进度，完全载入后控制按钮可用
function Loading(){
	
 var in_ID;
 bar.style.width=Math.round(movie.PercentLoaded())+"%";
 frameCount.innerText=Math.round(movie.PercentLoaded())+"%";
 if(movie.PercentLoaded() >= 100){
  PlayerButtons.document.all.tags('IMG')[0].disabled=false;
  PlayerButtons.document.all.tags('IMG')[1].disabled=false;
  PlayerButtons.document.all.tags('IMG')[2].disabled=false;
  PlayerButtons.document.all.tags('IMG')[3].disabled=false;
  PlayerButtons.document.all.tags('IMG')[4].disabled=false;
  PlayerButtons.document.all.tags('IMG')[5].disabled=false;
  PlayerButtons.document.all.tags('IMG')[6].disabled=false;
  PlayerButtons.document.all.tags('IMG')[7].disabled=false;
  PlayerButtons.document.all.tags('IMG')[8].disabled=false;

total=movie.TotalFrames;
  frame_number++;
  frameCount.innerText=frame_number+"/"+total;
  bar.style.background="";
  bar.innerHTML='<img src="/Images/Default/posbar1.gif" style="POSITION:relative;cursor:pointer;border:0;" id="element" class="drag" OnMouseOver="fnOnMouseOver()" OnMouseOut="fnOnMouseOut()">';
  document.onmousedown=drags
  document.onmouseup=new Function("dragapproved=false;Play()")
  ShowCount();
  clearTimeout(in_ID);
 }
 else
  in_ID=setTimeout("Loading();",1000);
}

//开始载入flash影片，载入过程中，播放控制按钮不可用
function LoadFlashUrl(FlashUrl,FlashWidth,FlashHeight){
 FlashSrc=FlashUrl;
 movie.LoadMovie(0, FlashUrl);
 movie.width=FlashWidth;
 movie.height=FlashHeight;
 PlayerButtons.document.all.tags('IMG')[0].disabled=true;
 PlayerButtons.document.all.tags('IMG')[1].disabled=true;
 PlayerButtons.document.all.tags('IMG')[2].disabled=true;
 PlayerButtons.document.all.tags('IMG')[3].disabled=true;
 PlayerButtons.document.all.tags('IMG')[4].disabled=true;
 PlayerButtons.document.all.tags('IMG')[5].disabled=true;
 PlayerButtons.document.all.tags('IMG')[6].disabled=true;
 PlayerButtons.document.all.tags('IMG')[7].disabled=true;
 PlayerButtons.document.all.tags('IMG')[8].disabled=true;

 frame_number=movie.CurrentFrame();
 Loading();
}
//显示层函数
function showMenu(menu){
menu.style.display='block';
}

//鼠标点击滚动条上的位置，影片相应播放到那个位置
function Jump(fnume){
 if(movie.IsPlaying()){
 Pause();
 movie.GotoFrame(fnume);
 Play();
 }
 else
 {
 movie.GotoFrame(fnume);
 Play();
 }
}

//以下两个函数是图片切换函数
function fnOnMouseOver(){
 element.src = "/Images/Default/posbar.gif";
}

function fnOnMouseOut(){
 element.src = "/Images/Default/posbar1.gif";
}


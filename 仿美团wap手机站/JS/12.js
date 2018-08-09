var GetRandomn = 1;
function GetRandom(n){GetRandomn=Math.floor(Math.random()*n+1)}
 var a12=new Array();
var t12=new Array();
var ts12=new Array();
a12[0]="<span onclick=\"addHits12(0,10)\"><a href=\"#\" target=\"_blank\"><img  alt=\"kesioncms右边广告位\"  border=\"0\"  height=250  width=275  src=\"/images/flashad.gif\"></a></span>";
t12[0]=0;
ts12[0]="2012-6-15";
var temp12=new Array();
var k=0;
for(var i=0;i<a12.length;i++){
if (t12[i]==1){
if (checkDate12(ts12[i])){
	temp12[k++]=a12[i];
}
	}else{
 temp12[k++]=a12[i];
}
}
if (temp12.length>0){
GetRandom(temp12.length);
document.write(a12[GetRandomn-1]);
}
function addHits12(c,id){if(c==1){try{jQuery.getScript('http://192.168.1.103/plus/ajaxs.asp?action=HitsGuangGao&id='+id,function(){});}catch(e){}}}
function checkDate12(date_arr){
 var date=new Date();
 date_arr=date_arr.split("-");
var year=parseInt(date_arr[0]);
var month=parseInt(date_arr[1])-1;
var day=0;
if (date_arr[2].indexOf(" ")!=-1)
day=parseInt(date_arr[2].split(" ")[0]);
else
day=parseInt(date_arr[2]);
var date1=new Date(year,month,day);
if(date.valueOf()>date1.valueOf())
 return false;
else
 return true
}

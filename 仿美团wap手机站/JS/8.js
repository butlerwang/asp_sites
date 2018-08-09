var GetRandomn = 1;
function GetRandom(n){GetRandomn=Math.floor(Math.random()*n+1)}
 var a8=new Array();
var t8=new Array();
var ts8=new Array();
a8[0]="<span onclick=\"addHits8(0,5)\"><a href=\"/shop/pack.asp?id=8\" target=\"_blank\"><img  alt=\"礼包广告\"  border=\"0\"  src=\"/images/lb950x90.jpg\"></a></span>";
t8[0]=0;
ts8[0]="2012-5-10";
var temp8=new Array();
var k=0;
for(var i=0;i<a8.length;i++){
if (t8[i]==1){
if (checkDate8(ts8[i])){
	temp8[k++]=a8[i];
}
	}else{
 temp8[k++]=a8[i];
}
}
if (temp8.length>0){
GetRandom(temp8.length);
document.write(a8[GetRandomn-1]);
}
function addHits8(c,id){if(c==1){try{jQuery.getScript('http://192.168.1.101/plus/ajaxs.asp?action=HitsGuangGao&id='+id,function(){});}catch(e){}}}
function checkDate8(date_arr){
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

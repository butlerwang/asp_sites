var GetRandomn = 1;
function GetRandom(n){GetRandomn=Math.floor(Math.random()*n+1)}
 var a16=new Array();
var t16=new Array();
var ts16=new Array();
a16[0]="<span onclick=\"addHits16(0,15)\"><a href=\"/shop/pack.asp?id=8\" target=\"_blank\"><img  alt=\"礼包\"  border=\"0\"  height=91  width=965  src=\"/images/lb950x90.jpg\"></a></span>";
t16[0]=0;
ts16[0]="2012-6-20";
var temp16=new Array();
var k=0;
for(var i=0;i<a16.length;i++){
if (t16[i]==1){
if (checkDate16(ts16[i])){
	temp16[k++]=a16[i];
}
	}else{
 temp16[k++]=a16[i];
}
}
if (temp16.length>0){
GetRandom(temp16.length);
document.write(a16[GetRandomn-1]);
}
function addHits16(c,id){if(c==1){try{jQuery.getScript('http://192.168.1.104/plus/ajaxs.asp?action=HitsGuangGao&id='+id,function(){});}catch(e){}}}
function checkDate16(date_arr){
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

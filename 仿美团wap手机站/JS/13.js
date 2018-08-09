var GetRandomn = 1;
function GetRandom(n){GetRandomn=Math.floor(Math.random()*n+1)}
 var a13=new Array();
var t13=new Array();
var ts13=new Array();
a13[0]="<span onclick=\"addHits13(0,11)\"><a href=\"#/down\" target=\"_blank\"><img  alt=\"V9下载试用\"  border=\"0\"  height=95  width=680  src=\"/images/ad680.gif\"></a></span>";
t13[0]=0;
ts13[0]="2012-6-15";
var temp13=new Array();
var k=0;
for(var i=0;i<a13.length;i++){
if (t13[i]==1){
if (checkDate13(ts13[i])){
	temp13[k++]=a13[i];
}
	}else{
 temp13[k++]=a13[i];
}
}
if (temp13.length>0){
GetRandom(temp13.length);
document.write(a13[GetRandomn-1]);
}
function addHits13(c,id){if(c==1){try{jQuery.getScript('http://192.168.1.103/plus/ajaxs.asp?action=HitsGuangGao&id='+id,function(){});}catch(e){}}}
function checkDate13(date_arr){
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

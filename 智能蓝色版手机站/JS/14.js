var GetRandomn = 1;
function GetRandom(n){GetRandomn=Math.floor(Math.random()*n+1)}
 var a14=new Array();
var t14=new Array();
var ts14=new Array();
a14[0]="<span onclick=\"addHits14(0,12)\"><a href=\"#/down\" target=\"_blank\"><img  alt=\"kesioncms\"  border=\"0\"  height=74  width=965  src=\"/images/banner1.png\"></a></span>";
t14[0]=0;
ts14[0]="2012-6-18";
var temp14=new Array();
var k=0;
for(var i=0;i<a14.length;i++){
if (t14[i]==1){
if (checkDate14(ts14[i])){
	temp14[k++]=a14[i];
}
	}else{
 temp14[k++]=a14[i];
}
}
if (temp14.length>0){
GetRandom(temp14.length);
document.write(a14[GetRandomn-1]);
}
function addHits14(c,id){if(c==1){try{jQuery.getScript('http://192.168.1.104/plus/ajaxs.asp?action=HitsGuangGao&id='+id,function(){});}catch(e){}}}
function checkDate14(date_arr){
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

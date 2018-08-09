var GetRandomn = 1;
function GetRandom(n){GetRandomn=Math.floor(Math.random()*n+1)}
 var a6=new Array();
var t6=new Array();
var ts6=new Array();
a6[0]="<span onclick=\"addHits6(1,4)\"><a href=\"#\" target=\"_blank\"><img  alt=\"横幅965*54\"  border=\"0\"  height=54  width=965  src=\"/images/banner.gif\"></a></span>";
t6[0]=0;
ts6[0]="2012-4-17";
var temp6=new Array();
var k=0;
for(var i=0;i<a6.length;i++){
if (t6[i]==1){
if (checkDate6(ts6[i])){
	temp6[k++]=a6[i];
}
	}else{
 temp6[k++]=a6[i];
}
}
if (temp6.length>0){
GetRandom(temp6.length);
document.write(a6[GetRandomn-1]);
}
function addHits6(c,id){if(c==1){try{jQuery.getScript('http://192.168.1.103:95/plus/ajaxs.asp?action=HitsGuangGao&id='+id,function(){});}catch(e){}}}
function checkDate6(date_arr){
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

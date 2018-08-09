var GetRandomn = 1;
function GetRandom(n){GetRandomn=Math.floor(Math.random()*n+1)}
 var a10=new Array();
var t10=new Array();
var ts10=new Array();
a10[0]="<span onclick=\"addHits10(0,8)\"><a href=\"/mnkc\" target=\"_blank\"><img  alt=\"考试广告右\"  border=\"0\"  height=80  width=480  src=\"/mnkc/images/mnkcbr.png\"></a></span>";
t10[0]=0;
ts10[0]="2012-5-14";
var temp10=new Array();
var k=0;
for(var i=0;i<a10.length;i++){
if (t10[i]==1){
if (checkDate10(ts10[i])){
	temp10[k++]=a10[i];
}
	}else{
 temp10[k++]=a10[i];
}
}
if (temp10.length>0){
GetRandom(temp10.length);
document.write(a10[GetRandomn-1]);
}
function addHits10(c,id){if(c==1){try{jQuery.getScript('http://192.168.1.103/plus/ajaxs.asp?action=HitsGuangGao&id='+id,function(){});}catch(e){}}}
function checkDate10(date_arr){
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

var GetRandomn = 1;
function GetRandom(n){GetRandomn=Math.floor(Math.random()*n+1)}
 var a11=new Array();
var t11=new Array();
var ts11=new Array();
a11[0]="<span onclick=\"addHits11(0,9)\"><a href=\"/user\" target=\"_blank\"><img  alt=\"互动\"  border=\"0\"  height=69  width=275  src=\"/images/fc.png\"></a></span>";
t11[0]=0;
ts11[0]="2012-6-15";
var temp11=new Array();
var k=0;
for(var i=0;i<a11.length;i++){
if (t11[i]==1){
if (checkDate11(ts11[i])){
	temp11[k++]=a11[i];
}
	}else{
 temp11[k++]=a11[i];
}
}
if (temp11.length>0){
GetRandom(temp11.length);
document.write(a11[GetRandomn-1]);
}
function addHits11(c,id){if(c==1){try{jQuery.getScript('http://192.168.1.103/plus/ajaxs.asp?action=HitsGuangGao&id='+id,function(){});}catch(e){}}}
function checkDate11(date_arr){
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

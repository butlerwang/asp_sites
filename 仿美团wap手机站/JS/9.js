var GetRandomn = 1;
function GetRandom(n){GetRandomn=Math.floor(Math.random()*n+1)}
 var a9=new Array();
var t9=new Array();
var ts9=new Array();
a9[0]="<span onclick=\"addHits9(0,13)\"><a href=\"#\" target=\"_blank\"><img  alt=\"考试广告位700\"  border=\"0\"  height=80  width=700  src=\"/mnkc/images/Dvrf.gif\"></a></span>";
t9[0]=0;
ts9[0]="2012-6-20";
a9[1]="<span onclick=\"addHits9(0,7)\"><a href=\"/mnkc\" target=\"_blank\"><img  alt=\"考试中心\"  border=\"0\"  height=80  width=480  src=\"/mnkc/images/mnkcbl.png\"></a></span>";
t9[1]=0;
ts9[1]="2012-5-14";
var temp9=new Array();
var k=0;
for(var i=0;i<a9.length;i++){
if (t9[i]==1){
if (checkDate9(ts9[i])){
	temp9[k++]=a9[i];
}
	}else{
 temp9[k++]=a9[i];
}
}
if (temp9.length>0){
GetRandom(temp9.length);
document.write(a9[GetRandomn-1]);
}
function addHits9(c,id){if(c==1){try{jQuery.getScript('http://192.168.1.104/plus/ajaxs.asp?action=HitsGuangGao&id='+id,function(){});}catch(e){}}}
function checkDate9(date_arr){
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

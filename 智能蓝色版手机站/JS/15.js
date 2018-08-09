var GetRandomn = 1;
function GetRandom(n){GetRandomn=Math.floor(Math.random()*n+1)}
 var a15=new Array();
var t15=new Array();
var ts15=new Array();
a15[0]="<span onclick=\"addHits15(0,14)\"><a href=\"#\" target=\"_blank\"><img  alt=\"考试系统广告700\"  border=\"0\"  height=80  width=700  src=\"/mnkc/images/Dvrf.gif\"></a></span>";
t15[0]=0;
ts15[0]="2012-6-20";
var temp15=new Array();
var k=0;
for(var i=0;i<a15.length;i++){
if (t15[i]==1){
if (checkDate15(ts15[i])){
	temp15[k++]=a15[i];
}
	}else{
 temp15[k++]=a15[i];
}
}
if (temp15.length>0){
GetRandom(temp15.length);
document.write(a15[GetRandomn-1]);
}
function addHits15(c,id){if(c==1){try{jQuery.getScript('http://192.168.1.104/plus/ajaxs.asp?action=HitsGuangGao&id='+id,function(){});}catch(e){}}}
function checkDate15(date_arr){
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

tmpDate = new Date();
date = tmpDate.getDate();
month= tmpDate.getMonth() + 1 ;
year= tmpDate.getFullYear();
document.write(year);
document.write("年");
document.write(month);
document.write("月");
document.write(date);
document.write("日 ");
myArray=new Array(6);
myArray[0]="星期日"
myArray[1]="星期一"
myArray[2]="星期二"
myArray[3]="星期三"
myArray[4]="星期四"
myArray[5]="星期五"
myArray[6]="星期六"
weekday=tmpDate.getDay();
if (weekday==0 | weekday==6)
{
document.write(myArray[weekday])
}
else
{document.write(myArray[weekday])
};
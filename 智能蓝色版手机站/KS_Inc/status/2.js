function KS_Status2(seed)
{ 
var m2 = "" ;
var msg=m1+m2;
var out = " ";
var c = 1;
if (seed > 100)
{ seed-=2;
var cmd="KS_Status2(" + seed + ")";
timerTwo=window.setTimeout(cmd,speed);}
else if (seed <= 100 && seed > 0)
{ for (c=0 ; c < seed ; c++)
{ out+=" ";}
out+=msg; seed-=2;
var cmd="KS_Status2(" + seed + ")";
window.status=out;
timerTwo=window.setTimeout(cmd,speed); }
else if (seed <= 0)
{ if (-seed < msg.length)
{
out+=msg.substring(-seed,msg.length);
seed-=2;
var cmd="KS_Status2(" + seed + ")";
window.status=out;
timerTwo=window.setTimeout(cmd,speed);}
else { window.status=" ";
timerTwo=window.setTimeout("KS_Status2(100)",speed);
}
}
}
KS_Status2(100);
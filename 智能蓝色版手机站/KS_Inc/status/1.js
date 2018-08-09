var spacelen = 120;
var space10=" ";
var seq=0;
function KS_Status1() {
len = msg.length;
window.status = msg.substring(0, seq+1);
seq++;
if ( seq >= len ) {
seq = 0;
window.status = '';
window.setTimeout("KS_Status1();", interval );
}
 else
window.setTimeout("KS_Status1();", interval );
}
KS_Status1();
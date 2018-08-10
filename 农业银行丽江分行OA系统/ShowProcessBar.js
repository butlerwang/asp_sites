AddProcessbar();
var bwidth=0;
var swidth = document.all.waiting.clientWidth;

function CheckIsProcessBar(obj)
{
if (obj.IsShowProcessBar=="True") 
{
return false;
} 
else
{
return true;
}
}

function CheckClick(e)
{
if (e == 1)
{
if (bwidth<swidth*0.98){
bwidth += (swidth - bwidth) * 0.025;
if (document.all)document.sbar.width = bwidth;
else document.rating.clip.width = bwidth;
setTimeout('CheckClick(1);',1000);

}
}
else
{
if(document.all)
{
if(document.all.waiting.style.visibility == 'visible')
{document.all.waiting.style.visibility = 'hidden';
bwidth = 1;}
whichIt = event.srcElement;

while (CheckIsProcessBar(whichIt))
{
whichIt = whichIt.parentElement;
if (whichIt == null)return true;
}


document.all.waiting.style.pixelTop = (document.body.offsetHeight - document.all.waiting.clientHeight) / 2 + document.body.scrollTop;
document.all.waiting.style.pixelLeft = (document.body.offsetWidth - document.all.waiting.clientWidth) / 2 + document.body.scrollLeft;
document.all.waiting.style.visibility = 'visible';
if(!bwidth)CheckClick(1);
bwidth = 1;

}

else
{

if(document.waiting.visibility == 'show')
{document.waiting.visibility = 'hide';
document.rating.visibility = 'hide';
bwidth = 1;}
if(e.target.href.toString() != '')
{
document.waiting.top = (window.innerHeight - document.waiting.clip.height) / 2 + self.pageYOffset;
document.waiting.left = (window.innerWidth - document.waiting.clip.width) / 2 + self.pageXOffset;
document.waiting.visibility = 'show';
document.rating.top = (window.innerHeight - document.waiting.clip.height) / 2 + self.pageYOffset+document.waiting.clip.height-10;
document.rating.left = (window.innerWidth - document.waiting.clip.width) / 2 + self.pageXOffset;
document.rating.visibility = 'show';
if(!bwidth)CheckClick(1);
bwidth = 1;
}
}
return true;
}
}

function AddProcessbar()
{

var Str=""
Str+= "<div id=waiting style=position:absolute;top:50px;left:100px;z-index:1;visibility:hidden >";
Str+= "<layer name=waiting visibility=visible zIndex=2 >"
Str+= "<table border=2 cellspacing=1 cellpadding=0 bordercolorlight=#FFFFFF bordercolordark=#C0C0C0 bgcolor=#E0E0E0>"
Str+= " <tr>"
Str+= " <td bgcolor=#E0E0E0 height=30px width=300px align=center>"
Str+= " <font color=black>数据正在处理中...</font>"
Str+= " </td>"
Str+= " </tr>"
Str+= " <tr>"
Str+= " <td bgcolor=#E0E0E0>"
Str+= " <img width=1 height=10 name=sbar style=background-color:#6699cc>"
Str+= " </td>"
Str+= " </tr>"
Str+= "</table> "
Str+= "</layer>"
Str+= "</div>"
document.write(Str)

if(document.all)document.onclick = CheckClick; 
}
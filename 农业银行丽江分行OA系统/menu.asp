<!--#include file="check.asp"-->
<html><head>
<meta content="text/html; charset=gb2312" http-equiv=Content-Type>
<script language=JavaScript>
function OpenWindows(url)
{
  var 
newwin=window.open(url,"_blank","toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,top=50,left=120,width=250,height=120");
 return false;
}
function DynLayer(id,nestref,frame) {
    if (!DynLayer.set && !frame) DynLayerInit()
    this.frame = frame || self
    if (is.ns) {
        if (is.ns4) {
            if (!frame) {
                if (!nestref) var nestref = DynLayer.nestRefArray[id]
            if (!DynLayerTest(id,nestref)) return
                this.css = (nestref)? eval("document."+nestref+".document."+id) : document.layers[id]
            }
            else this.css = (nestref)? eval("frame.document."+nestref+".document."+id) : frame.document.layers[id]
            this.elm = this.event = this.css
            this.doc = this.css.document
        }
        if (is.ns5) {
            this.elm = document.getElementById(id)
            this.css = this.elm.style
            this.doc = document
        }
        this.x = this.css.left
        this.y = this.css.top
        this.w = this.css.clip.width
        this.h = this.css.clip.height
    }
    else if (is.ie) {
        this.elm = this.event = this.frame.document.all[id]
        this.css = this.frame.document.all[id].style
        this.doc = document
        this.x = this.elm.offsetLeft
        this.y = this.elm.offsetTop
        this.w = (is.ie4)? this.css.pixelWidth : this.elm.offsetWidth
        this.h = (is.ie4)? this.css.pixelHeight : this.elm.offsetHeight
    }
    this.id = id
    this.nestref = nestref
    this.obj = id + "DynLayer"
    eval(this.obj + "=this")
}
function DynLayerSetWidth(w) {
    this.css.width = w>0?w:0+"px"
}
function DynLayerSetHeight(h) {
    this.css.height = h>0?h:0+"px"
}
function DynLayerMoveTo(x,y) {
    if (x!=null) {
        this.x = x
        if (is.ns) this.css.left = this.x
        else this.css.pixelLeft = this.x
    }
    if (y!=null) {
        this.y = y
        if (is.ns) this.css.top = this.y
        else this.css.pixelTop = this.y
    }
}
function DynLayerMoveX(x) {
    if (x!=null) {
        this.x = x
        this.css.left = this.x
    }
}
function DynLayerMoveY(y) {
    if (y!=null) {
        this.y = y
    this.css.top = this.y
    }
}
function DynLayerMoveBy(x,y) {
    this.moveTo(this.x+x,this.y+y)
}
function DynLayerShow() {
    this.css.visibility = (is.ns)? "show" : "visible"
}
function DynLayerHide() {
    this.css.visibility = (is.ns)? "hide" : "hidden"
}
DynLayer.prototype.moveTo = DynLayerMoveTo
DynLayer.prototype.moveX = DynLayerMoveX
DynLayer.prototype.moveY = DynLayerMoveY
DynLayer.prototype.moveBy = DynLayerMoveBy
DynLayer.prototype.show = DynLayerShow
DynLayer.prototype.hide = DynLayerHide
DynLayer.prototype.setWidth = DynLayerSetWidth
DynLayer.prototype.setHeight = DynLayerSetHeight
DynLayerTest = new Function('return true')
// DynLayerInit Function
function DynLayerInit(nestref) {
    if (!DynLayer.set) DynLayer.set = true
    if (is.ns) {
        if (nestref) ref = eval('document.'+nestref+'.document')
        else {nestref = ''; ref = document;}
        for (var i=0; i<ref.layers.length; i++) {
            var divname = ref.layers[i].name
            DynLayer.nestRefArray[divname] = nestref
            var index = divname.indexOf("Div")
            if (index > 0) {
                eval(divname.substr(0,index)+' = new DynLayer("'+divname+'","'+nestref+'")')
            }

            if (ref.layers[i].document.layers.length > 0) {

                DynLayer.refArray[DynLayer.refArray.length] = (nestref=='')? ref.layers[i].name : nestref+'.document.'+ref.layers[i].name
            }
        }
        if (DynLayer.refArray.i < DynLayer.refArray.length) {
            DynLayerInit(DynLayer.refArray[DynLayer.refArray.i++])
        }
    }
else if (is.ie) {
        for (var i=0; i<document.all.tags("DIV").length; i++) {
            var divname = document.all.tags("DIV")[i].id
            var index = divname.indexOf("Div")
            if (index > 0) {
                eval(divname.substr(0,index)+' = new DynLayer("'+divname+'")')
            }
        }
    }
    return true
}
DynLayer.nestRefArray = new Array()
DynLayer.refArray = new Array()
DynLayer.refArray.i = 0
DynLayer.set = false
// Slide Methods
function DynLayerSlideTo(endx,endy,inc,speed,fn) {
    if (endx==null) endx = this.x
    if (endy==null) endy = this.y
    var distx = endx-this.x
    var disty = endy-this.y
    this.slideStart(endx,endy,distx,disty,inc,speed,fn)
}
function DynLayerSlideBy(distx,disty,inc,speed,fn) {
    var endx = this.x + distx
    var endy = this.y + disty
    this.slideStart(endx,endy,distx,disty,inc,speed,fn)
}
function DynLayerSlideStart(endx,endy,distx,disty,inc,speed,fn) {
    if (this.slideActive) return
    if (!inc) inc = 10
    if (!speed) speed = 5
    var num = Math.sqrt(Math.pow(distx,2) + Math.pow(disty,2))/inc
    if (num==0) return
    var dx = distx/num
    var dy = disty/num
    if (!fn) fn = null
    this.slideActive = true
    this.slide(dx,dy,endx,endy,num,1,speed,fn)
}
function DynLayerSlide(dx,dy,endx,endy,num,i,speed,fn) {
    if (!this.slideActive) return
    if (i++ < num) {
        this.moveBy(dx,dy)
        this.onSlide()
        if (this.slideActive) setTimeout(this.obj+".slide("+dx+","+dy+","+endx+","+endy+","+num+","+i+","+speed+",\""+fn+"\")",speed)
        else this.onSlideEnd()
    }
    else {
        this.slideActive = false
        this.moveTo(endx,endy)
        this.onSlide()
        this.onSlideEnd()
        eval(fn)
    }
}
DynLayerSlideInit = new Function()
DynLayer.prototype.slideInit = new Function()
DynLayer.prototype.slideTo = DynLayerSlideTo
DynLayer.prototype.slideBy = DynLayerSlideBy
DynLayer.prototype.slideStart = DynLayerSlideStart
DynLayer.prototype.slide = DynLayerSlide
DynLayer.prototype.onSlide = new Function()
DynLayer.prototype.onSlideEnd = new Function()
// Clip Methods
function DynLayerClipInit(clipTop,clipRight,clipBottom,clipLeft) {
    if (is.ie) {
        if (arguments.length==4) this.clipTo(clipTop,clipRight,clipBottom,clipLeft)
        else if (is.ie4) this.clipTo(0,this.css.pixelWidth,this.css.pixelHeight,0)
    }
}
function DynLayerClipTo(t,r,b,l) {
    if (t==null) t = this.clipValues('t')
    if (r==null) r = this.clipValues('r')
    if (b==null) b = this.clipValues('b')
    if (l==null) l = this.clipValues('l')
    if (is.ns) {
        this.css.clip.top = t
        this.css.clip.right = r
        this.css.clip.bottom = b
        this.css.clip.left = l
    }
    else if (is.ie) this.css.clip = "rect("+t+"px "+r+"px "+b+"px "+l+"px)"
}
function DynLayerClipBy(t,r,b,l) {
    this.clipTo(this.clipValues('t')+t,this.clipValues('r')+r,this.clipValues('b')+b,this.clipValues('l')+l)
}
function DynLayerClipValues(which) {
    if (is.ie) var clipv = this.css.clip.split("rect(")[1].split(")")[0].split("px")
    if (which=="t") return (is.ns)? this.css.clip.top : Number(clipv[0])
    if (which=="r") return (is.ns)? this.css.clip.right : Number(clipv[1])
    if (which=="b") return (is.ns)? this.css.clip.bottom : Number(clipv[2])
    if (which=="l") return (is.ns)? this.css.clip.left : Number(clipv[3])
}
DynLayer.prototype.clipInit = DynLayerClipInit
DynLayer.prototype.clipTo = DynLayerClipTo
DynLayer.prototype.clipBy = DynLayerClipBy
DynLayer.prototype.clipValues = DynLayerClipValues
// Write Method

function DynLayerWrite(html) {

    if (is.ns) {

        this.doc.open()

        this.doc.write(html)

        this.doc.close()

    }

    else if (is.ie) {

        this.event.innerHTML = html

    }

}

DynLayer.prototype.write = DynLayerWrite



// BrowserCheck Object

function BrowserCheck() {

    var b = navigator.appName

    if (b=="Netscape") this.b = "ns"

    else if (b=="Microsoft Internet Explorer") this.b = "ie"

    else this.b = b

    this.version = navigator.appVersion

    this.v = parseInt(this.version)

    this.ns = (this.b=="ns" && this.v>=4)

    this.ns4 = (this.b=="ns" && this.v==4)

    this.ns5 = (this.b=="ns" && this.v==5)

    this.ie = (this.b=="ie" && this.v>=4)

    this.ie4 = (this.version.indexOf('MSIE 4')>0)

    this.ie5 = (this.version.indexOf('MSIE 5')>0)

    this.min = (this.ns||this.ie)

}

is = new BrowserCheck()



// CSS Function

function css(id,left,top,width,height,color,vis,z,other) {

    if (id=="START") return '<STYLE TYPE="text/css">\n'

    else if (id=="END") return '</STYLE>'

    var str = (left!=null && top!=null)? '#'+id+' {position:absolute; left:'+left+'px; top:'+top+'px;' : '#'+id+' {position:relative;'

    if (arguments.length>=4 && width!=null) str += ' width:'+width+'px;'

    if (arguments.length>=5 && height!=null) {

        str += ' height:'+height+'px;'

        if (arguments.length<9 || other.indexOf('clip')==-1) str += ' clip:rect(0px '+width+'px '+height+'px 0px);'

    }

    if (arguments.length>=6 && color!=null) str += (is.ns)? ' layer-background-color:'+color+';' : ' background-color:'+color+';'

    if (arguments.length>=7 && vis!=null) str += ' visibility:'+vis+';'

    if (arguments.length>=8 && z!=null) str += ' z-index:'+z+';'

    if (arguments.length==9 && other!=null) str += ' '+other

    str += '}\n'

    return str

}

function writeCSS(str,showAlert) {

    str = css('START')+str+css('END')

    document.write(str)

    if (showAlert) alert(str)

}

</script>

<script language=JavaScript1.2>
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
  
}


function pop1(pageurl) {
  var popwin=window.open(pageurl,"popWin","scrollbars=yes,toolbar=no,location=no,directories=no,status=no,menubar=no,resizable=no,width=540,height=370");
  return false;
  }

//-->
</script>

<style type=text/css>#menulayer0Div {
    BACKGROUND-COLOR: #ffffff; CLIP: rect(0px 4000px 3000px 0px); HEIGHT: 3000px; LEFT: 0px; POSITION: absolute; TOP: 0px; layer-background-color: #ffffff
}
#iconlayer0Div {
    BACKGROUND-COLOR: #ffffff; CLIP: rect(0px 4000px 3000px 0px); HEIGHT: 3000px; LEFT: 0px; POSITION: absolute; TOP: 22px; layer-background-color: #ffffff
}
#barlayer0Div {
    BACKGROUND-COLOR: #ffffff; CLIP: rect(0px 4000px 22px 0px); HEIGHT: 22px; LEFT: 0px; POSITION: absolute; TOP: 0px; layer-background-color: #ffffff
}
#uplayer0Div {
    BACKGROUND-COLOR: #cccccc; CLIP: rect(0px 16px 16px 0px); HEIGHT: 16px; LEFT: -20px; POSITION: absolute; TOP: 26px; WIDTH: 16px; layer-background-color: #cccccc
}
#downlayer0Div {
    BACKGROUND-COLOR: #cccccc; CLIP: rect(0px 16px 16px 0px); HEIGHT: 16px; LEFT: -20px; POSITION: absolute; TOP: 42px; WIDTH: 16px; layer-background-color: #cccccc
}
#menulayer1Div {
    BACKGROUND-COLOR: #ffffff; CLIP: rect(0px 4000px 3000px 0px); HEIGHT: 3000px; LEFT: 0px; POSITION: absolute; TOP: 0px; layer-background-color: #ffffff
}
#iconlayer1Div {
    BACKGROUND-COLOR: #ffffff; CLIP: rect(0px 4000px 3000px 0px); HEIGHT: 3000px; LEFT: 0px; POSITION: absolute; TOP: 22px; layer-background-color: #ffffff
}
#barlayer1Div {
    BACKGROUND-COLOR: #ffffff; CLIP: rect(0px 4000px 22px 0px); HEIGHT: 22px; LEFT: 0px; POSITION: absolute; TOP: 0px; layer-background-color: #ffffff
}
#uplayer1Div {
    BACKGROUND-COLOR: #cccccc; CLIP: rect(0px 16px 16px 0px); HEIGHT: 16px; LEFT: -20px; POSITION: absolute; TOP: 26px; WIDTH: 16px; layer-background-color: #cccccc
}
#downlayer1Div {
    BACKGROUND-COLOR: #cccccc; CLIP: rect(0px 16px 16px 0px); HEIGHT: 16px; LEFT: -20px; POSITION: absolute; TOP: 42px; WIDTH: 16px; layer-background-color: #cccccc
}
#menulayer2Div {
    BACKGROUND-COLOR: #ffffff; CLIP: rect(0px 4000px 3000px 0px); HEIGHT: 3000px; LEFT: 0px; POSITION: absolute; TOP: 0px; layer-background-color: #ffffff
}
#iconlayer2Div {
    BACKGROUND-COLOR: #ffffff; CLIP: rect(0px 4000px 3000px 0px); HEIGHT: 3000px; LEFT: 0px; POSITION: absolute; TOP: 22px; layer-background-color: #ffffff
}
#barlayer2Div {
    BACKGROUND-COLOR: #ffffff; CLIP: rect(0px 4000px 22px 0px); HEIGHT: 22px; LEFT: 0px; POSITION: absolute; TOP: 0px; layer-background-color: #ffffff
}
#uplayer2Div {
    BACKGROUND-COLOR: #cccccc; CLIP: rect(0px 16px 16px 0px); HEIGHT: 16px; LEFT: -20px; POSITION: absolute; TOP: 26px; WIDTH: 16px; layer-background-color: #cccccc
}
#downlayer2Div {
    BACKGROUND-COLOR: #cccccc; CLIP: rect(0px 16px 16px 0px); HEIGHT: 16px; LEFT: -20px; POSITION: absolute; TOP: 42px; WIDTH: 16px; layer-background-color: #cccccc
}
#menulayer3Div {
    BACKGROUND-COLOR: #ffffff; CLIP: rect(0px 4000px 3000px 0px); HEIGHT: 3000px; LEFT: 0px; POSITION: absolute; TOP: 0px; layer-background-color: #ffffff
}
#iconlayer3Div {
    BACKGROUND-COLOR: #ffffff; CLIP: rect(0px 4000px 3000px 0px); HEIGHT: 3000px; LEFT: 0px; POSITION: absolute; TOP: 22px; layer-background-color: #ffffff
}
#barlayer3Div {
    BACKGROUND-COLOR: #ffffff; CLIP: rect(0px 4000px 22px 0px); HEIGHT: 22px; LEFT: 0px; POSITION: absolute; TOP: 0px; layer-background-color: #ffffff
}
#uplayer3Div {
    BACKGROUND-COLOR: #cccccc; CLIP: rect(0px 16px 16px 0px); HEIGHT: 16px; LEFT: -20px; POSITION: absolute; TOP: 26px; WIDTH: 16px; layer-background-color: #cccccc
}
#downlayer3Div {
    BACKGROUND-COLOR: #cccccc; CLIP: rect(0px 16px 16px 0px); HEIGHT: 16px; LEFT: -20px; POSITION: absolute; TOP: 42px; WIDTH: 16px; layer-background-color: #cccccc
}
#menulayer4Div {
    BACKGROUND-COLOR: #ffffff; CLIP: rect(0px 4000px 3000px 0px); HEIGHT: 3000px; LEFT: 0px; POSITION: absolute; TOP: 0px; layer-background-color: #ffffff
}
#iconlayer4Div {
    BACKGROUND-COLOR: #ffffff; CLIP: rect(0px 4000px 3000px 0px); HEIGHT: 3000px; LEFT: 0px; POSITION: absolute; TOP: 22px; layer-background-color: #ffffff
}
#barlayer4Div {
    BACKGROUND-COLOR: #ffffff; CLIP: rect(0px 4000px 22px 0px); HEIGHT: 22px; LEFT: 0px; POSITION: absolute; TOP: 0px; layer-background-color: #ffffff
}
#uplayer4Div {
    BACKGROUND-COLOR: #cccccc; CLIP: rect(0px 16px 16px 0px); HEIGHT: 16px; LEFT: -20px; POSITION: absolute; TOP: 26px; WIDTH: 16px; layer-background-color: #cccccc
}
#downlayer4Div {
    BACKGROUND-COLOR: #cccccc; CLIP: rect(0px 16px 16px 0px); HEIGHT: 16px; LEFT: -20px; POSITION: absolute; TOP: 42px; WIDTH: 16px; layer-background-color: #cccccc
}
#menulayer5Div {
    BACKGROUND-COLOR: #ffffff; CLIP: rect(0px 4000px 3000px 0px); HEIGHT: 3000px; LEFT: 0px; POSITION: absolute; TOP: 0px; layer-background-color: #ffffff
}
#iconlayer5Div {
    BACKGROUND-COLOR: #ffffff; CLIP: rect(0px 4000px 3000px 0px); HEIGHT: 3000px; LEFT: 0px; POSITION: absolute; TOP: 22px; layer-background-color: #ffffff
}
#barlayer5Div {
    BACKGROUND-COLOR: #ffffff; CLIP: rect(0px 4000px 22px 0px); HEIGHT: 22px; LEFT: 0px; POSITION: absolute; TOP: 0px; layer-background-color: #ffffff
}
#uplayer5Div {
    BACKGROUND-COLOR: #cccccc; CLIP: rect(0px 16px 16px 0px); HEIGHT: 16px; LEFT: -20px; POSITION: absolute; TOP: 26px; WIDTH: 16px; layer-background-color: #cccccc
}
#downlayer5Div {
    BACKGROUND-COLOR: #cccccc; CLIP: rect(0px 16px 16px 0px); HEIGHT: 16px; LEFT: -20px; POSITION: absolute; TOP: 42px; WIDTH: 16px; layer-background-color: #cccccc
}
#menulayer6Div {
    BACKGROUND-COLOR: #ffffff; CLIP: rect(0px 4000px 3000px 0px); HEIGHT: 3000px; LEFT: 0px; POSITION: absolute; TOP: 0px; layer-background-color: #ffffff
}
#iconlayer6Div {
    BACKGROUND-COLOR: #ffffff; CLIP: rect(0px 4000px 3000px 0px); HEIGHT: 3000px; LEFT: 0px; POSITION: absolute; TOP: 22px; layer-background-color: #ffffff
}
#barlayer6Div {
    BACKGROUND-COLOR: #ffffff; CLIP: rect(0px 4000px 22px 0px); HEIGHT: 22px; LEFT: 0px; POSITION: absolute; TOP: 0px; layer-background-color: #ffffff
}
#uplayer6Div {
    BACKGROUND-COLOR: #cccccc; CLIP: rect(0px 16px 16px 0px); HEIGHT: 16px; LEFT: -20px; POSITION: absolute; TOP: 26px; WIDTH: 16px; layer-background-color: #cccccc
}
#downlayer6Div {
    BACKGROUND-COLOR: #cccccc; CLIP: rect(0px 16px 16px 0px); HEIGHT: 16px; LEFT: -20px; POSITION: absolute; TOP: 42px; WIDTH: 16px; layer-background-color: #cccccc
}
</style>

<script language=JavaScript>

<!--



var menubarheight = 0;

var menubarsum = 0;

var menuspeed = 10;

var menuinc = 100;

var scrollspeed = 100;

var scrollinc = 60;

var menuchoose = 0;

var iconX = new Array(menubarsum);

var menuIconWidth = new Array(menubarsum);

var menuIconHeight = new Array(menubarsum);

var menuscroll = 0;

var iconareaheight = 0;

var iconrightpos = 0;

var maxscroll = 0;

var scrolling = false;

var scrollTimerID = 0;



function init(mnum, mheight) {

    menubarheight = mnum

    menubarsum = mheight

    menulayer = new Array(menubarsum)

    iconlayer = new Array(menubarsum)

    barlayer = new Array(menubarsum)

    uplayer = new Array(menubarsum)

    downlayer = new Array(menubarsum)

    for (var i=0; i<menubarsum; i++) {

        menulayer[i] = new DynLayer("menulayer" + i + "Div")

        menulayer[i].slideInit()



        iconlayer[i] = new DynLayer("iconlayer" + i + "Div", "menulayer" + i + "Div")

        iconlayer[i].slideInit()

        /*iconlayer[i].setWidth(document.body.clientWidth);*/



        if (menuIconWidth[i] > document.body.clientWidth) {

            iconlayer[i].setWidth(menuIconWidth[i])

            iconX[i] = (document.body.clientWidth-menuIconWidth[i])/2

        } else {

            iconlayer[i].setWidth(document.body.clientWidth)

            iconX[i] = 0

        }

        iconlayer[i].moveTo(iconX[i], menubarheight)



        barlayer[i] = new DynLayer("barlayer" + i + "Div", "menulayer" + i + "Div")

        barlayer[i].slideInit()



        uplayer[i] = new DynLayer("uplayer" + i + "Div", "menulayer" + i + "Div")

        uplayer[i].slideInit()



        downlayer[i] = new DynLayer("downlayer" + i + "Div", "menulayer" + i + "Div")

        downlayer[i].slideInit()

        

    }

    menureload()

    

    

}



function menubarpush(num) {

    if (num != menuchoose && num >= 0 && num < menubarsum) {

    

        iconlayer[menuchoose].moveTo(iconX[menuchoose],menubarheight)

        menuscroll = 0

        scrolling = false

    

        for (var i=0; i <=num; i++) {

            menulayer[i].slideTo(0, i*menubarheight, menuinc, menuspeed)

        }

        nAdCornerOriginY = document.body.clientHeight;

        nAdCornerOriginY += document.body.scrollTop;

        for (var i=menubarsum-1; i>num; i--) {

            nAdCornerOriginY -= menubarheight

            menulayer[i].slideTo(0,nAdCornerOriginY, menuinc, menuspeed)

        }

        menuchoose = num

        menuscrollbar()

    }

}





function menureload() {

    nAdCornerOriginY = document.body.clientHeight;

    nAdCornerOriginY += document.body.scrollTop;

    for (var i=menubarsum-1; i>menuchoose; i--) {

        nAdCornerOriginY -= menubarheight

        menulayer[i].moveTo(0, nAdCornerOriginY)

    }

    for (var i=0; i<menubarsum; i++) {

        if (menuIconWidth[i] > document.body.clientWidth) {

            iconlayer[i].setWidth(menuIconWidth[i])

            iconX[i] = (document.body.clientWidth-menuIconWidth[i])/2

        } else {

            iconlayer[i].setWidth(document.body.clientWidth)

            iconX[i] = 0

        }

        iconlayer[i].moveX(iconX[i], menubarheight)

    }

    

    

    menuscrollbar()

}





function menuscrollbar() {

    iconareaheight = document.body.clientHeight-menubarheight*(menubarsum);

    iconrightpos = document.body.clientWidth-16-4;

    maxscroll = menuIconHeight[menuchoose] - iconareaheight

    

    

    

    if (maxscroll > 0) {

        if (menuscroll > 0) {

            uplayer[menuchoose].moveTo(iconrightpos, menubarheight+4) 

        } else {

            uplayer[menuchoose].moveTo(-20, 0)

        }

        if (menuscroll < maxscroll) {

            downlayer[menuchoose].moveTo(iconrightpos, iconareaheight+2)

        } else {

            downlayer[menuchoose].moveTo(-20, 0)

        }

    } else {

        if (menuscroll <= 0) 

            uplayer[menuchoose].moveTo(-20, 0)

        downlayer[menuchoose].moveTo(-20, 0)

    }

}





function menuscrollup() {

    if (menuscroll > 0) {

        scrolling = true

        menuscroll -= scrollinc

        iconlayer[menuchoose].moveTo(iconX[menuchoose], menubarheight-menuscroll)

        

        scrollTimerID = setTimeout("menuscrollup()", scrollspeed)

    } else {

        menuscrollstop()    

    }

    menuscrollbar()

    

}




function menuscrolldown() {

    if (menuscroll < maxscroll) {

        scrolling = true

        menuscroll += scrollinc

        if (menuscroll < maxscroll) {

            iconlayer[menuchoose].moveTo(iconX[menuchoose], menubarheight-menuscroll)

        } else {

            iconlayer[menuchoose].moveTo(iconX[menuchoose], menubarheight-maxscroll)

        }

        

        scrollTimerID = setTimeout("menuscrolldown()", scrollspeed)

    } else {

        menuscrollstop()    

    }



    menuscrollbar()

    

}



function menuscrollstop() {

    scrolling = false

    if (scrollTimerID) {

        clearTimeout(scrollTimerID)

        scrollTimerID = 0;

    }

    

}



//-->

</script>

<script id=clientEventHandlersJS language=javascript>

<!--



function window_onresize() {

    menureload()

}



//-->

</script>

<style type=text/css>A:link {
    COLOR: #0000b7; TEXT-DECORATION: none
}
A:visited {
    COLOR: #0000b7; TEXT-DECORATION: none
}
A:active {
    COLOR: #0000b7; TEXT-DECORATION: none
}
A:hover {
    COLOR: #0000b7; TEXT-DECORATION: none
}
TD {
    COLOR: white; FONT-SIZE: 9pt
}
TH {
    COLOR: white; FONT-SIZE: 9pt
}
FONT {
    FONT-SIZE: 9pt
}
.chinese_text13 {
    FILTER: Blur(Add=0, Direction=0, Strength=0); FONT-FAMILY: Verdana,宋体; FONT-SIZE: 9pt
}
</style>

<meta content="MSHTML 5.00.2614.3500" name=GENERATOR></head>
<body bgColor=#ffffff onload=init(22,7) onresize="return window_onresize()">
<div id=menulayer0Div>
<div id=iconlayer0Div>
<table align=center border=0 cellPadding=1 cellSpacing=0 width="100%">
   <TBODY>



<tr>
       <tr>
    <td align=middle>
      <table bgColor=#ffffff border=1 borderColorDark=#ffffff 
      borderColorLight=#ffffff cellPadding=0 cellSpacing=0 
      onmousedown="this.borderColorLight='#000000';this.borderColorDark='#cccccc'" 
      onmouseout="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'" 
      onmouseover="this.borderColorLight='#cccccc';this.borderColorDark='#000000'" 
      onmouseup="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'">
        <tbody>
        <tr>
          <td bgColor=#ffffff borderColorDark=#ffffff 
            borderColorLight=#ffffff><a 
            href="../LW-TJ/index-TJ.ASP" 
            target=main><img align=middle alt=统计数据查询 border=0 height=32 
            src="images/spb.gif" style="FILTER: alpha(opacity=100)" 
            width=32></a> </td></tr></tbody></table></td></tr>
  <tr>
    <td align=middle class=chinese_text13><a 
      href="../LW-TJ/index-TJ.ASP" target=main>统计查询 
</a></td></tr>

<tr>
    <td align=middle>
      <table bgColor=#ffffff border=1 borderColorDark=#ffffff 
      borderColorLight=#ffffff cellPadding=0 cellSpacing=0 
      onmousedown="this.borderColorLight='#000000';this.borderColorDark='#cccccc'" 
      onmouseout="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'" 
      onmouseover="this.borderColorLight='#cccccc';this.borderColorDark='#000000'" 
      onmouseup="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'">
        <tbody>
        <tr>
          <td bgColor=#ffffff borderColorDark=#ffffff 
            borderColorLight=#ffffff><a 
            href="#" 
            target=main><img align=middle alt=（信贷电子化资料查询） border=0 height=32 
            src="images/advise.gif" style="FILTER: alpha(opacity=100)" 
            width=32></a> </td></tr></tbody></table></td></tr>
  <tr>
    <td align=middle class=chinese_text13><a 
      href="#" target=main>信贷查询 
    </a></td></tr>

<tr>
    <td align=middle>
      <table bgColor=#ffffff border=1 borderColorDark=#ffffff 
      borderColorLight=#ffffff cellPadding=0 cellSpacing=0 
      onmousedown="this.borderColorLight='#000000';this.borderColorDark='#cccccc'" 
      onmouseout="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'" 
      onmouseover="this.borderColorLight='#cccccc';this.borderColorDark='#000000'" 
      onmouseup="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'">
        <tbody>
        <tr>
          <td bgColor=#ffffff borderColorDark=#ffffff 
            borderColorLight=#ffffff><a 
            href="../lw-gw/index.asp" 
            target=main><img align=middle alt=（旧系统下的原公文查询） border=0 height=32 
            src="images/advise.gif" style="FILTER: alpha(opacity=100)" 
            width=32></a> </td></tr></tbody></table></td></tr>
  <tr>
    <td align=middle class=chinese_text13><a 
      href="../lw-gw/index.asp" target=main>原公文查询 
    </a></td></tr>

<TR>
    <TD align=middle>
      <TABLE bgColor=#ffffff border=1 borderColorDark=#ffffff 
      borderColorLight=#ffffff cellPadding=0 cellSpacing=0 
      onmousedown="this.borderColorLight='#000000';this.borderColorDark='#cccccc'" 
      onmouseout="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'" 
      onmouseover="this.borderColorLight='#cccccc';this.borderColorDark='#000000'" 
      onmouseup="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'">
        <TBODY>
        <TR>
          <TD bgColor=#ffffff borderColorDark=#ffffff 
            borderColorLight=#ffffff><A 
            href="tongzhi.asp" 
            target=main><IMG align=middle alt=内部通知 border=0 height=32 
            src="images/bulletin.gif" style="FILTER: alpha(opacity=100)" 
            width=32></A> </TD></TR></TBODY></TABLE></TD></TR>
  <TR>
    <TD align=middle class=chinese_text13><A 
      href="tongzhi.asp" target=main>内部通知
      </A></TD></TR>
  <TR>
    <TD height=4></TD></TR>
     
  <TR>
    <TD align=middle>
      <table bgColor=#ffffff border=1 borderColorDark=#ffffff 
      borderColorLight=#ffffff cellPadding=0 cellSpacing=0 
      onmousedown="this.borderColorLight='#000000';this.borderColorDark='#cccccc'" 
      onmouseout="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'" 
      onmouseover="this.borderColorLight='#cccccc';this.borderColorDark='#000000'" 
      onmouseup="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'">
        <tbody>
        <tr>
          <td bgColor=#ffffff borderColorDark=#ffffff 
            borderColorLight=#ffffff><a 
            href="learn.asp" 
            target=main><img align=middle alt=文件学习 border=0 height=32 
            src="images/school.gif" style="FILTER: alpha(opacity=100)" 
            width=32></a> </td></tr></tbody></table></td></tr>
  <tr>
    <td align=middle class=chinese_text13><a 
      href="learn.asp" target=main>文件学习
      </a></td></tr>
  <tr>
    <td height=4></td></tr>
  <tr>
    <td align=middle>
      <table bgColor=#ffffff border=1 borderColorDark=#ffffff 
      borderColorLight=#ffffff cellPadding=0 cellSpacing=0 
      onmousedown="this.borderColorLight='#000000';this.borderColorDark='#cccccc'" 
      onmouseout="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'" 
      onmouseover="this.borderColorLight='#cccccc';this.borderColorDark='#000000'" 
      onmouseup="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'">
        <tbody>
        <tr>
          <td bgColor=#ffffff borderColorDark=#ffffff 
            borderColorLight=#ffffff><a href="addfile.asp" 
            target=main><img align=middle alt=上报文件 border=0 height=32 
            src="images/criterion.gif" style="FILTER: alpha(opacity=100)" 
            width=32></a> </td></tr></tbody></table></td></tr>
  <tr>
    <td align=middle class=chinese_text13><a 
      href="addfile.asp" target=main>上报文件
</a></td></tr></tbody></table></div>
<div id=uplayer0Div><img height=16 
onmousedown="javascript:this.src='images/scrollup2.gif';menuscrollup()" 
onmouseout="javascript:this.src='images/scrollup.gif';menuscrollstop()" 
onmouseup="javascript:this.src='images/scrollup.gif';menuscrollstop()" 
src="images/scrollup.gif" title=更多 width=16> </div>
<div id=downlayer0Div><img height=16 
onmousedown="javascript:this.src='images/scrolldown2.gif';menuscrolldown()" 
onmouseout="javascript:this.src='images/scrolldown.gif';menuscrollstop()" 
onmouseup="javascript:this.src='images/scrolldown.gif';menuscrollstop()" 
src="images/scrolldown.gif" title=更多 width=16> </div>
<div id=barlayer0Div>
<table bgColor=#cccccc border=0 borderColorDark=#505050 borderColorLight=white 
cellPadding=0 cellSpacing=0 height=22 onclick=javascript:menubarpush(0) 
onmousedown="javascript:this.borderColorDark='White';this.borderColorLight='#505050'" 
onmouseout="javascript:this.borderColorDark='#505050';this.borderColorLight='White'" 
onmouseup="javascript:this.borderColorDark='#505050';this.borderColorLight='White'" 
style="CURSOR: hand" width="100%">
  <tbody>
  <tr>
    <td align=middle background=images/menu_bg.gif borderColorDark=#cccccc 
    borderColorLight=#cccccc colSpan=0 noWrap rowSpan=0><font 
      class=chinese_text13 color=#ffffff width="100%">行政管理 </font></td></tr></tbody></table></div>
<script id=clientEventHandlersJS language=javascript>

<!--

menuIconWidth[0] = iconlayer0Div.scrollWidth + 0;

menuIconHeight[0] = iconlayer0Div.scrollHeight + 0;

//-->

</script>
</div>
<div id=menulayer1Div>
<div id=iconlayer1Div>
<table align=center border=0 cellPadding=1 cellSpacing=0 width="100%">
  <tbody>
<tr>
    <td align=middle>
      <table bgColor=#ffffff border=1 borderColorDark=#ffffff 
      borderColorLight=#ffffff cellPadding=0 cellSpacing=0 
      onmousedown="this.borderColorLight='#000000';this.borderColorDark='#cccccc'" 
      onmouseout="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'" 
      onmouseover="this.borderColorLight='#cccccc';this.borderColorDark='#000000'" 
      onmouseup="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'">
        <tbody>
        <tr>
          <td bgColor=#ffffff borderColorDark=#ffffff 
            borderColorLight=#ffffff><a 
            href="tel.asp" 
            target=main><img align=middle alt=常用电话 border=0 height=32 
            src="images/phone.gif" style="FILTER: alpha(opacity=100)" 
            width=32></a> </td></tr></tbody></table></td></tr>
  <tr>
    <td align=middle class=chinese_text13><a 
      href="tel.asp" target=main>常用电话 
  </a></td></tr>
  <tr>
    <td height=4></td></tr>
    <tr>
    <td align=middle>
      <table bgColor=#ffffff border=1 borderColorDark=#ffffff 
      borderColorLight=#ffffff cellPadding=0 cellSpacing=0 
      onmousedown="this.borderColorLight='#000000';this.borderColorDark='#cccccc'" 
      onmouseout="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'" 
      onmouseover="this.borderColorLight='#cccccc';this.borderColorDark='#000000'" 
      onmouseup="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'">
        <tbody>
        <tr>
          <td bgColor=#ffffff borderColorDark=#ffffff 
            borderColorLight=#ffffff><a 
            href="url.asp" 
            target=main><img align=middle alt=常用网址 border=0 height=32 
            src="images/url.gif" style="FILTER: alpha(opacity=100)" 
            width=32></a> </td></tr></tbody></table></td></tr>
  <tr>
    <td align=middle class=chinese_text13><a 
      href="url.asp" target=main>常用网址 
</a></td></tr>
<tr>
       <tr>
    <td align=middle>
      <table bgColor=#ffffff border=1 borderColorDark=#ffffff 
      borderColorLight=#ffffff cellPadding=0 cellSpacing=0 
      onmousedown="this.borderColorLight='#000000';this.borderColorDark='#cccccc'" 
      onmouseout="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'" 
      onmouseover="this.borderColorLight='#cccccc';this.borderColorDark='#000000'" 
      onmouseup="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'">
        <tbody>
        <tr>
          <td bgColor=#ffffff borderColorDark=#ffffff 
            borderColorLight=#ffffff><a 
            href="index.asp" 
            target=main><img align=middle alt=邮政编码电话区号查询 border=0 height=32 
            src="images/spb.gif" style="FILTER: alpha(opacity=100)" 
            width=32></a> </td></tr></tbody></table></td></tr>
  <tr>
    <td align=middle class=chinese_text13><a 
      href="index.asp" target=main>邮编区号查询 
</a></td></tr>


<tr>
       <tr>
    <td align=middle>
      <table bgColor=#ffffff border=1 borderColorDark=#ffffff 
      borderColorLight=#ffffff cellPadding=0 cellSpacing=0 
      onmousedown="this.borderColorLight='#000000';this.borderColorDark='#cccccc'" 
      onmouseout="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'" 
      onmouseover="this.borderColorLight='#cccccc';this.borderColorDark='#000000'" 
      onmouseup="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'">
        <tbody>
        <tr>
          <td bgColor=#ffffff borderColorDark=#ffffff 
            borderColorLight=#ffffff><a 
            href="../LW-TJ/index-TJ.ASP" 
            target=main><img align=middle alt=统计数据查询 border=0 height=32 
            src="images/spb.gif" style="FILTER: alpha(opacity=100)" 
            width=32></a> </td></tr></tbody></table></td></tr>
  <tr>
    <td align=middle class=chinese_text13><a 
      href="../LW-TJ/index-TJ.ASP" target=main>统计查询 
</a></td></tr>

<tr>
    <td align=middle>
      <table bgColor=#ffffff border=1 borderColorDark=#ffffff 
      borderColorLight=#ffffff cellPadding=0 cellSpacing=0 
      onmousedown="this.borderColorLight='#000000';this.borderColorDark='#cccccc'" 
      onmouseout="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'" 
      onmouseover="this.borderColorLight='#cccccc';this.borderColorDark='#000000'" 
      onmouseup="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'">
        <tbody>
        <tr>
          <td bgColor=#ffffff borderColorDark=#ffffff 
            borderColorLight=#ffffff><a 
            href="#" 
            target=main><img align=middle alt=（信贷电子化资料查询） border=0 height=32 
            src="images/advise.gif" style="FILTER: alpha(opacity=100)" 
            width=32></a> </td></tr></tbody></table></td></tr>
  <tr>
    <td align=middle class=chinese_text13><a 
      href="#" target=main>信贷查询 
    </a></td></tr>

<tr>
    <td align=middle>
      <table bgColor=#ffffff border=1 borderColorDark=#ffffff 
      borderColorLight=#ffffff cellPadding=0 cellSpacing=0 
      onmousedown="this.borderColorLight='#000000';this.borderColorDark='#cccccc'" 
      onmouseout="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'" 
      onmouseover="this.borderColorLight='#cccccc';this.borderColorDark='#000000'" 
      onmouseup="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'">
        <tbody>
        <tr>
          <td bgColor=#ffffff borderColorDark=#ffffff 
            borderColorLight=#ffffff><a 
            href="../lw-gw/index.asp" 
            target=main><img align=middle alt=（旧系统下的原公文查询） border=0 height=32 
            src="images/advise.gif" style="FILTER: alpha(opacity=100)" 
            width=32></a> </td></tr></tbody></table></td></tr>
  <tr>
    <td align=middle class=chinese_text13><a 
      href="../lw-gw/index.asp" target=main>原公文查询 
    </a></td></tr>

<tr>
    <td height=4></td></tr>
  </tbody></table></div>
<div id=uplayer1Div><img height=16 
onmousedown="javascript:this.src='images/scrollup2.gif';menuscrollup()" 
onmouseout="javascript:this.src='images/scrollup.gif';menuscrollstop()" 
onmouseup="javascript:this.src='images/scrollup.gif';menuscrollstop()" 
src="images/scrollup.gif" title=更多 width=16> </div>
<div id=downlayer1Div><img height=16 
onmousedown="javascript:this.src='images/scrolldown2.gif';menuscrolldown()" 
onmouseout="javascript:this.src='images/scrolldown.gif';menuscrollstop()" 
onmouseup="javascript:this.src='images/scrolldown.gif';menuscrollstop()" 
src="images/scrolldown.gif" title=更多 width=16> </div>
<div id=barlayer1Div>
<table bgColor=#cccccc border=0 borderColorDark=#505050 borderColorLight=white 
cellPadding=0 cellSpacing=0 height=22 onclick=javascript:menubarpush(1) 
onmousedown="javascript:this.borderColorDark='White';this.borderColorLight='#505050'" 
onmouseout="javascript:this.borderColorDark='#505050';this.borderColorLight='White'" 
onmouseup="javascript:this.borderColorDark='#505050';this.borderColorLight='White'" 
style="CURSOR: hand" width="100%">
  <tbody>
  <tr>
    <td align=middle background=images/menu_bg.gif borderColorDark=#cccccc 
    borderColorLight=#cccccc colSpan=0 noWrap rowSpan=0><font 
      class=chinese_text13 color=white>公共信息 </font></td></tr></tbody></table></div>
<script id=clientEventHandlersJS language=javascript>

<!--

menuIconWidth[1] = iconlayer1Div.scrollWidth + 0;

menuIconHeight[1] = iconlayer1Div.scrollHeight + 0;

//-->

</script>
</div>
<div id=menulayer2Div>
<div id=iconlayer2Div>
<table align=center border=0 cellPadding=1 cellSpacing=0 width="100%">
  <tbody>
  <tr>
    <td align=middle>
      <table bgColor=#ffffff border=1 borderColorDark=#ffffff 
      borderColorLight=#ffffff cellPadding=0 cellSpacing=0 
      onmousedown="this.borderColorLight='#000000';this.borderColorDark='#cccccc'" 
      onmouseout="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'" 
      onmouseover="this.borderColorLight='#cccccc';this.borderColorDark='#000000'" 
      onmouseup="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'">
        <tbody>
        <tr>
          <td bgColor=#ffffff borderColorDark=#ffffff 
            borderColorLight=#ffffff><a 
            href="bbs.asp" 
            target=main><img align=middle alt=讨论中心 border=0 height=32 
            src="images/advise.gif" style="FILTER: alpha(opacity=100)" 
            width=32></a> </td></tr></tbody></table></td></tr>
  <tr>
    <td align=middle class=chinese_text13><a 
      href="bbs.asp" target=main>讨论中心 
    </a></td></tr>
  <tr>
    <td height=4></td></tr>
  <tr>
    <td align=middle>
      <table bgColor=#ffffff border=1 borderColorDark=#ffffff 
      borderColorLight=#ffffff cellPadding=0 cellSpacing=0 
      onmousedown="this.borderColorLight='#000000';this.borderColorDark='#cccccc'" 
      onmouseout="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'" 
      onmouseover="this.borderColorLight='#cccccc';this.borderColorDark='#000000'" 
      onmouseup="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'">
        <tbody>
        <tr>
          <td bgColor=#ffffff borderColorDark=#ffffff 
            borderColorLight=#ffffff><a 
            href="chat.asp" target=main><img 
            align=middle alt=会议中心 border=0 height=32 src="images/chat.gif" 
            style="FILTER: alpha(opacity=100)" width=32></a> 
    </td></tr></tbody></table></td></tr>
  <tr>
    <td align=middle class=chinese_text13><a 
      href="chat.asp" target=main>会议中心 
    </a></td></tr>
  <tr>
    <td height=4></td></tr>

    <tr>
    <td align=middle>
      <table bgColor=#ffffff border=1 borderColorDark=#ffffff 
      borderColorLight=#ffffff cellPadding=0 cellSpacing=0 
      onmousedown="this.borderColorLight='#000000';this.borderColorDark='#cccccc'" 
      onmouseout="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'" 
      onmouseover="this.borderColorLight='#cccccc';this.borderColorDark='#000000'" 
      onmouseup="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'">
        <tbody>
        <tr>
          <td bgColor=#ffffff borderColorDark=#ffffff 
            borderColorLight=#ffffff><a 
            href="file.asp" target=main><img 
            align=middle alt=软件下载 border=0 height=32 src="images/soft.gif" 
            style="FILTER: alpha(opacity=100)" width=32></a> 
    </td></tr></tbody></table></td></tr>
  <tr>
    <td align=middle class=chinese_text13><a 
      href="file.asp" target=main>软件下载 
    </a></td></tr>
  <tr>
    <td height=4></td></tr>
  
 
  
<tbody></table></div>
<div id=uplayer2Div><img height=16 
onmousedown="javascript:this.src='images/scrollup2.gif';menuscrollup()" 
onmouseout="javascript:this.src='images/scrollup.gif';menuscrollstop()" 
onmouseup="javascript:this.src='images/scrollup.gif';menuscrollstop()" 
src="images/scrollup.gif" title=更多 width=16> </div>
<div id=downlayer2Div><img height=16 
onmousedown="javascript:this.src='images/scrolldown2.gif';menuscrolldown()" 
onmouseout="javascript:this.src='images/scrolldown.gif';menuscrollstop()" 
onmouseup="javascript:this.src='images/scrolldown.gif';menuscrollstop()" 
src="images/scrolldown.gif" title=更多 width=16> </div>
<div id=barlayer2Div>
<table bgColor=#cccccc border=0 borderColorDark=#505050 borderColorLight=white 
cellPadding=0 cellSpacing=0 height=22 onclick=javascript:menubarpush(2) 
onmousedown="javascript:this.borderColorDark='White';this.borderColorLight='#505050'" 
onmouseout="javascript:this.borderColorDark='#505050';this.borderColorLight='White'" 
onmouseup="javascript:this.borderColorDark='#505050';this.borderColorLight='White'" 
style="CURSOR: hand" width="100%">
  <tbody>
  <tr>
    <td align=middle background=images/menu_bg.gif borderColorDark=#cccccc 
    borderColorLight=#cccccc colSpan=0 noWrap rowSpan=0><font 
      class=chinese_text13 color=white>交流中心 </font></td></tr></tbody></table></div>
<script id=clientEventHandlersJS language=javascript>

<!--

menuIconWidth[2] = iconlayer2Div.scrollWidth + 0;

menuIconHeight[2] = iconlayer2Div.scrollHeight + 0;

//-->

</script>
</div>
<div id=menulayer3Div>
<div id=iconlayer3Div>
<table align=center border=0 cellPadding=1 cellSpacing=0 width="100%">
  <tbody>
  <tr>
    <td align=middle>
      <table bgColor=#ffffff border=1 borderColorDark=#ffffff 
      borderColorLight=#ffffff cellPadding=0 cellSpacing=0 
      onmousedown="this.borderColorLight='#000000';this.borderColorDark='#cccccc'" 
      onmouseout="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'" 
      onmouseover="this.borderColorLight='#cccccc';this.borderColorDark='#000000'" 
      onmouseup="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'">
        <tbody>
        <tr>
          <td bgColor=#ffffff borderColorDark=#ffffff 
            borderColorLight=#ffffff><a 
            href="txl.asp" 
            target=main><img align=middle alt=个人通讯录 border=0 height=32 
            src="images/P_address.gif" style="FILTER: alpha(opacity=100)" 
            width=32></a> </td></tr></tbody></table></td></tr>
  <tr>
    <td align=middle class=chinese_text13><a 
      href="txl.asp" 
      target=main>个人通讯录 </a></td></tr>
 <tr>
    <td height=4></td></tr>
<tr>
    <td align=middle>
      <table bgColor=#ffffff border=1 borderColorDark=#ffffff 
      borderColorLight=#ffffff cellPadding=0 cellSpacing=0 
      onmousedown="this.borderColorLight='#000000';this.borderColorDark='#cccccc'" 
      onmouseout="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'" 
      onmouseover="this.borderColorLight='#cccccc';this.borderColorDark='#000000'" 
      onmouseup="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'">
        <tbody>
        <tr>
          <td bgColor=#ffffff borderColorDark=#ffffff 
            borderColorLight=#ffffff><a 
            href="calendar.asp" 
            target=main><img align=middle alt=日程安排 border=0 height=32 
            src="images/calendar.gif" style="FILTER: alpha(opacity=100)" 
            width=32></a> </td></tr></tbody></table></td></tr>
  <tr>
    <td align=middle class=chinese_text13><a 
      href="calendar.asp" 
      target=main>日程安排</a></td></tr>
 <tr>
  <td height=4></td></tr>
   <tr>
    <td align=middle>
      <table bgColor=#ffffff border=1 borderColorDark=#ffffff 
      borderColorLight=#ffffff cellPadding=0 cellSpacing=0 
      onmousedown="this.borderColorLight='#000000';this.borderColorDark='#cccccc'" 
      onmouseout="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'" 
      onmouseover="this.borderColorLight='#cccccc';this.borderColorDark='#000000'" 
      onmouseup="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'">
        <tbody>
        <tr>
          <td bgColor=#ffffff borderColorDark=#ffffff 
            borderColorLight=#ffffff><a 
            href="passwd.asp" 
            target=main><img align=middle alt=修改资料 border=0 height=32 
            src="images/M_password.gif" style="FILTER: alpha(opacity=100)" 
            width=32></a> </td></tr></tbody></table></td></tr>
  <tr>
    <td align=middle class=chinese_text13><a 
      href="passwd.asp" 
      target=main>修改资料</a></td></tr>
 <tr>
 <td height=4></td></tr>
  <tr>
    <td align=middle>
      <table bgColor=#ffffff border=1 borderColorDark=#ffffff 
      borderColorLight=#ffffff cellPadding=0 cellSpacing=0 
      onmousedown="this.borderColorLight='#000000';this.borderColorDark='#cccccc'" 
      onmouseout="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'" 
      onmouseover="this.borderColorLight='#cccccc';this.borderColorDark='#000000'" 
      onmouseup="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'">
        <tbody>
        <tr>
          <td bgColor=#ffffff borderColorDark=#ffffff 
            borderColorLight=#ffffff><a 
            href="http://www.yearcon.com/sms/index.php" target=main><img 
            align=middle alt=手机短消息 border=0 height=32 src="images/game.gif" 
            style="FILTER: alpha(opacity=100)" width=32></a> 
    </td></tr></tbody></table></td></tr>
  <tr>
    <td align=middle class=chinese_text13><a 
      href="http://www.yearcon.com/sms/index.php" target=main>手机短消息 
    </a></td></tr>
  <tr>
    <td height=4></td></tr>
  </tbody></table></div>
<div id=uplayer3Div><img height=16 
onmousedown="javascript:this.src='images/scrollup2.gif';menuscrollup()" 
onmouseout="javascript:this.src='images/scrollup.gif';menuscrollstop()" 
onmouseup="javascript:this.src='images/scrollup.gif';menuscrollstop()" 
src="images/scrollup.gif" title=更多 width=16> </div>
<div id=downlayer3Div><img height=16 
onmousedown="javascript:this.src='images/scrolldown2.gif';menuscrolldown()" 
onmouseout="javascript:this.src='images/scrolldown.gif';menuscrollstop()" 
onmouseup="javascript:this.src='images/scrolldown.gif';menuscrollstop()" 
src="images/scrolldown.gif" title=更多 width=16> </div>
<div id=barlayer3Div>
<table bgColor=#cccccc border=0 borderColorDark=#505050 borderColorLight=white 
cellPadding=0 cellSpacing=0 height=22 onclick=javascript:menubarpush(3) 
onmousedown="javascript:this.borderColorDark='White';this.borderColorLight='#505050'" 
onmouseout="javascript:this.borderColorDark='#505050';this.borderColorLight='White'" 
onmouseup="javascript:this.borderColorDark='#505050';this.borderColorLight='White'" 
style="CURSOR: hand" width="100%">
  <tbody>
  <tr>
    <td align=middle background=images/menu_bg.gif borderColorDark=#cccccc 
    borderColorLight=#cccccc colSpan=0 noWrap rowSpan=0><font 
      class=chinese_text13 color=white>个人助理</font></td></tr></tbody></table></div>
<script id=clientEventHandlersJS language=javascript>

<!--

menuIconWidth[3] = iconlayer3Div.scrollWidth + 0;

menuIconHeight[3] = iconlayer3Div.scrollHeight + 0;

//-->

</script>
</div>
<div id=menulayer4Div>
<div id=iconlayer4Div>
<table align=center border=0 cellPadding=1 cellSpacing=0 width="100%">
  <tbody>
  <tr>
    <td height=4></td></tr>
 <tr>
    <td align=middle>
      <table bgColor=#ffffff border=1 borderColorDark=#ffffff 
      borderColorLight=#ffffff cellPadding=0 cellSpacing=0 
      onmousedown="this.borderColorLight='#000000';this.borderColorDark='#cccccc'" 
      onmouseout="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'" 
      onmouseover="this.borderColorLight='#cccccc';this.borderColorDark='#000000'" 
      onmouseup="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'">
        <tbody>
        <tr>
          <td bgColor=#ffffff borderColorDark=#ffffff 
            borderColorLight=#ffffff><a 
            href="mailbox.asp?mailbox=common" 
            target=main><img align=middle alt=公共信件 border=0  
            src="images/mail2.gif" style="FILTER: alpha(opacity=100)" 
            ></a> </td></tr></tbody></table></td></tr>
  <tr>
    <td align=middle class=chinese_text13><a 
      href="mailbox.asp?mailbox=common" 
      target=main>公共信件</a></td></tr>
 <tr>
 <td height=4></td>
  <tr>
    <td align=middle>
      <table bgColor=#ffffff border=1 borderColorDark=#ffffff 
      borderColorLight=#ffffff cellPadding=0 cellSpacing=0 
      onmousedown="this.borderColorLight='#000000';this.borderColorDark='#cccccc'" 
      onmouseout="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'" 
      onmouseover="this.borderColorLight='#cccccc';this.borderColorDark='#000000'" 
      onmouseup="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'">
        <tbody>
        <tr>
          <td bgColor=#ffffff borderColorDark=#ffffff 
            borderColorLight=#ffffff><a 
            href="write.asp" 
            target=main><img align=middle alt=写邮件 border=0 
            src="images/mail6.gif" style="FILTER: alpha(opacity=100)" 
            ></a> </td></tr></tbody></table></td></tr>
  <tr>
    <td align=middle class=chinese_text13><a 
      href="write.asp" 
      target=main>写邮件</a></td></tr>
 <tr>
 <td height=4></td></tr>
  <tr>
    <td align=middle>
      <table bgColor=#ffffff border=1 borderColorDark=#ffffff 
      borderColorLight=#ffffff cellPadding=0 cellSpacing=0 
      onmousedown="this.borderColorLight='#000000';this.borderColorDark='#cccccc'" 
      onmouseout="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'" 
      onmouseover="this.borderColorLight='#cccccc';this.borderColorDark='#000000'" 
      onmouseup="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'">
        <tbody>
        <tr>
          <td bgColor=#ffffff borderColorDark=#ffffff 
            borderColorLight=#ffffff><a 
            href="mailbox.asp?mailbox=recived" 
            target=main><img align=middle alt=收件箱 border=0 
            src="images/mail1.gif" style="FILTER: alpha(opacity=100)" 
            ></a> </td></tr></tbody></table></td></tr>
  <tr>
    <td align=middle class=chinese_text13><a 
      href="mailbox.asp?mailbox=recived" 
      target=main>收件箱</a></td></tr>
 <tr>
 <td  height=4></td></tr>
  <tr>
    <td align=middle>
      <table bgColor=#ffffff border=1 borderColorDark=#ffffff 
      borderColorLight=#ffffff cellPadding=0 cellSpacing=0 
      onmousedown="this.borderColorLight='#000000';this.borderColorDark='#cccccc'" 
      onmouseout="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'" 
      onmouseover="this.borderColorLight='#cccccc';this.borderColorDark='#000000'" 
      onmouseup="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'">
        <tbody>
        <tr>
          <td bgColor=#ffffff borderColorDark=#ffffff 
            borderColorLight=#ffffff><a 
            href="mailbox.asp?mailbox=sendout" 
            target=main><img align=middle alt=发件箱 border=0 
            src="images/mail3.gif" style="FILTER: alpha(opacity=100)" 
            ></a> </td></tr></tbody></table></td></tr>
  <tr>
    <td align=middle class=chinese_text13><a 
      href="mailbox.asp?mailbox=sendout" 
      target=main>发件箱</a></td></tr>
 <tr>
 <td height=4></td></tr>
  <tr>
    <td align=middle>
      <table bgColor=#ffffff border=1 borderColorDark=#ffffff 
      borderColorLight=#ffffff cellPadding=0 cellSpacing=0 
      onmousedown="this.borderColorLight='#000000';this.borderColorDark='#cccccc'" 
      onmouseout="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'" 
      onmouseover="this.borderColorLight='#cccccc';this.borderColorDark='#000000'" 
      onmouseup="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'">
        <tbody>
        <tr>
          <td bgColor=#ffffff borderColorDark=#ffffff 
            borderColorLight=#ffffff><a 
            href="mailbox.asp?mailbox=del" 
            target=main><img align=middle alt=回收站 border=0 
            src="images/mail4.gif" style="FILTER: alpha(opacity=100)" 
            ></a> </td></tr></tbody></table></td></tr>
  <tr>
    <td align=middle class=chinese_text13><a 
      href="mailbox.asp?mailbox=del" 
      target=main>回收站</a></td></tr>
 <tr>
    <td height=4></td></tr>  
    
    </tbody></table></div>
<div id=uplayer4Div><img height=16 
onmousedown="javascript:this.src='images/scrollup2.gif';menuscrollup()" 
onmouseout="javascript:this.src='images/scrollup.gif';menuscrollstop()" 
onmouseup="javascript:this.src='images/scrollup.gif';menuscrollstop()" 
src="images/scrollup.gif" title=更多 width=16> </div>
<div id=downlayer4Div><img height=16 
onmousedown="javascript:this.src='images/scrolldown2.gif';menuscrolldown()" 
onmouseout="javascript:this.src='images/scrolldown.gif';menuscrollstop()" 
onmouseup="javascript:this.src='images/scrolldown.gif';menuscrollstop()" 
src="images/scrolldown.gif" title=更多 width=16> </div>
<div id=barlayer4Div>
<table bgColor=#cccccc border=0 borderColorDark=#505050 borderColorLight=white 
cellPadding=0 cellSpacing=0 height=22 onclick=javascript:menubarpush(4) 
onmousedown="javascript:this.borderColorDark='White';this.borderColorLight='#505050'" 
onmouseout="javascript:this.borderColorDark='#505050';this.borderColorLight='White'" 
onmouseup="javascript:this.borderColorDark='#505050';this.borderColorLight='White'" 
style="CURSOR: hand" width="100%">
  <tbody>
  <tr>
    <td align=middle background=images/menu_bg.gif borderColorDark=#cccccc 
    borderColorLight=#cccccc colSpan=0 noWrap rowSpan=0><font 
      class=chinese_text13 color=white>个人信箱</font></td></tr></tbody></table></div>
<script id=clientEventHandlersJS language=javascript>

<!--

menuIconWidth[4] = iconlayer4Div.scrollWidth + 0;

menuIconHeight[4] = iconlayer4Div.scrollHeight + 0;

//-->

</script>
</div>


<div id=menulayer5Div>
<div id=iconlayer5Div>
<table align=center border=0 cellPadding=1 cellSpacing=0 width="100%">
  <tbody> <tr>
      <td align=middle>
        <table bgColor=#ffffff border=1 borderColorDark=#ffffff 
      borderColorLight=#ffffff cellPadding=0 cellSpacing=0 
      onmousedown="this.borderColorLight='#000000';this.borderColorDark='#cccccc'" 
      onmouseout="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'" 
      onmouseover="this.borderColorLight='#cccccc';this.borderColorDark='#000000'" 
      onmouseup="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'">
        <tbody>
          <tr>
            <td bgColor=#ffffff borderColorDark=#ffffff 
            borderColorLight=#ffffff><a 
            href="mlearn.asp" target=main><img 
            align=middle alt=学习文件的管理 border=0 height=32 
            src="images/filemana.gif" style="FILTER: alpha(opacity=100)" 
            width=32></a> 
            </td>
           </tr>
        </tbody>
        </table>
       </td>
     </tr>
     <tr>
       <td align=middle class=chinese_text13><a 
      href="mlearn.asp" target=main>文件管理
       </a></td>
     </tr>
     <tr>
      <td height=4>
      </td>
     </tr>
     </tbody>
    </table>
    <table align=center border=0 cellPadding=1 cellSpacing=0 width="100%">
  <tbody>
    <tr>
      <td align=middle>
        <table bgColor=#ffffff border=1 borderColorDark=#ffffff 
      borderColorLight=#ffffff cellPadding=0 cellSpacing=0 
      onmousedown="this.borderColorLight='#000000';this.borderColorDark='#cccccc'" 
      onmouseout="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'" 
      onmouseover="this.borderColorLight='#cccccc';this.borderColorDark='#000000'" 
      onmouseup="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'">
        <tbody>
          <tr>
            <td bgColor=#ffffff borderColorDark=#ffffff 
            borderColorLight=#ffffff><a 
            href="userchk.asp" target=main><img 
            align=middle alt=注册用户的审核和注销 border=0 height=32 
            src="images/foot.gif" style="FILTER: alpha(opacity=100)" 
            width=32></a> 
            </td>
           </tr>
        </tbody>
        </table>
       </td>
     </tr>
     <tr>
       <td align=middle class=chinese_text13><a 
      href="userchk.asp" target=main>用户管理 
       </a>
       </td>
     </tr>
     <tr>
      <td height=4>
      </td>
     </tr>
     </tbody>
    </table>
    
    
    
<table align=center border=0 cellPadding=1 cellSpacing=0 width="100%">
  <tbody>
    <tr>
      <td align=middle>
        <table bgColor=#ffffff border=1 borderColorDark=#ffffff 
      borderColorLight=#ffffff cellPadding=0 cellSpacing=0 
      onmousedown="this.borderColorLight='#000000';this.borderColorDark='#cccccc'" 
      onmouseout="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'" 
      onmouseover="this.borderColorLight='#cccccc';this.borderColorDark='#000000'" 
      onmouseup="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'">
        <tbody>
          <tr>
            <td bgColor=#ffffff borderColorDark=#ffffff 
            borderColorLight=#ffffff><a 
            href="shouqu.asp" target=main><img 
            align=middle alt=管理各个单位所报送的统计数据 border=0 height=32 
            src="images/attendance.gif" style="FILTER: alpha(opacity=100)" 
            width=32></a> 
            </td>
           </tr>
        </tbody>
        </table>
       </td>
     </tr>
     <tr>
       <td align=middle class=chinese_text13><a 
      href="shouqu.asp" target=main>报文管理 
       </a>
       </td>
     </tr>
     <tr>
      <td height=4>
      </td>
     </tr>
     </tbody>
    </table><table align=center border=0 cellPadding=1 cellSpacing=0 width="100%">
  <tbody>
    <tr>
      <td align=middle>
        <table bgColor=#ffffff border=1 borderColorDark=#ffffff 
      borderColorLight=#ffffff cellPadding=0 cellSpacing=0 
      onmousedown="this.borderColorLight='#000000';this.borderColorDark='#cccccc'" 
      onmouseout="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'" 
      onmouseover="this.borderColorLight='#cccccc';this.borderColorDark='#000000'" 
      onmouseup="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'">
        <tbody>
          <tr>
            <td bgColor=#ffffff borderColorDark=#ffffff 
            borderColorLight=#ffffff><a 
            href="mm.asp" target=main><img 
            align=middle alt=单位管理 border=0 height=32 
            src="images/docu.gif" style="FILTER: alpha(opacity=100)" 
            width=32></a> 
            </td>
           </tr>
        </tbody>
        </table>
       </td>
     </tr>
     <tr>
       <td align=middle class=chinese_text13><a 
      href="mm.asp" target=main>单位管理 
       </a>
       </td>
     </tr>
     <tr>
      <td height=4>
      </td>
     </tr>
     </tbody>
    </table>
  </div>
<div id=uplayer5Div><img height=16 
onmousedown="javascript:this.src='images/scrollup2.gif';menuscrollup()" 
onmouseout="javascript:this.src='images/scrollup.gif';menuscrollstop()" 
onmouseup="javascript:this.src='images/scrollup.gif';menuscrollstop()" 
src="images/scrollup.gif" title=更多 width=16> </div>
<div id=downlayer5Div><img height=16 
onmousedown="javascript:this.src='images/scrolldown2.gif';menuscrolldown()" 
onmouseout="javascript:this.src='images/scrolldown.gif';menuscrollstop()" 
onmouseup="javascript:this.src='images/scrolldown.gif';menuscrollstop()" 
src="images/scrolldown.gif" title=更多 width=16> </div>
<div id=barlayer5Div>
<table bgColor=#cccccc border=0 borderColorDark=#505050 borderColorLight=white 
cellPadding=0 cellSpacing=0 height=22 onclick=javascript:menubarpush(5) 
onmousedown="javascript:this.borderColorDark='White';this.borderColorLight='#505050'" 
onmouseout="javascript:this.borderColorDark='#505050';this.borderColorLight='White'" 
onmouseup="javascript:this.borderColorDark='#505050';this.borderColorLight='White'" 
style="CURSOR: hand" width="100%">
  <tbody>
  <tr>
    <td align=middle background=images/menu_bg.gif borderColorDark=#cccccc 
    borderColorLight=#cccccc colSpan=0 noWrap rowSpan=0><font 
      class=chinese_text13 color=white>超级管理 </font></td></tr></tbody></table></div>
<script id=clientEventHandlersJS language=javascript>

<!--

menuIconWidth[5] = iconlayer5Div.scrollWidth + 0;

menuIconHeight[5] = iconlayer5Div.scrollHeight + 0;

//-->

</script>
</div>
  <div id=menulayer6Div>
  <div id=iconlayer6Div> 
    <table align=center border=0 cellPadding=1 cellSpacing=0 width="100%">
  <tbody>
    <tr>
      <td align=middle>
        <table bgColor=#ffffff border=1 borderColorDark=#ffffff 
      borderColorLight=#ffffff cellPadding=0 cellSpacing=0 
      onmousedown="this.borderColorLight='#000000';this.borderColorDark='#cccccc'" 
      onmouseout="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'" 
      onmouseover="this.borderColorLight='#cccccc';this.borderColorDark='#000000'" 
      onmouseup="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'">
        <tbody>
          <tr>
            <td bgColor=#ffffff borderColorDark=#ffffff 
            borderColorLight=#ffffff><a 
            href="adrot.asp" target=main><img 
            align=middle alt=广告管理 border=0 height=32 
            src="images/docu.gif" style="FILTER: alpha(opacity=100)" 
            width=32></a> 
            </td>
           </tr>
        </tbody>
        </table>
       </td>
     </tr>
     <tr>
       <td align=middle class=chinese_text13><a 
      href="adrot.asp" target=main>广告管理
       </a>
       </td>
     </tr>
     <tr>
      <td height=4>
      </td>
     </tr>
     </tbody>
    </table><table align=center border=0 cellPadding=1 cellSpacing=0 width="100%">
  <tbody>
    <tr>
      <td align=middle>
        <table bgColor=#ffffff border=1 borderColorDark=#ffffff 
      borderColorLight=#ffffff cellPadding=0 cellSpacing=0 
      onmousedown="this.borderColorLight='#000000';this.borderColorDark='#cccccc'" 
      onmouseout="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'" 
      onmouseover="this.borderColorLight='#cccccc';this.borderColorDark='#000000'" 
      onmouseup="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'">
        <tbody>
          <tr>
            <td bgColor=#ffffff borderColorDark=#ffffff 
            borderColorLight=#ffffff><a 
            href="backupdb.asp" target=main><img 
            align=middle alt=数据备份 border=0 height=32 
            src="images/back.gif" style="FILTER: alpha(opacity=100)" 
            width=32></a> 
            </td>
           </tr>
        </tbody>
        </table>
       </td>
     </tr>
     <tr>
       <td align=middle class=chinese_text13><a 
      href="backupdb.asp" target=main>数据备份
       </a>
       </td>
     </tr>
      <tr>
      <td height=4>
      </td>
     </tr>
     </tbody>
    </table><table align=center border=0 cellPadding=1 cellSpacing=0 width="100%">
  <tbody>
    <tr>
      <td align=middle>
        <table bgColor=#ffffff border=1 borderColorDark=#ffffff 
      borderColorLight=#ffffff cellPadding=0 cellSpacing=0 
      onmousedown="this.borderColorLight='#000000';this.borderColorDark='#cccccc'" 
      onmouseout="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'" 
      onmouseover="this.borderColorLight='#cccccc';this.borderColorDark='#000000'" 
      onmouseup="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'">
        <tbody>
          <tr>
            <td bgColor=#ffffff borderColorDark=#ffffff 
            borderColorLight=#ffffff><a 
            href="restoredb.asp" target=main><img 
            align=middle alt=数据恢复 border=0 height=32 
            src="images/restore.gif" style="FILTER: alpha(opacity=100)" 
            width=32></a> 
            </td>
           </tr>
        </tbody>
        </table>
       </td>
     </tr>
     <tr>
       <td align=middle class=chinese_text13><a 
      href="restoredb.asp" target=main>数据恢复
       </a>
       </td>
     </tr>
      <tr>
      <td height=4>
      </td>
     </tr>
     </tbody>
    </table><table align=center border=0 cellPadding=1 cellSpacing=0 width="100%">
  <tbody>
    <tr>
      <td align=middle>
        <table bgColor=#ffffff border=1 borderColorDark=#ffffff 
      borderColorLight=#ffffff cellPadding=0 cellSpacing=0 
      onmousedown="this.borderColorLight='#000000';this.borderColorDark='#cccccc'" 
      onmouseout="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'" 
      onmouseover="this.borderColorLight='#cccccc';this.borderColorDark='#000000'" 
      onmouseup="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'">
        <tbody>
          <tr>
            <td bgColor=#ffffff borderColorDark=#ffffff 
            borderColorLight=#ffffff><a 
            href="sorry.asp" target=main><img 
            align=middle alt=短信管理 border=0 height=32 
            src="images/support.gif" style="FILTER: alpha(opacity=100)" 
            width=32></a> 
            </td>
           </tr>
        </tbody>
        </table>
       </td>
     </tr>
     <tr>
       <td align=middle class=chinese_text13><a 
      href="sorry.asp" target=main>短信管理
       </a>
       </td>
     </tr>
     <tr>
      <td height=4>
      </td>
     </tr>
     <tr>
      <td align=middle>
        <table bgColor=#ffffff border=1 borderColorDark=#ffffff 
      borderColorLight=#ffffff cellPadding=0 cellSpacing=0 
      onmousedown="this.borderColorLight='#000000';this.borderColorDark='#cccccc'" 
      onmouseout="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'" 
      onmouseover="this.borderColorLight='#cccccc';this.borderColorDark='#000000'" 
      onmouseup="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'">
        <tbody>
          <tr>
            <td bgColor=#ffffff borderColorDark=#ffffff 
            borderColorLight=#ffffff><a 
            href="mailto:lj_lw@ynmail.com"><img
            align=middle alt=技术支持 border=0 height=32 
            src="images/suc.gif" style="FILTER: alpha(opacity=100)" 
            width=32></a> 
            </td>
           </tr>
        </tbody>
        </table>
       </td>
     </tr>
     <tr>
       <td align=middle class=chinese_text13><a 
      href="mailto:lj_lw@ynmail.com">技术支持
       </a>
       </td>
     </tr>
     </tbody>
    </table>
    </div>

<div id=uplayer6Div><img height=16 
onmousedown="javascript:this.src='images/scrollup2.gif';menuscrollup()" 
onmouseout="javascript:this.src='images/scrollup.gif';menuscrollstop()" 
onmouseup="javascript:this.src='images/scrollup.gif';menuscrollstop()" 
src="images/scrollup.gif" title=更多 width=16> </div>
<div id=downlayer6Div><img height=16 
onmousedown="javascript:this.src='images/scrolldown2.gif';menuscrolldown()" 
onmouseout="javascript:this.src='images/scrolldown.gif';menuscrollstop()" 
onmouseup="javascript:this.src='images/scrolldown.gif';menuscrollstop()" 
src="images/scrolldown.gif" title=更多 width=16> </div>
<div id=barlayer6Div>
<table bgColor=#cccccc border=0 borderColorDark=#505050 borderColorLight=white 
cellPadding=0 cellSpacing=0 height=22 onclick=javascript:menubarpush(6) 
onmousedown="javascript:this.borderColorDark='White';this.borderColorLight='#505050'" 
onmouseout="javascript:this.borderColorDark='#505050';this.borderColorLight='White'" 
onmouseup="javascript:this.borderColorDark='#505050';this.borderColorLight='White'" 
style="CURSOR: hand" width="100%">
  <tbody>
  <tr>
    <td align=middle background=images/menu_bg.gif borderColorDark=#cccccc 
    borderColorLight=#cccccc colSpan=0 noWrap rowSpan=0><font 
      class=chinese_text13 color=white>系统管理 </font></td></tr></tbody></table></div>

<script id=clientEventHandlersJS language=javascript>

<!--

menuIconWidth[6] = iconlayer6Div.scrollWidth + 0;

menuIconHeight[6] = iconlayer6Div.scrollHeight + 0;

//-->

</script>
<div id=menulayer7Div>
  <div id=iconlayer7Div> 
    <table align=center border=0 cellPadding=1 cellSpacing=0 width="100%">
  <tbody>
    <tr>
      <td align=middle>
        <table bgColor=#ffffff border=1 borderColorDark=#ffffff 
      borderColorLight=#ffffff cellPadding=0 cellSpacing=0 
      onmousedown="this.borderColorLight='#000000';this.borderColorDark='#cccccc'" 
      onmouseout="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'" 
      onmouseover="this.borderColorLight='#cccccc';this.borderColorDark='#000000'" 
      onmouseup="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'">
        <tbody>
          <tr>
            <td bgColor=#ffffff borderColorDark=#ffffff 
            borderColorLight=#ffffff><a 
            href="adrot.asp" target=main><img 
            align=middle alt=广告管理 border=0 height=32 
            src="images/docu.gif" style="FILTER: alpha(opacity=100)" 
            width=32></a> 
            </td>
           </tr>
        </tbody>
        </table>
       </td>
     </tr>
     <tr>
       <td align=middle class=chinese_text13><a 
      href="adrot.asp" target=main>广告管理
       </a>
       </td>
     </tr>
     <tr>
      <td height=4>
      </td>
     </tr>
     </tbody>
    </table><table align=center border=0 cellPadding=1 cellSpacing=0 width="100%">
  <tbody>
    <tr>
      <td align=middle>
        <table bgColor=#ffffff border=1 borderColorDark=#ffffff 
      borderColorLight=#ffffff cellPadding=0 cellSpacing=0 
      onmousedown="this.borderColorLight='#000000';this.borderColorDark='#cccccc'" 
      onmouseout="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'" 
      onmouseover="this.borderColorLight='#cccccc';this.borderColorDark='#000000'" 
      onmouseup="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'">
        <tbody>
          <tr>
            <td bgColor=#ffffff borderColorDark=#ffffff 
            borderColorLight=#ffffff><a 
            href="message_manage.asp" target=main><img 
            align=middle alt=短信管理 border=0 height=32 
            src="images/support.gif" style="FILTER: alpha(opacity=100)" 
            width=32></a> 
            </td>
           </tr>
        </tbody>
        </table>
       </td>
     </tr>
     <tr>
       <td align=middle class=chinese_text13><a 
      href="message_manage.asp" target=main>短信管理
       </a>
       </td>
     </tr>
     <tr>
      <td height=4>
      </td>
     </tr>
     <tr>
      <td align=middle>
        <table bgColor=#ffffff border=1 borderColorDark=#ffffff 
      borderColorLight=#ffffff cellPadding=0 cellSpacing=0 
      onmousedown="this.borderColorLight='#000000';this.borderColorDark='#cccccc'" 
      onmouseout="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'" 
      onmouseover="this.borderColorLight='#cccccc';this.borderColorDark='#000000'" 
      onmouseup="this.borderColorLight='#ffffff';this.borderColorDark='#ffffff'">
        <tbody>
          <tr>
            <td bgColor=#ffffff borderColorDark=#ffffff 
            borderColorLight=#ffffff><a 
            href="mailto:lj_lw@ynmail.com"><img
            align=middle alt=技术支持 border=0 height=32 
            src="images/suc.gif" style="FILTER: alpha(opacity=100)" 
            width=32></a> 
            </td>
           </tr>
        </tbody>
        </table>
       </td>
     </tr>
     <tr>
       <td align=middle class=chinese_text13><a 
      href="mailto:lj_lw@ynmail.com">技术支持
       </a>
       </td>
     </tr>
     </tbody>
    </table>

</div>
<div id=uplayer7Div><img height=16 
onmousedown="javascript:this.src='images/scrollup2.gif';menuscrollup()" 
onmouseout="javascript:this.src='images/scrollup.gif';menuscrollstop()" 
onmouseup="javascript:this.src='images/scrollup.gif';menuscrollstop()" 
src="images/scrollup.gif" title=更多 width=16> </div>
<div id=downlayer7Div><img height=16 
onmousedown="javascript:this.src='images/scrolldown2.gif';menuscrolldown()" 
onmouseout="javascript:this.src='images/scrolldown.gif';menuscrollstop()" 
onmouseup="javascript:this.src='images/scrolldown.gif';menuscrollstop()" 
src="images/scrolldown.gif" title=更多 width=16> </div>
<div id=barlayer7Div>
<table bgColor=#cccccc border=0 borderColorDark=#505050 borderColorLight=white 
cellPadding=0 cellSpacing=0 height=22 onclick=javascript:menubarpush(6) 
onmousedown="javascript:this.borderColorDark='White';this.borderColorLight='#505050'" 
onmouseout="javascript:this.borderColorDark='#505050';this.borderColorLight='White'" 
onmouseup="javascript:this.borderColorDark='#505050';this.borderColorLight='White'" 
style="CURSOR: hand" width="100%">
  <tbody>
  <tr>
    <td align=middle background=images/menu_bg.gif borderColorDark=#cccccc 
    borderColorLight=#cccccc colSpan=0 noWrap rowSpan=0><font 
      class=chinese_text13 color=white>系统管理 </font></td></tr></tbody></table></div>

</div>
<script id=clientEventHandlersJS language=javascript>

<!--

menuIconWidth[7] = iconlayer7Div.scrollWidth + 0;

menuIconHeight[7] = iconlayer7Div.scrollHeight + 0;

//-->

</script>
</body></html>

var lang = new Array();
var userAgent = navigator.userAgent.toLowerCase();
var is_opera = userAgent.indexOf('opera') != -1 && opera.version();
var is_moz = (navigator.product == 'Gecko') && userAgent.substr(userAgent.indexOf('firefox') + 8, 3);
var is_ie = (userAgent.indexOf('msie') != -1 && !is_opera) && userAgent.substr(userAgent.indexOf('msie') + 5, 3);
var is_mac = userAgent.indexOf('mac') != -1;

//FixPrototypeForGecko
if(is_moz && window.HTMLElement) {
	HTMLElement.prototype.__defineSetter__('outerHTML', function(sHTML) {
        	var r = this.ownerDocument.createRange();
		r.setStartBefore(this);
		var df = r.createContextualFragment(sHTML);
		this.parentNode.replaceChild(df,this);
		return sHTML;
	});

	HTMLElement.prototype.__defineGetter__('outerHTML', function() {
		var attr;
		var attrs = this.attributes;
		var str = '<' + this.tagName.toLowerCase();
		for(var i = 0;i < attrs.length;i++){
			attr = attrs[i];
			if(attr.specified)
			str += ' ' + attr.name + '="' + attr.value + '"';
		}
		if(!this.canHaveChildren) {
			return str + '>';
		}
		return str + '>' + this.innerHTML + '</' + this.tagName.toLowerCase() + '>';
        });

	HTMLElement.prototype.__defineGetter__('canHaveChildren', function() {
		switch(this.tagName.toLowerCase()) {
			case 'area':case 'base':case 'basefont':case 'col':case 'frame':case 'hr':case 'img':case 'br':case 'input':case 'isindex':case 'link':case 'meta':case 'param':
			return false;
        	}
		return true;
	});
	HTMLElement.prototype.click = function(){
		var evt = this.ownerDocument.createEvent('MouseEvents');
		evt.initMouseEvent('click', true, true, this.ownerDocument.defaultView, 1, 0, 0, 0, 0, false, false, false, false, 0, null);
		this.dispatchEvent(evt);
	}
}

Array.prototype.push = function(value) {
	this[this.length] = value;
	return this.length;
}

function $(id) {
	return document.getElementById(id);
}

// 完成操作,停止事件, 取消事件的默认动作
function doane(event) {
	e = event ? event : window.event;
	if(is_ie) {
		e.returnValue = false;
		e.cancelBubble = true;
	} else if(e) {
		e.stopPropagation();
		e.preventDefault();
	}
}

function fetchCheckbox(cbn) {
	return $(cbn) && $(cbn).checked == true ? 1 : 0;
}

// 取得指定cookie值
function getcookie(name) {
	var cookie_start = document.cookie.indexOf(name);
	var cookie_end = document.cookie.indexOf(";", cookie_start);
	return cookie_start == -1 ? '' : unescape(document.cookie.substring(cookie_start + name.length + 1, (cookie_end > cookie_start ? cookie_end : document.cookie.length)));
}

// 设置cookie
function setcookie(cookieName, cookieValue, seconds, path, domain, secure) {
	var expires = new Date();
	expires.setTime(expires.getTime() + seconds);
	domain = !domain ? cookiedomain : domain;
	path = !path ? cookiepath : path;
	document.cookie = escape(cookieName) + '=' + escape(cookieValue)
		+ (expires ? '; expires=' + expires.toGMTString() : '')
		+ (path ? '; path=' + path : '/')
		+ (domain ? '; domain=' + domain : '')
		+ (secure ? '; secure' : '');
}

function in_array(needle, haystack) {
	if(typeof needle == 'string' || typeof needle == 'number') {
		for(var i in haystack) {
			if(haystack[i] == needle) {
					return true;
			}
		}
	}
	return false;
}

function isUndefined(variable) {
	return typeof variable == 'undefined' ? true : false;
}

// 取得多字节字符串长度
function mb_strlen(str) {
	var len = 0;
	for(var i = 0; i < str.length; i++) {
		len += str.charCodeAt(i) < 0 || str.charCodeAt(i) > 255 ? (charset == 'utf-8' ? 3 : 2) : 1;
	}
	return len;
}

// 截取多字节字符串到指定长度
function mb_cutstr(str, maxlen, dot) {
	var len = 0;
	var ret = '';
	var dot = !dot ? '...' : '';
	maxlen = maxlen - dot.length;
	for(var i = 0; i < str.length; i++) {
		len += str.charCodeAt(i) < 0 || str.charCodeAt(i) > 255 ? (charset == 'utf-8' ? 3 : 2) : 1;
		if(len > maxlen) {
			ret += dot;
			break;
		}
		ret += str.substr(i, 1);
	}
	return ret;
}

function strlen(str) {
	return (is_ie && str.indexOf('\n') != -1) ? str.replace(/\r?\n/g, '_').length : str.length;
}

function updatestring(str1, str2, clear) {
	str2 = '_' + str2 + '_';
	return clear ? str1.replace(str2, '') : (str1.indexOf(str2) == -1 ? str1 + str2 : str1);
}

// 删除字符串两边的空格
function trim(str) {
	return (str + '').replace(/(\s+)$/g, '').replace(/^\s+/g, '');
}

function _attachEvent(obj, evt, func, eventobj) {
	eventobj = !eventobj ? obj : eventobj;
	if(obj.addEventListener) {
		obj.addEventListener(evt, func, false);
	} else if(eventobj.attachEvent) {
		obj.attachEvent("on" + evt, func);
	}
}

var jsmenu = new Array();
var ctrlobjclassName;
jsmenu['active'] = new Array();
jsmenu['timer'] = new Array();
jsmenu['iframe'] = new Array();

function initCtrl(ctrlobj, click, duration, timeout, layer) {
	if(ctrlobj && !ctrlobj.initialized) {
		ctrlobj.initialized = true;
		ctrlobj.unselectable = true;

		ctrlobj.outfunc = typeof ctrlobj.onmouseout == 'function' ? ctrlobj.onmouseout : null;
		ctrlobj.onmouseout = function() {
			if(this.outfunc) this.outfunc();
			if(duration < 3) jsmenu['timer'][ctrlobj.id] = setTimeout('hideMenu(' + layer + ')', timeout);
		}

		ctrlobj.overfunc = typeof ctrlobj.onmouseover == 'function' ? ctrlobj.onmouseover : null;
		ctrlobj.onmouseover = function(e) {
			doane(e);
			if(this.overfunc) this.overfunc();
			if(click) {
				clearTimeout(jsmenu['timer'][this.id]);
			} else {
				for(var id in jsmenu['timer']) {
					if(jsmenu['timer'][id]) clearTimeout(jsmenu['timer'][id]);
				}
			}
		}
	}
}

function initMenu(ctrlid, menuobj, duration, timeout, layer, drag) {
	if(menuobj && !menuobj.initialized) {
		menuobj.initialized = true;
		menuobj.ctrlkey = ctrlid;
		menuobj.onclick = ebygum;
		menuobj.style.position = 'absolute';
		if(duration < 3) {
			if(duration > 1) {
				menuobj.onmouseover = function() {
					clearTimeout(jsmenu['timer'][ctrlid]);
				}
			}
			if(duration != 1) {
				menuobj.onmouseout = function() {
					jsmenu['timer'][ctrlid] = setTimeout('hideMenu(' + layer + ')', timeout);
				}
			}
		}
		menuobj.style.zIndex = 999;
		if(drag) {
			menuobj.onmousedown = function(event) {try{menudrag(menuobj, event, 1);}catch(e){}};
			menuobj.onmousemove = function(event) {try{menudrag(menuobj, event, 2);}catch(e){}};
			menuobj.onmouseup = function(event) {try{menudrag(menuobj, event, 3);}catch(e){}};
		}
	}
}

var menudragstart = new Array();
function menudrag(menuobj, e, op) {
	if(op == 1) {
		if(in_array(is_ie ? event.srcElement.tagName : e.target.tagName, ['TEXTAREA', 'INPUT', 'BUTTON', 'SELECT'])) {
			return;
		}
		menudragstart = is_ie ? [event.clientX, event.clientY] : [e.clientX, e.clientY];
		menudragstart[2] = parseInt(menuobj.style.left);
		menudragstart[3] = parseInt(menuobj.style.top);
		doane(e);
	} else if(op == 2 && menudragstart[0]) {
		var menudragnow = is_ie ? [event.clientX, event.clientY] : [e.clientX, e.clientY];
		menuobj.style.left = (menudragstart[2] + menudragnow[0] - menudragstart[0]) + 'px';
		menuobj.style.top = (menudragstart[3] + menudragnow[1] - menudragstart[1]) + 'px';
		doane(e);
	} else if(op == 3) {
		menudragstart = [];
		doane(e);
	}
}

function showMenu(ctrlid, click, offset, duration, timeout, layer, showid, maxh, drag) {
	var ctrlobj = $(ctrlid);
	if(!ctrlobj) return;
	if(isUndefined(click)) click = false;
	if(isUndefined(offset)) offset = 0;
	if(isUndefined(duration)) duration = 2;
	if(isUndefined(timeout)) timeout = 250;
	if(isUndefined(layer)) layer = 0;
	if(isUndefined(showid)) showid = ctrlid;
	var showobj = $(showid);
	var menuobj = $(showid + '_menu');
	if(!showobj|| !menuobj) return;
	if(isUndefined(maxh)) maxh = 400;
	if(isUndefined(drag)) drag = false;

	if(click && jsmenu['active'][layer] == menuobj) {
		hideMenu(layer);
		return;
	} else {
		hideMenu(layer);
	}

	var len = jsmenu['timer'].length;
	if(len > 0) {
		for(var i=0; i<len; i++) {
			if(jsmenu['timer'][i]) clearTimeout(jsmenu['timer'][i]);
		}
	}

	initCtrl(ctrlobj, click, duration, timeout, layer);
	ctrlobjclassName = ctrlobj.className;
	ctrlobj.className += ' hover';
	initMenu(ctrlid, menuobj, duration, timeout, layer, drag);

	menuobj.style.display = '';
	if(!is_opera) {
		menuobj.style.clip = 'rect(auto, auto, auto, auto)';
	}

	setMenuPosition(showid, offset);

	if(maxh && menuobj.scrollHeight > maxh) {
		menuobj.style.height = maxh + 'px';
		if(is_opera) {
			menuobj.style.overflow = 'auto';
		} else {
			menuobj.style.overflowY = 'auto';
		}
	}

	if(!duration) {
		setTimeout('hideMenu(' + layer + ')', timeout);
	}

	jsmenu['active'][layer] = menuobj;
}

function setMenuPosition(showid, offset) {
	var showobj = $(showid);
	var menuobj = $(showid + '_menu');
	if(isUndefined(offset)) offset = 0;
	if(showobj) {
		showobj.pos = fetchOffset(showobj);
		showobj.X = showobj.pos['left'];
		showobj.Y = showobj.pos['top'];
		if($(InFloat) != null) {
			var InFloate = InFloat.split('_');
			if(!floatwinhandle[InFloate[1] + '_1']) {
				floatwinnojspos = fetchOffset($('floatwinnojs'));
				floatwinhandle[InFloate[1] + '_1'] = floatwinnojspos['left'];
				floatwinhandle[InFloate[1] + '_2'] = floatwinnojspos['top'];
			}
			showobj.X = showobj.X - $(InFloat).scrollLeft - parseInt(floatwinhandle[InFloate[1] + '_1']);
			showobj.Y = showobj.Y - $(InFloat).scrollTop - parseInt(floatwinhandle[InFloate[1] + '_2']);
			InFloat = '';
		}
		showobj.w = showobj.offsetWidth;
		showobj.h = showobj.offsetHeight;
		menuobj.w = menuobj.offsetWidth;
		menuobj.h = menuobj.offsetHeight;
		if(offset < 3) {
			menuobj.style.left = (showobj.X + menuobj.w > document.body.clientWidth) && (showobj.X + showobj.w - menuobj.w >= 0) ? showobj.X + showobj.w - menuobj.w + 'px' : showobj.X + 'px';
			menuobj.style.top = offset == 1 ? showobj.Y + 'px' : (offset == 2 || ((showobj.Y + showobj.h + menuobj.h > document.documentElement.scrollTop + document.documentElement.clientHeight) && (showobj.Y - menuobj.h >= 0)) ? (showobj.Y - menuobj.h) + 'px' : showobj.Y + showobj.h + 'px');
		} else if(offset == 3) {
			menuobj.style.left = (document.body.clientWidth - menuobj.clientWidth) / 2 + document.body.scrollLeft + 'px';
			menuobj.style.top = (document.body.clientHeight - menuobj.clientHeight) / 2 + document.body.scrollTop + 'px';
		}
		
		if(menuobj.style.clip && !is_opera) {
			menuobj.style.clip = 'rect(auto, auto, auto, auto)';
		}
	}
}

function hideMenu(layer) {
	if(isUndefined(layer)) layer = 0;
	if(jsmenu['active'][layer]) {
		try {
			$(jsmenu['active'][layer].ctrlkey).className = ctrlobjclassName;
		} catch(e) {}
		clearTimeout(jsmenu['timer'][jsmenu['active'][layer].ctrlkey]);
		jsmenu['active'][layer].style.display = 'none';
		if(is_ie && is_ie < 7 && jsmenu['iframe'][layer]) {
			jsmenu['iframe'][layer].style.display = 'none';
		}
		jsmenu['active'][layer] = null;
	}
}

function fetchOffset(obj) {
	var left_offset = obj.offsetLeft;
	var top_offset = obj.offsetTop;
	while((obj = obj.offsetParent) != null) {
		left_offset += obj.offsetLeft;
		top_offset += obj.offsetTop;
	}
	return { 'left' : left_offset, 'top' : top_offset };
}

function ebygum(eventobj) {
	if(!eventobj || is_ie) {
		window.event.cancelBubble = true;
		return window.event;
	} else {
		if(eventobj.target.type == 'submit') {
			eventobj.target.form.submit();
		}
		eventobj.stopPropagation();
		return eventobj;
	}
}

//得到一个定长的hash值， 依赖于 stringxor()
function hash(string, length) {
	var length = length ? length : 32;
	var start = 0;
	var i = 0;
	var result = '';
	filllen = length - string.length % length;
	for(i = 0; i < filllen; i++){
		string += "0";
	}
	while(start < string.length) {
		result = stringxor(result, string.substr(start, length));
		start += length;
	}
	return result;
}

function stringxor(s1, s2) {
	var s = '';
	var hash = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ';
	var max = Math.max(s1.length, s2.length);
	for(var i=0; i<max; i++) {
		var k = s1.charCodeAt(i) ^ s2.charCodeAt(i);
		s += hash.charAt(k % 52);
	}
	return s;
}


//PageScroll
function pagescroll_class(obj, pagewidth, pageheight) {
	this.ctrlobj = $(obj);
	this.speed = 2;
	this.pagewidth = pagewidth;
	this.times = 1;
	this.pageheight = pageheight;
	this.running = 0;
	this.defaultleft = 0;
	this.defaulttop = 0;
	this.script = '';
	this.start = function(times) {
		if(this.running) return 0;
		this.times = !times ? 1 : times;
		this.scrollpx = 0;
		return this.running = 1;
	}
	this.left = function(times, script) {
		if(!this.start(times)) return;
		this.stepv = -(this.step = this.pagewidth * this.times / this.speed);
		this.script = !script ? '' : script;
		setTimeout('pagescroll.h()', 1);
	}
	this.right = function(times, script) {
		if(!this.start(times)) return;
		this.stepv = this.step = this.pagewidth * this.times / this.speed;
		this.script = !script ? '' : script;
		setTimeout('pagescroll.h()', 1);
	}
	this.up = function(times, script) {
		if(!this.start(times)) return;
		this.stepv = -(this.step = this.pageheight * this.times / this.speed);
		this.script = !script ? '' : script;
		setTimeout('pagescroll.v()', 1);
	}
	this.down = function(times, script) {
		if(!this.start(times)) return;
		this.stepv = this.step = this.pageheight * this.times / this.speed;
		this.script = !script ? '' : script;
		setTimeout('pagescroll.v()', 1);
	}
	this.h = function() {
		if(this.scrollpx <= this.pagewidth * this.times) {
			this.scrollpx += Math.abs(this.stepv);
			patch = this.scrollpx > this.pagewidth * this.times ? this.scrollpx - this.pagewidth * this.times : 0;
			patch = patch > 0 && this.stepv < 0 ? -patch : patch;
			oldscrollLeft = this.ctrlobj.scrollLeft;
			this.ctrlobj.scrollLeft = this.ctrlobj.scrollLeft + this.stepv - patch;
			if(oldscrollLeft != this.ctrlobj.scrollLeft) {
				setTimeout('pagescroll.h()', 1);
				return;
			}
		}
		if(this.script) {
			eval(this.script);
		}
		this.running = 0;
	}
	this.v = function() {
		if(this.scrollpx <= this.pageheight * this.times) {
			this.scrollpx += Math.abs(this.stepv);
			patch = this.scrollpx > this.pageheight * this.times ? this.scrollpx - this.pageheight * this.times : 0;
			patch = patch > 0 && this.stepv < 0 ? -patch : patch;
			oldscrollTop = this.ctrlobj.scrollTop;
			this.ctrlobj.scrollTop = this.ctrlobj.scrollTop + this.stepv - patch;
			if(oldscrollTop != this.ctrlobj.scrollTop) {
				setTimeout('pagescroll.v()', 1);
				return;
			}
		}
		if(this.script) {
			eval(this.script);
		}
		this.running = 0;
	}
	this.init = function() {
		this.ctrlobj.scrollLeft = this.defaultleft;
		this.ctrlobj.scrollTop = this.defaulttop;
	}

}

//FloatWin
var hiddenobj = new Array();
var floatwinhandle = new Array();
var floatscripthandle = new Array();
var floattabs = new Array();
var floatwins = new Array();
var InFloat = '';
var floatwinreset = 0;
var floatwinopened = 0;

function floatwin_scroll(pos) {
	var pose = pos.split(',');
	try {
		pagescroll.defaultleft = pose[0];
		pagescroll.defaulttop = pose[1];
		pagescroll.init();
	} catch(e) {}
}

//Smilies
function smilies_show(id, smcols, method, seditorkey) {
	if(seditorkey && !$(seditorkey + 'smilies_menu')) {
		var div = document.createElement("div");
		div.id = seditorkey + 'smilies_menu';
		div.style.display = 'none';
		div.className = 'smilieslist';
		$('append_parent').appendChild(div);
		var div = document.createElement("div");
		div.id = id;
		div.style.overflow = 'hidden';
		$(seditorkey + 'smilies_menu').appendChild(div);
	}
	if(typeof smilies_type == 'undefined') {
		var scriptNode = document.createElement("script");
		scriptNode.type = "text/javascript";
		scriptNode.charset = charset ? charset : (is_moz ? document.characterSet : document.charset);
		scriptNode.src = 'images/smilies_var.js';
		$('append_parent').appendChild(scriptNode);
		if(is_ie) {
			scriptNode.onreadystatechange = function() {
				smilies_onload(id, smcols, method, seditorkey);
			}
		} else {
			scriptNode.onload = function() {
				smilies_onload(id, smcols, method, seditorkey);
			}
		}
	} else {
		smilies_onload(id, smcols, method, seditorkey);
	}
}

var currentstype = null;
function smilies_onload(id, smcols, method, seditorkey) {
	smile = getcookie('smile').split('D');
	if(typeof smilies_type != 'undefined') {
		currentstype = smile[0] ? smile[0] : 1;
		smiliestype = '<div class="smiliesgroup" style="margin-right: 0"><ul>';
		for(i in smilies_type) {
			smiliestype += '<li><a href="javascript:;" hidefocus="ture" ' + (currentstype == i ? 'class="current"' : '') + ' id="stype'+i+'" onclick="smilies_switch(\'' + id + '\', \'' + smcols + '\', '+i+', 1, ' + method + ', \'' + seditorkey + '\');if(currentstype) {$(\'stype\'+currentstype).className=\'\';}this.className=\'current\';currentstype='+i+';">'+smilies_type[i][0]+'</a></li>';
		}
		smiliestype += '</ul></div>';
		$(id).innerHTML = smiliestype + '<div style="clear: both" class="float_typeid" id="' + id + '_data"></div><table class="smilieslist_table" id="' + id + '_preview_table" style="display: none"><tr><td class="smilieslist_preview" id="' + id + '_preview"></td></tr></table>';
		smilies_switch(id, smcols, smile[0], smile[1], method, seditorkey);
	}
}

function smilies_switch(id, smcols, type, page, method, seditorkey) {
	type = type? type : 1;
	page = page ? page : 1;
	setcookie('smile', type + 'D' + page, 31536000);
	smiliesdata = '<table id="' + id + '_table" cellpadding="0" cellspacing="0" style="clear: both"><tr>';
	j = 0;
	for(i in smilies_array[type][page]) {
		if(j >= smcols) {
			smiliesdata += '<tr>';
			j = 0;
		}
		s = smilies_array[type][page][i];
		smiliesdata += s && s[0] ? '<td onmouseover="smilies_preview(\'' + id + '\', this, ' + s[5] + ')" onmouseout="smilies_preview(\'' + id + '\')" onclick="' + (method ? 'insertSmiley(' + s[0] + ')': 'seditor_insertunit(\'' + seditorkey + '\', \'' + s[1].replace(/'/, '\\\'') + '\')') +
			'"><img id="smilie_' + s[0] + '" width="' + s[3] +'" height="' + s[4] +'" src="images/smilies/' + smilies_type[type][1] + '/' + s[2] + '" alt="' + s[1] + '" />' : '<td>';
		j++;
	}
	smiliesdata += '</table>';
	$(id + '_data').innerHTML = smiliesdata;
}

function smilies_preview(id, obj, v) {
	if(!obj) {
		$(id + '_preview_table').style.display = 'none';
	} else {
		$(id + '_preview_table').style.display = '';
		$(id + '_preview').innerHTML = '<img width="' + v + '" src="' + obj.childNodes[0].src + '" />';
	}
}

//SEditor
function seditor_ctlent(event, script) {
	if(event.ctrlKey && event.keyCode == 13 || event.altKey && event.keyCode == 83) {
		eval(script);
	}
}

function seditor_insertunit(key, text, textend, moveend) {	
	$(key + 'message').focus();
	textend = isUndefined(textend) ? '' : textend;
	moveend = isUndefined(textend) ? 0 : moveend;
	startlen = strlen(text);
	endlen = strlen(textend);
	if(!isUndefined($(key + 'message').selectionStart)) {
		var opn = $(key + 'message').selectionStart + 0;
		if(textend != '') {
			text = text + $(key + 'message').value.substring($(key + 'message').selectionStart, $(key + 'message').selectionEnd) + textend;
		}
		$(key + 'message').value = $(key + 'message').value.substr(0, $(key + 'message').selectionStart) + text + $(key + 'message').value.substr($(key + 'message').selectionEnd);
		if(!moveend) {
			$(key + 'message').selectionStart = opn + strlen(text) - endlen;
			$(key + 'message').selectionEnd = opn + strlen(text) - endlen;
		}
	} else if(document.selection && document.selection.createRange) {
		var sel = document.selection.createRange();
		if(textend != '') {
			text = text + sel.text + textend;
		}
		sel.text = text.replace(/\r?\n/g, '\r\n');
		if(!moveend) {
			sel.moveStart('character', -endlen);
			sel.moveEnd('character', -endlen);
		}
		sel.select();
	} else {
		$(key + 'message').value += text;
	}
	hideMenu();
}

var cookiedomain = isUndefined(cookiedomain) ? '' : cookiedomain;
var cookiepath = isUndefined(cookiepath) ? '' : cookiepath;
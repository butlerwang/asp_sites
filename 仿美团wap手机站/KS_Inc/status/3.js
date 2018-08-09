   var place=1; 
	function scrollIn() {
	window.status=Message.substring(0, place);
	if (place >= Message.length) {
	place=1;
	window.setTimeout("KS_Status3()",300);
	} else {
	place++;
	window.setTimeout("scrollIn()",speed);
	}
	}
	function KS_Status3() {
	window.status=Message.substring(place, Message.length);
	if (place >= Message.length) {
	place=1;
	window.setTimeout("scrollIn()", 100);
	} else {
	place++;
	window.setTimeout("KS_Status3()", speed);
	}
	}
	KS_Status3();
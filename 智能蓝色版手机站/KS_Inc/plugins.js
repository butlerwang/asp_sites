﻿/////////////////// Plug-in file for CalendarXP 6.0 /////////////////
// This file is totally configurable. You may remove all the comments in this file to shrink the download size.
/////////////////////////////////////////////////////////////////////

///////////// Calendar Onchange Handler ////////////////////////////
// It's triggered whenever the calendar gets changed.
// d = 0 means the calendar is about to switch to the month of (y,m); 
// d > 0 means a specific date [y,m,d] is about to be selected.
////////////////////////////////////////////////////////////////////
function fOnChange(y,m,d) {
	return false;  // return true to cancel the change.
	''//alert("sdfsadlsdfa ")
	}



///////////// Calendar AfterSelected Handler ///////////////////////
// It's triggered whenever a date gets fully selected.
// The selected date is passed in as y(ear),m(onth),d(ay)
////////////////////////////////////////////////////////////////////
function fAfterSelected(y,m,d) {
''//var s;
//if (m<9) {
//m='0'+m
//}
//if (d<9) {
//d='0'+d
//}
//s=y+m+d
//window.open('helloworld.htm#'+s)
//window.location='helloworld.htm'
}

// ====== Following are self-defined and/or custom-built functions! =======

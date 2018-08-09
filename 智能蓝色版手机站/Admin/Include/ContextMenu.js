var DocPopupContextMenu=window.createPopup();
function document.oncontextmenu()
{	var width=120;
	var height=0;
	var left=event.clientX;
	var top=event.clientY;
	 DocDisabledContextMenu();
	var ObjPopDocument=DocPopupContextMenu.document;
	var ContextMenuStr='';
	for (var i=0;i<DocMenuArr.length;i++)
	{
		if (DocMenuArr[i].ExeFunction=='seperator')
		{
			ContextMenuStr+=FormatSeperator();
			height+=16;
		}
		else
		{
			ContextMenuStr+=FormatMenuRow(DocMenuArr[i].ExeFunction,DocMenuArr[i].Description.replace('\(','\(<U>').replace(')','</U>)'),DocMenuArr[i].EnabledStr);
			height+=20;
		}
	}
	ContextMenuStr="<TABLE border=0 cellpadding=0 cellspacing=0 class=Menu width=120>"+ContextMenuStr
	ContextMenuStr=ContextMenuStr+"<\/TABLE>";
	ObjPopDocument.open();
	ObjPopDocument.write("<link href=\"../Include/ContextMenu.css\" type=\"text/css\" rel=\"stylesheet\"><body scroll=\"no\" onConTextMenu=\"window.event.returnValue=false;\" onselectstart=\"window.event.returnValue=false;\">"+ContextMenuStr);
	ObjPopDocument.close();
	height+=4;
	if(left+width > document.body.clientWidth) left-=width;
	if(top+height > document.body.clientHeight) top-=height;
	DocPopupContextMenu.show(left, top, width, height, document.body);
	return false;
}
function FormatSeperator()
{
	var MenuRowStr="<tr><td height=16 valign=middle ><hr class=\"Seperator\" width=95%><\/td><\/tr>";
	return MenuRowStr;
}
function FormatMenuRow(MenuOperation,MenuDescription,EnabledStr)
{
	var MenuRowStr="<tr "+EnabledStr+"><td align=left height=20 class=MouseOut onMouseOver=this.className='MouseOver'; onMouseOut=this.className='MouseOut'; valign=middle"
	if (EnabledStr=='') MenuRowStr+=" onclick=\""+MenuOperation+"parent.DocPopupContextMenu.hide();\">&nbsp;&nbsp;&nbsp;&nbsp;";
	else MenuRowStr+=">&nbsp;&nbsp;&nbsp;&nbsp;";
	MenuRowStr=MenuRowStr+MenuDescription+"<\/td><\/tr>";
	return MenuRowStr;
}
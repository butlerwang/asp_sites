self.onError=null; 
currentX = currentY = 0; 
CurrElement = null; 
function mousedown(El) { 
CurrElement = El; 
currentX = (event.clientX + document.body.scrollLeft); 
currentY = (event.clientY + document.body.scrollTop);   
return true; 
} 
/*
function document.onmousemove() { 
if (CurrElement == null) { return false; } 
newX = (event.clientX + document.body.scrollLeft); 
newY = (event.clientY + document.body.scrollTop); 
distanceX = (newX - currentX); 
distanceY = (newY - currentY); 
currentX = newX; currentY = newY; 
CurrElement.style.pixelLeft += distanceX; 
CurrElement.style.pixelTop += distanceY;
 event.returnvalue = false; 
return false; 
} 
function document.onmouseup() { 
CurrElement = null; 
return true; 
 }
 */
 
/*********************************************************************************** 
函 数: SelectElement 
选 择:无
作 用:选择对象(目录或文件)
*************************************************************************************/
function SelectElement()
{  
  var el=event.srcElement;
	var i=0;
	//alert(el.tagName);
	//alert('前一个'+el.parentElement.tagName);
	if ((event.ctrlKey==true)||(event.shiftKey==true))
	{
	  if (event.ctrlKey==true)
		{
			for (i=0;i<DocElementArr.length;i++)
              {
				if (el.parentElement==DocElementArr[i].Obj)
				{
					if (DocElementArr[i].Selected==false)
					 { if (el.tagName=='IMG')
					     {el.className='FolderSelected';el.parentElement.children[1].className='FolderSelectItem';}
					    else
					     {el.className='FolderSelectItem';el.parentElement.children[0].className='FolderSelected';}
		                 DocElementArr[i].Selected=true;
					  }
					else
					{ DocElementArr[i].Obj.children[0].className='';DocElementArr[i].Obj.children[1].className='FolderItem';DocElementArr[i].Selected=false;}
				}
			}
		}
		if (event.shiftKey==true)
		{ var MaxIndex=0,ObjInArray=false,EndIndex=0,ElIndex=-1;
			for (i=0;i<DocElementArr.length;i++)
			{
				if (DocElementArr[i].Selected==true)
				{if (DocElementArr[i].Index>=MaxIndex) MaxIndex=DocElementArr[i].Index;}
				if (el.parentElement==DocElementArr[i].Obj)
				{
				  ObjInArray=true;
				  ElIndex=i;
				  EndIndex=DocElementArr[i].Index;
				}
			}
			if (ElIndex>MaxIndex)
			{
				if (MaxIndex>0)
				   {if (el.tagName=='IMG')
				 	   for (i=MaxIndex-1;i<EndIndex;i++)
					   { DocElementArr[i].Obj.children[0].className='FolderSelected';DocElementArr[i].Obj.children[1].className='FolderSelectItem';DocElementArr[i].Selected=true;}
					 else
					    for(i=MaxIndex-1;i<EndIndex;i++)
					    {DocElementArr[i].Obj.children[1].className='FolderSelectItem';DocElementArr[i].Obj.children[0].className='FolderSelected';DocElementArr[i].Selected=true;}
					}
				else
				{  
                        if (el.tagName=='IMG')
					     {el.className='FolderSelected';DocElementArr[ElIndex].Obj.children[1].className='FolderSelectItem';}
					    else
					     {el.className='FolderSelectItem';DocElementArr[ElIndex].Obj.children[0].className='FolderSelected';}
		                 DocElementArr[ElIndex].Selected=true;
				}
			}
			else
			{  
				if (ObjInArray)
				{
					  if (el.tagName=='IMG')
					    for (i=EndIndex;i<MaxIndex-1;i++)
					     {el.className='FolderSelected';DocElementArr[i].Obj.children[1].className='FolderSelectItem';DocElementArr[i].Selected=true;}
					  else
					  	for (i=EndIndex;i<MaxIndex-1;i++)
					     {el.className='FolderSelectItem';DocElementArr[i].Obj.children[0].className='FolderSelected';DocElementArr[i].Selected=true;}
					if (ElIndex>=0)
					{
                        if (el.tagName=='IMG')
					     {el.className='FolderSelected';DocElementArr[ElIndex].Obj.children[1].className='FolderSelectItem';}
					    else
					     {el.className='FolderSelectItem';DocElementArr[ElIndex].Obj.children[0].className='FolderSelected';}
		                 DocElementArr[ElIndex].Selected=true;
					}
				}
			}
		}
    }	
	else
	{
		for (i=0;i<DocElementArr.length;i++)
		 {
			if (el.parentElement==DocElementArr[i].Obj)
			    {  if (el.tagName=='IMG')
					   {el.className='FolderSelected';el.parentElement.children[1].className='FolderSelectItem';}
				   else
					   {el.className='FolderSelectItem';el.parentElement.children[0].className='FolderSelected';}
		            DocElementArr[i].Selected=true;
			    }
			else
			{DocElementArr[i].Obj.children[0].className='';DocElementArr[i].Obj.children[1].className='FolderItem';
		     DocElementArr[i].Selected=false;
			}
		}
	}
}
 /*********************************************************************************** 
函 数: SelectElement 
选 择:无
作 用:选择对象(目录或文件)
*************************************************************************************/
function SelectAllElement()
{
		for (i=0;i<DocElementArr.length;i++)
		 { 
		  DocElementArr[i].Obj.children[0].className='FolderSelected';
		  DocElementArr[i].Obj.children[1].className='FolderSelectItem';
		  DocElementArr[i].Selected=true;
		  }
}
function ContextMenuItem(ExeFunction,Description,EnabledStr)
{
	this.ExeFunction=ExeFunction;
	this.Description=Description;
	this.EnabledStr=EnabledStr;
}
function ElementObj(Obj,Index,Selected)
{
	this.Obj=Obj;
	this.Index=Index;
	this.Selected=Selected;
}
/**********************************************************************
函 数:InitialDocElementArr
参 数:FolderID--目录ID,FileID----- 文件ID,
功 能:初始化对象
*************************************************************************/
function InitialDocElementArr(FolderID,FileID)
{   var CurrObj=null,j=1;
	var AllElement=document.body.getElementsByTagName('span');
	for (var i=0;i<AllElement.length;i++)
	{
		CurrObj=AllElement(i);
		if ((eval('CurrObj.'+FolderID)!=null)||(eval('CurrObj.'+FileID)!=null))
		{
			DocElementArr[DocElementArr.length]=new ElementObj(CurrObj,j,false);
			j++;
		}
	}
}
/**********************************************************************************************
函 数:DisabledContextMenu
参 数:FolderID目录ID,FileID 文件ID,
     BNS---Both None Selected FSS---File SIngal Selected,FMS---File Multiple Selected
     BSS---Both Selected, DSS---Directory singal Selected,DMS---Directory Multiple Selected
作 用:根据选择对象不同,决定菜单项是否有效	 
**************************************************************************************************/
function DisabledContextMenu(FolderID,FileID,BNS,FSS,FMS,BSS,DSS,DMS)
{	var el=event.srcElement;EventObjInArray=false,SelectFolder='',SelectFile='',DisabledContextMenuStr='';
	for (var i=0;i<DocElementArr.length;i++)
	{
		if (el.parentElement==DocElementArr[i].Obj)
		{
			if (DocElementArr[i].Selected==true) EventObjInArray=true;
			break;
		}
	}
	for (var i=0;i<DocElementArr.length;i++)
	{
		if (el.parentElement==DocElementArr[i].Obj)
		{    DocElementArr[i].Obj.children[0].className='FolderSelected';
		     DocElementArr[i].Obj.children[1].className='FolderSelectItem';
			 DocElementArr[i].Selected=true;
			if (eval('DocElementArr[i].Obj.'+FolderID)!=null)
			{
				if (SelectFolder=='') SelectFolder=eval('DocElementArr[i].Obj.'+FolderID);
				else SelectFolder=SelectFolder+','+eval('DocElementArr[i].Obj.'+FolderID)
			}
			if (eval('DocElementArr[i].Obj.'+FileID)!=null)
			{
				if (SelectFile=='') SelectFile=eval('DocElementArr[i].Obj.'+FileID);
				else SelectFile=SelectFile+','+eval('DocElementArr[i].Obj.'+FileID)
			}
		}
		else
		{
			if (!EventObjInArray)
			{   DocElementArr[i].Obj.children[0].className='';
		        DocElementArr[i].Obj.children[1].className='FolderItem';
				DocElementArr[i].Selected=false;
			}
			else
			{
				if (DocElementArr[i].Selected==true)
				{
					if (eval('DocElementArr[i].Obj.'+FolderID)!=null)
					{
						if (SelectFolder=='') SelectFolder=eval('DocElementArr[i].Obj.'+FolderID);
						else SelectFolder=SelectFolder+','+eval('DocElementArr[i].Obj.'+FolderID)
					}
					if (eval('DocElementArr[i].Obj.'+FileID)!=null)
					{
						if (SelectFile=='') SelectFile=eval('DocElementArr[i].Obj.'+FileID);
						else SelectFile=SelectFile+','+eval('DocElementArr[i].Obj.'+FileID)
					}
				}
			}
		}
	}
	if ((SelectFolder=='')&&(SelectFile=='')) DisabledContextMenuStr=BNS;
	else
	{
		if ((SelectFile!='')&&(SelectFolder==''))
		{
			if (SelectFile.indexOf(',')!=-1) 
			   DisabledContextMenuStr=FSS;
			else DisabledContextMenuStr=FMS;
		}
		if ((SelectFolder!='')&&(SelectFile!='')) DisabledContextMenuStr=DisabledContextMenuStr+BSS;
		if ((SelectFolder!='')&&(SelectFile==''))
		{
			if (SelectFolder.indexOf(',')!=-1) 
			  DisabledContextMenuStr=DSS;
			else DisabledContextMenuStr=DMS;
		}
	}
	for (var i=0;i<DocMenuArr.length;i++)
	{
		if (DisabledContextMenuStr.indexOf(DocMenuArr[i].Description)!=-1) DocMenuArr[i].EnabledStr='disabled';
		else  DocMenuArr[i].EnabledStr='';
	}
}
/************************************************************************
函 数:GetSelectStatus
参 数:FolderID---栏目或目录ID,FileID----文件ID
作 用: 返回当前选择项的ID集合,包括目录和文件
**************************************************************************/
function GetSelectStatus(FolderID,FileID)
{
  for (var i=0;i<DocElementArr.length;i++)
	{
		if (DocElementArr[i].Selected==true)
		{
			if (eval('DocElementArr[i].Obj.'+FileID)!=null)
			{
				if (SelectedFile=='') SelectedFile=eval('DocElementArr[i].Obj.'+FileID);
				else  SelectedFile=SelectedFile+','+eval('DocElementArr[i].Obj.'+FileID);
			}
			if (eval('DocElementArr[i].Obj.'+FolderID)!=null)
			{
				if (SelectedFolder=='') SelectedFolder=eval('DocElementArr[i].Obj.'+FolderID);
				else  SelectedFolder=SelectedFolder+','+eval('DocElementArr[i].Obj.'+FolderID);
			}
		}
	}
}
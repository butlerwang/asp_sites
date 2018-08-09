
 //空间更多
 function SpacePage(curPage,action)
 {
   this._username = null;
   this._action   = action;
   this._c_obj    = "spacemain";
   this._p_obj    = "kspage";
   this._page     = curPage;
   this._url      = "ajax.asp";
   loadDate(1);
 }


//当前页,动作，用户名
function Page(curPage,action,username)
   {
   this._username = username;
   this._action   = action;
   this._c_obj    = ksblog._mainlist;
   this._p_obj    = ksblog._pagelist;
   this._page     = curPage;
   this._url      = ksblog._url;
   loadDate(1);
   }

 //圈子主题
 function TeamPage(curPage,action)
 {
   this._username = null;
   this._action   = action;
   this._c_obj    = "teammain";
   this._p_obj    = "kspage";
   this._page     = curPage;
   this._url      = "groupajax.asp";
   loadDate(1);
 }
function loadDate(p)
{  this._page=p;
   var xhr=new ksblog.Ajax();
   xhr.open("get",_url+"?action="+_action+"&username="+escape(_username)+"&page="+p,true);
   xhr.onreadystatechange=function (){
	         if(xhr.readyState==1)
			  {
				document.getElementById(_c_obj).innerHTML="<div align='center'><img src='images/loading.gif'>正在加载...</div>";
			  }
			  else if(xhr.readyState==2 || xhr.readyState==3)
			  {
				  if (p==1)
				 document.getElementById(_c_obj).innerHTML="<div align='center'><img src='images/loading.gif'>正在读取数据...</div>";
				}
			  else if(xhr.readyState==4)
			  {
			 if (xhr.status==200)
			 {   
				  var pagearr=xhr.responseText.split("{ks:page}")
				  var pageparamarr=pagearr[1].split("|");
				  count=pageparamarr[0];    
				  perpagenum=pageparamarr[1];
				  pagecount=pageparamarr[2];
				  itemunit=pageparamarr[3];   
				  itemname=pageparamarr[4];
				  pagestyle=pageparamarr[5];
				  document.getElementById(_c_obj).innerHTML=pagearr[0];
				  pagelist();
			 }
			}
	   }
    xhr.send(null);  	
}

function pagelist()
{
 var n=1;	
 var statushtml=null;
 switch(parseInt(this.pagestyle))
 {
  case 1:	
     statushtml="共"+this.count+this.itemunit+" <a href=\"javascript:homePage();\" title=\"首页\">首页</a> <a href=\"javascript:previousPage()\" title=\"上一页\">上一页</a>&nbsp;<a href=\"javascript:nextPage()\" title=\"下一页\">下一页</a> <a href=\"javascript:lastPage();\" title=\"最后一页\">尾页</a> 页次:<font color=red>"+this._page+"</font>/"+this.pagecount+"页 "+this.perpagenum+this.itemunit+this.itemname+"/页";
		break;
  case 2:
	 statushtml="共"+this.pagecount+"页/"+this.count+this.itemunit+this.itemname+" <a href=\"javascript:homePage();\" title=\"首页\"><font face=webdings>9</font></a> <a href=\"javascript:previousPage()\" title=\"上一页\"><font face=webdings>7</font></a>&nbsp;";
	 var startpage=1;
	 if (this._page>10)
	   startpage=(parseInt(this._page/10)-1)*10+parseInt(this._page%10)+1;
	  for(var i=startpage;i<=this.pagecount;i++){ 
		  if (i==this._page)
		   statushtml+="<a href=\"javascript:turn("+i+")\"><font color=\"#ff0000\">"+i+"</font></a>&nbsp;"
		  else
			statushtml+="<a href=\"javascript:turn("+i+")\">"+i+"</a>&nbsp;"
			n=n+1;
		  if (n>10) break;
	  }
	 statushtml+="<a href=\"javascript:nextPage()\" title=\"下一页\"><font face=webdings>8</font></a> <a href=\"javascript:lastPage();\" title=\"最后一页\"><font face=webdings>:</font></a>";
	break;	 
  case 3:
     statushtml="第<font color=#ff000>"+this._page+"</font>页 共"+this.pagecount+"页 <a href=\"javascript:homePage();\" title=\"首页\"><<</a> <a href=\"javascript:previousPage()\" title=\"上一页\"><</a>&nbsp;<a href=\"javascript:nextPage()\" title=\"下一页\">></a> <a href=\"javascript:lastPage();\" title=\"最后一页\">>></a> "+this.perpagenum+this.itemunit+this.itemname+"/页";
  case 4:
     statushtml="页次:<font color=red>"+this._page+"</font>/"+this.pagecount+"页 [ <a href=\"javascript:homePage();\" title=\"首页\">首页</a> <a href=\"javascript:previousPage()\" title=\"上一页\">上一页</a>&nbsp;<a href=\"javascript:nextPage()\" title=\"下一页\">下一页</a> <a href=\"javascript:lastPage();\" title=\"最后一页\">尾页</a> ]";
   break;
 }
	 statushtml+="&nbsp;跳转到第<select name=\"goto\" onchange=\"turn(parseInt(this.value));\">";
	  for(var i=1;i<=this.pagecount;i++){
		 if (i==this._page)
		 statushtml+="<option value='"+i+"' selected>"+i+"</option>";
		 else
		 statushtml+="<option value='"+i+"'>"+i+"</option>";
	  }	
	 statushtml+="</select>页";
	// if (this.pagecount!="")
	// {
	 document.getElementById(this._p_obj).innerHTML=statushtml;
	// }
}
function homePage()
{
   if(this._page==1)
    alert("已经是首页了！")
   else
   loadDate(1);
} 
function lastPage()
{
   if(this._page==this.pagecount)
    alert("已经是最后一页了！")
   else
   loadDate(this.pagecount);
} 
function previousPage()
{
   if (this._page>1)
      loadDate(this._page-1);
   else
      alert("已经是第一页了");      
}

function nextPage()
{
   if(this._page<this.pagecount)
      loadDate(this._page+1);
   else
      alert("已经到最后一页了");
}
function turn(i)
{
      loadDate(i);
}
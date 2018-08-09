
//ajax 控件
function PageAjax(){
	if(window.XMLHttpRequest){
		return new XMLHttpRequest();
	} else if(window.ActiveXObject){
		return new ActiveXObject("Microsoft.XMLHTTP");
	} 
	throw new Error("XMLHttp object could be created.");
}

function Page(curPage,labelid,classid,installdir,url,refreshtype,specialid,infoid)
   {
   this.labelid=labelid;
   this.classid=classid;
   this.url=url;
   if (labelid.substring(0,5)=="{SQL_")
   {
	var slabelid=labelid.split('(')[0];
    slabelid=slabelid.replace("{","");
    this.c_obj="c_"+slabelid;
    this.p_obj="p_"+slabelid;
   }
   else
   {
   this.c_obj="c_"+labelid;
   this.p_obj="p_"+labelid;
   }
   this.installdir=installdir;
   this.refreshtype=refreshtype;
   this.specialid=specialid;
   this.infoid=infoid;
   this.page=curPage;
   loadData(1);
   }
function loadData(p)
{  this.page=p;
   var xhr=new PageAjax();
   var senddata=installdir+url+"?labelid="+escape(labelid)+"&infoid="+infoid+"&classid="+classid+"&refreshtype="+refreshtype+"&specialid=" +specialid+"&curpage="+p+getUrlParam();
   xhr.open("get",senddata,true);
   xhr.onreadystatechange=function (){
	         if(xhr.readyState==1)
			  {
				 if (p==1)
				eval('document.all.'+c_obj).innerHTML="<div align='center'><img src='"+installdir+"images/loading.gif'>正在连接服务器...</div>";
			  }
			  else if(xhr.readyState==2 || xhr.readyState==3)
			  {
				if (p==1)
				eval('document.all.'+c_obj).innerHTML="<div align='center'><img src='"+installdir+"images/loading.gif'>正在读取数据...</div>";
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
				  getObject(c_obj).innerHTML=pagearr[0];
				  pagelist();
			 }
			}
	   }
    xhr.send(null); 
}
//取url传的参数
function getUrlParam()
{
	var URLParams = new Object() ;
	var aParams = document.location.search.substr(1).split('&') ;//substr(n,m)截取字符从n到m,split('o')以o为标记,分割字符串为数组
	if(aParams!=''&&aParams!=null){
		var sum=new Array(aParams.length);//定义数组
		for (i=0 ; i < aParams.length ; i++) {
		sum[i]=new Array();
		var aParam = aParams[i].split('=') ;//以等号分割
		URLParams[aParam[0]] = aParam[1] ;
		sum[i][0]=aParam[0];
		sum[i][1]=aParam[1];
		}
		var p='';
		for(i=0;i<sum.length;i++)
		{
		  p=p+'&'+sum[i][0]+"="+sum[i][1]
		}
	   return p;
	}else{
	   return "";
	}
}
function getObject(id) 
{
	if(document.getElementById) 
	{
		return document.getElementById(id);
	}
	else if(document.all)
	{
		return document.all[id];
	}
	else if(document.layers)
	{
		return document.layers[id];
	}
}

function pagelist()
{
 var n=1;	
 var statushtml=null;
 switch(parseInt(this.pagestyle))
 {
  case 1:	
     statushtml="共"+this.count+this.itemunit+" <a href=\"javascript:homePage(1);\" title=\"首页\">首页</a> <a href=\"javascript:previousPage()\" title=\"上一页\">上一页</a>&nbsp;<a href=\"javascript:nextPage()\" title=\"下一页\">下一页</a> <a href=\"javascript:lastPage();\" title=\"最后一页\">尾页</a> 页次:<font color=red>"+this.page+"</font>/"+this.pagecount+"页 "+this.perpagenum+this.itemunit+this.itemname+"/页";
		break;
  case 2:
	 statushtml="<a href='#'>"+this.pagecount+"页/"+this.count+this.itemunit+"</a> <a href=\"javascript:homePage(1);\" title=\"首页\"><span style='font-family:webdings;font-size:14px'>9</span></a> <a href=\"javascript:previousPage()\" title=\"上一页\"><span style='font-family:webdings;font-size:14px'>7</span></a>&nbsp;";
	 var startpage=1;
	 if (this.page==10)
	   startpage=2;
	 else if(this.page>10)
	   startpage=eval((parseInt(this.page/10)-1)*10+parseInt((this.page)%10)+2);
	  for(var i=startpage;i<=this.pagecount;i++){ 
		  if (i==this.page)
		   statushtml+="<a href=\"#\"><font color=\"#ff0000\">"+i+"</font></a>&nbsp;"
		  else
			statushtml+="<a href=\"javascript:turn("+i+")\">"+i+"</a>&nbsp;"
			n=n+1;
		  if (n>10) break;
	  }
	 statushtml+="<a href=\"javascript:nextPage()\" title=\"下一页\"><font face=webdings>8</font></a> <a href=\"javascript:lastPage();\" title=\"最后一页\"><span style='font-family:webdings;font-size:14px'>:</span></a>";
	 statushtml="<span class='kspage'>"+statushtml+"</span>";
	break;	 
  case 4:
	 statushtml="<table border='0' align='right'><tr><td><a class='prev' href='javascript:previousPage();'>上一页</a>";
	 statushtml+="<a class='prev' href='javascript:nextPage();'>下一页</a>";
	 statushtml+="<a class='prev' href='javascript:homePage(1);'>首 页</a>";
	 var startpage=1;
	 if (this.page>7) startpage=page-5;
	 if (this.pagecout-this.page<5) startpage=this.pagecount-9;
	  for(var i=startpage;i<=this.pagecount;i++){ 
		  if (i==this.page)
		   statushtml+="<a href='javascript:void(0)' class='curr'><font color=\"#ff0000\">"+i+"</font></a>"
		  else
			statushtml+="<a class='num' href=\"javascript:turn("+i+")\">"+i+"</a>"
			n=n+1;
		  if (n>10) break;
	  }
	 statushtml+="<a href=\"javascript:lastPage();\" class='next' title=\"最后一页\">末 页</a><span>共有" +this.pagecount+"页</td></tr></table>";
	break;	 
  case 3:
     statushtml="第<font color=#ff000>"+this.page+"</font>页 共"+this.pagecount+"页 <a href=\"javascript:homePage(1);\" title=\"首页\"><<</a> <a href=\"javascript:previousPage()\" title=\"上一页\"><</a>&nbsp;<a href=\"javascript:nextPage()\" title=\"下一页\">></a> <a href=\"javascript:lastPage();\" title=\"最后一页\">>></a> "+this.perpagenum+this.itemunit+this.itemname+"/页";
   break;
 }
  if (parseInt(this.pagestyle)!=4){
	 statushtml+="&nbsp;第<select name=\"goto\" onchange=\"turn(parseInt(this.value));\">";
	  for(var i=1;i<=this.pagecount;i++){
		 if (i==this.page)
		 statushtml+="<option value='"+i+"' selected>"+i+"</option>";
		 else
		 statushtml+="<option value='"+i+"'>"+i+"</option>";
	  }	
	 statushtml+="</select>页";
  }
	 getObject(this.p_obj).innerHTML=statushtml;
}
function homePage()
{
   if(this.page==1)
    alert("已经是首页了！")
   else
   loadData(1);
} 
function lastPage()
{
   if(this.page==this.pagecount)
    alert("已经是最后一页了！")
   else
   loadData(this.pagecount);
} 
function previousPage()
{
   if (this.page>1)
      loadData(this.page-1);
   else
      alert("已经是第一页了");      
}

function nextPage()
{
   if(this.page<this.pagecount)
      loadData(this.page+1);
   else
      alert("已经到最后一页了");
}
function turn(i)
{
     loadData(i);
}
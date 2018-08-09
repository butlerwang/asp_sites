//ajax 控件
function $$(element) {
	var elements = new Array();
	if (arguments.length > 1) {
		for (var i = 0, elements = [], length = arguments.length; i < length; i++)
		elements.push($$(arguments[i]));
		return elements;
	}
	element = document.getElementById(element);
	return element;
}
function DiggAjax(){
	if(window.XMLHttpRequest){
		return new XMLHttpRequest();
	} else if(window.ActiveXObject){
		return new ActiveXObject("Microsoft.XMLHTTP");
	} 
	throw new Error("XMLHttp object could be created.");
}
var loader=new DiggAjax;
function ajaxLoadPage(url,request,method,fun)
{ 
	method=method.toUpperCase();
	if (method=='GET')
	{
		urls=url.split("?");
		if (urls[1]=='' || typeof urls[1]=='undefined')
		{
			url=urls[0]+"?"+request;
		}
		else
		{
			url=urls[0]+"?"+urls[1]+"&"+request;
		}
		
		request=null;
	}
	loader.open(method,url,true);
	if (method=="POST")
	{
		loader.setRequestHeader("Content-Type","application/x-www-form-urlencoded");
	}
	loader.onreadystatechange=function(){
	     eval(fun+'()');
	}
	loader.send(request);
 }
 function digbacks()
  {
  if (loader.readyState==4)
  {
	var s=loader.responseText;
	switch (s)
	{
	    case "err":
		 alert('digg error!');
		 break;
		case "over":
		 alert('您已投票过了！');
		 break;
		case "nologin":
		  alert('您还没有登录，不能推荐!');
		  break;
		default:
		   var sarr=s.split('|');
		   $$("s"+sarr[0]).innerHTML=sarr[1];
		   try{
			$$("c"+sarr[0]).innerHTML=sarr[2];
			var znum=document.getElementById("s"+sarr[0]).innerHTML;
			var cnum=document.getElementById("c"+sarr[0]).innerHTML
			var totalnum=parseInt(znum)+parseInt(cnum);
			$$("perz"+sarr[0]).innerHTML=((znum*100)/totalnum).toFixed(2)+'%';
			$$("perc"+sarr[0]).innerHTML=((cnum*100)/totalnum).toFixed(2)+'%';
			$$("digzcimg").style.width = parseInt((znum/totalnum)*55);
			$$("digcimg").style.width = parseInt((cnum/totalnum)*55);
		  }catch(e){
		  }

	}

	}
  }
  

//Digg
function digg(channelid,infoid,installdir){dig(channelid,infoid,installdir,0)}
function cai(channelid,infoid,installdir){dig(channelid,infoid,installdir,1)}
function dig(channelid,infoid,installdir,type)
{
 try{
  ajaxLoadPage(installdir+'plus/digg.asp','digtype='+type+'&action=hits&ChannelID='+channelid+'&infoid=' +infoid,'post','digbacks');
 }catch(e){
	 var head = document.getElementsByTagName("head")[0];        
	 var js = document.createElement("script"); 
	 js.src = installdir+'plus/digg.asp?printout=js&digtype='+type+'&action=hits&ChannelID='+channelid+'&infoid=' +infoid; 
	 head.appendChild(js);   
  }
}
function show_digg(channelid,infoid,installdir)
{ 
 var url=installdir+"plus/digg.asp?channelid="+channelid+"&infoid="+infoid+"&action=show";
 try
 {
   var xhr=new DiggAjax();
   xhr.open("get",url,true);
   xhr.onreadystatechange=function (){
	         if(xhr.readyState==1)
			  {
				$$("s"+infoid).innerHTML="<img src='"+installdir+"images/loading.gif'>";
			  }
			  else if(xhr.readyState==2 || xhr.readyState==3)
			  {
				$$("s"+infoid).innerHTML="<img src='"+installdir+"images/loading.gif'>";
			  }
			  else if(xhr.readyState==4)
			  {
			 if (xhr.status==200)
			 {   
				  var r=xhr.responseText
				  var rarr=r.split('|');
			      $$("s"+infoid).innerHTML=rarr[1];
				   try{
			        $$("c"+infoid).innerHTML=rarr[2];
					var znum=document.getElementById("s"+infoid).innerHTML;
					var cnum=document.getElementById("c"+infoid).innerHTML
					var totalnum=parseInt(znum)+parseInt(cnum);
					if (parseInt(znum)==0){
					$$("perz"+infoid).innerHTML='0%';
					}else{
					$$("perz"+infoid).innerHTML=((znum*100)/totalnum).toFixed(2)+'%';
					}
					if (parseInt(cnum)==0){
					$$("perc"+infoid).innerHTML='0%';
					}else{
					$$("perc"+infoid).innerHTML=((cnum*100)/totalnum).toFixed(2)+'%';
					}
					$$("digzcimg").style.width = parseInt((znum/totalnum)*55);
					$$("digcimg").style.width = parseInt((cnum/totalnum)*55);
				  }catch(e){
				  }

			 }
			}
	   }
    xhr.send(null);  
 }
 catch(e){
	 var head = document.getElementsByTagName("head")[0];        
	 var js = document.createElement("script"); 
	 js.src = url+"&printout=js"; 
	 head.appendChild(js);   
	}
}

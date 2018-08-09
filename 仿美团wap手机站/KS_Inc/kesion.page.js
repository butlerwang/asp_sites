
function pageinfo(pagestyle,perpagenum,FExt,LinkUrlFname)
{ 
 //得到当前页码
  var url=window.location.href;
  var urls=url.split("_");
  var page=urls[urls.length-1].split('.')[0];
  if (checkRate(page)==false) page=TotalPage;
  page=parseInt(TotalPage-page+1);

  if (page<0 || (typeof(urls[2])=="undefined") && urls[0].toLowerCase().indexOf("index")<0 && urls[0].toLowerCase().indexOf("default")<0 && isNaN(page)) page=1

  var HomeLink=LinkUrlFname+FExt;
  if (pagestyle==1){
  document.getElementById('totalrecord').innerHTML=TotalPut;
  document.getElementById('currpage').innerHTML=page;
  document.getElementById('totalpage').innerHTML=TotalPage;
  document.getElementById('perpagenum').innerHTML=perpagenum;
  }
  else if(pagestyle==2||pagestyle==3){
	  if (pagestyle==2){
			 var p_str=' ';
			 var startpage=1;
			 var n=1;
			 if (page>10)
			   startpage=(parseInt(page/10)-1)*10+parseInt(page%10)+1;
			  for(var i=startpage;i<=TotalPage;i++){ 
				  if (i==1){
					 if (i==page)
						p_str+=('<a href="' + HomeLink + '"><font color="#ff0000">['+i+']</font></a>&nbsp;');
					  else
						p_str+=('<a href="' + HomeLink + '">['+i+']</a>&nbsp;');
				  }
				  else{
					  if (i==page)
						p_str+=('<a href="' + LinkUrlFname +'_'+ eval(TotalPage-i+1) + FExt + '"><font color="#ff0000">['+i+']</font></a>&nbsp;');
					  else
						p_str+=('<a href="' + LinkUrlFname +'_'+ eval(TotalPage-i+1) + FExt + '">['+i+']</a>&nbsp;');
				  } 
					n=n+1;
				  if (n>10) break;
			  }
			 document.getElementById('pagelist').innerHTML=p_str;
	  }
		 document.getElementById('currpage').innerHTML=page;
		 document.getElementById('totalpage').innerHTML=TotalPage;
  }
  else if(pagestyle==4){
             var p_str=' ';
			 var startpage=1;
			 var n=1;
			  if (page>1)
				   if (page==2){
					p_str+= '<a href="'+ HomeLink +'" class="prev">上一页</a>';
				   }else{
					p_str+= '<a href="'+ LinkUrlFname + '_' +eval(TotalPage-page+2) + FExt +'" class="prev">上一页</a>';
				   }
			   
			  if (page!=TotalPage) p_str+='<a href="'+ LinkUrlFname +'_' + eval(TotalPage-page)+FExt+'" class="next">下一页</a>';
			  p_str+= '<a href="'+HomeLink+'" class="prev">首 页</a>';
			 
			 if (page>7) startpage=page-5;
			 if (TotalPage-page<5) startpage=TotalPage-9;
			  if (startpage<=0) startpage=1;
			  for(var i=startpage;i<=TotalPage;i++){ 
				  if (i==1){
					 if (i==page)
						p_str+=('<a href="' + HomeLink + '" class="curr"><font color="#ff0000">'+i+'</font></a>');
					  else
						p_str+=('<a href="' + HomeLink + '" class="num">'+i+'</a>');
				  }
				  else{
					  if (i==page)
						p_str+=('<a href="' + LinkUrlFname +'_'+ eval(TotalPage-i+1) + FExt + '" class="curr"><font color="#ff0000">'+i+'</font></a>');
					  else
						p_str+=('<a href="' + LinkUrlFname +'_'+ eval(TotalPage-i+1) + FExt + '" class="num">'+i+'</a>');
				  } 
					n=n+1;
				  if (n>10) break;
			  }
			  if (TotalPage==1){
			  p_str+='<a href="' + LinkUrlFname + FExt +'" class="prev">末页</a>';
			  }else{
			  p_str+='<a href="' + LinkUrlFname + '_1' + FExt +'" class="prev">末页</a>';
			  }
			  p_str+=' <span>总共'+ TotalPage +'页</span>';
			  
			 document.getElementById('pagelist').innerHTML=p_str; 
			 
			}
  
  
  if (pagestyle!=4){
	  for(var i=1;i<=TotalPage;i++)
	  { 
		 if (i==1)
		 document.getElementById('turnpage').options[i-1]=new Option('第'+i+'页',HomeLink);
		 else
		 document.getElementById('turnpage').options[i-1]=new Option('第'+i+'页',LinkUrlFname+'_'+eval(TotalPage-i+1)+FExt);
	  }
	  document.getElementById('turnpage').options(page-1).selected=true;
  }
}


function page(pagestyle,perpagenum,itemunit,FExt,LinkUrlFname)
{
 //得到当前页码
  var url=window.location.href;
  var urls=url.split("_");
  var page=urls[urls.length-1].split('.')[0];
  if (checkRate(page)==false) page=TotalPage;
  page=TotalPage-page+1;	
  if (page<0 || (typeof(urls[2])=="undefined") && urls[0].indexOf("index")<0 && urls[0].indexOf("default")<0) page=1
  var HomeLink=LinkUrlFname+FExt;
  document.write('<div style="text-align:right;">');
  switch (pagestyle)
   {
	  case 1:  //分页样式一
	    document.write('共 '+TotalPut+' '+itemunit+' 页次:<font color=red>' + page + '</font>/'+ TotalPage +' 页 '+  perpagenum + itemunit +'/页 ');
		if(page==1 && page!=TotalPage)
		 document.writeln('首页  上一页 <a href="'+LinkUrlFname+'_'+ eval(TotalPage -1)  + FExt +'">下一页</a>  <a href= "'+ LinkUrlFname+'_1'+ FExt +'">尾页</a>');
		else if(page==1 && page==TotalPage)
		 document.writeln('首页  上一页 下一页 尾页');
		else if (page==TotalPage && page != 2)
		 document.writeln('<a href="' + HomeLink +'">首页</a>  <a href="'+ LinkUrlFname + '_' + eval(TotalPage-page+2) + FExt +'">上一页</a> 下一页  尾页')
		else if(page == TotalPage && page == 2)
		 document.writeln('<a href="'+ HomeLink +'">首页</a>  <a href="'+ HomeLink +'">上一页</a> 下一页  尾页');
		else if(page == 2)
		 document.writeln('<a href="'+ HomeLink +'">首页</a>  <a href="'+ HomeLink +'">上一页</a> <a href="' + LinkUrlFname + '_'  +eval(TotalPage-page) + FExt +'">下一页</a>  <a href= "'+ LinkUrlFname +'_1'+ FExt +'">尾页</a>');
		else
		 document.writeln('<a href="'+ HomeLink +'">首页</a>  <a href="'+ LinkUrlFname +'_'+ eval(TotalPage-page+2) + FExt +'">上一页</a> <a href="'+ LinkUrlFname +'_'+eval(TotalPage -page)+ FExt +'">下一页</a>  <a href= "' + LinkUrlFname + '_1' + FExt + '">尾页</a>');
		 break;
	 case 2:   //分页样式二
	    document.writeln('第<font color=red>'+ page +'</font>页 共'+TotalPage +'页');
		if(page==1)
		 document.writeln('<span style="font-family:webdings;font-size:14px">9</span> <span style="font-family:webdings;font-size:14px">7</span>');
		else if(page==2)
		 document.writeln('<a href="'+HomeLink +'" title="首页"><span style="font-family:webdings;font-size:14px">9</span></a> <a href="'+ HomeLink +'" title="上一页"><span style="font-family:webdings;font-size:14px">7</span></a>');
		else
		 document.writeln('<a href="'+HomeLink +'" title="首页"><span style="font-family:webdings;font-size:14px">9</span></a> <a href="' + LinkUrlFname +'_'+ eval(TotalPage-page+2) + FExt + '" title=""上一页""><span style="font-family:webdings;font-size:14px">7</span></a> ');
		 
		 var startpage=1;
		 var n=1;
		 if (page>10)
		   startpage=(parseInt(page/10)-1)*10+parseInt(page%10)+1;
		  for(var i=startpage;i<=TotalPage;i++){ 
			  if (i==1){
				 if (i==page)
					document.write('<a href="' + HomeLink + '"><font color="#ff0000">['+i+']</font></a>&nbsp;');
				  else
					document.write('<a href="' + HomeLink + '">['+i+']</a>&nbsp;');
			  }
			  else{
				  if (i==page)
					document.write('<a href="' + LinkUrlFname +'_'+ eval(TotalPage-i+1) + FExt + '"><font color="#ff0000">['+i+']</font></a>&nbsp;');
				  else
					document.write('<a href="' + LinkUrlFname +'_'+ eval(TotalPage-i+1) + FExt + '">['+i+']</a>&nbsp;');
			  } 
				n=n+1;
			  if (n>10) break;
		  }
			 
			if (page==TotalPage)
			 document.writeln('<span style="font-family:webdings;font-size:14px">8</span> <span style="font-family:webdings;font-size:14px">:</span>');
			else
			 document.writeln('<a href="' + LinkUrlFname + '_' + eval(TotalPage -page) + FExt + '" title="下一页"><span style="font-family:webdings;font-size:14px">8</span></a> <a href="' + LinkUrlFname + '_1' + FExt + '"><span style="font-family:webdings;font-size:14px">:</span></a> ');
		   break;
	 case 3:   //分页样式三
	    document.writeln('第<font color=red>'+ page +'</font>页 共'+TotalPage +'页');
        if (page==1)
			 document.writeln('<span style="font-family:webdings;font-size:14px">9</span> <span style="font-family:webdings;font-size:14px">7</span>');
		else if (page==2)
		    document.writeln('<a href="' + HomeLink +'" title="首页"><span style="font-family:webdings;font-size:14px">9</span></a> <a href="' + HomeLink + '" title="上一页"><span style="font-family:webdings;font-size:14px">7</span></a>');
		else
			document.writeln('<a href="' + HomeLink + '" title="首页"><span style="font-family:webdings;font-size:14px">9</span></a> <a href="' + LinkUrlFname + '_' + eval(TotalPage-page+2) + FExt +'" title="上一页"><span style="font-family:webdings;font-size:14px">7</span></a> ');
		if (page==TotalPage)
		   document.writeln(' <span style="font-family:webdings;font-size:14px">8</span> <span style="font-family:webdings;font-size:14px">:</span>');
		else
		   document.writeln(' <a href="' + LinkUrlFname + '_'+eval(TotalPage -page) + FExt + '" title="下一页"><span style="font-family:webdings;font-size:14px">8</span></a> <a href="' + LinkUrlFname + '_1' + FExt + '"><span style="font-family:webdings;font-size:14px">:</span></a> ')
	  break;
   }
		 document.writeln(' 转到：<select name="page" size="1" onchange="javascript:window.location=this.options[this.selectedIndex].value;">');
		 for(var i=1;i<=TotalPage;i++)
		 {  var s="";
			if (page==i)
			 s=" selected";
		  if (i==1)
		   document.writeln('<option value="'+HomeLink+'"'+s+'>第'+i+'页</option>');
		  else
		  document.writeln('<option value="'+LinkUrlFname+'_'+eval(TotalPage-i+1)+FExt+'"'+s+'>第'+i+'页</option>');
		 }
		 document.writeln('</select>');
		 document.writeln('</div>');

}
function checkRate(value)
{
     var re = /^[0-9]+.?[0-9]*$/;   //判断字符串是否为数字     //判断正整数 /^[1-9]+[0-9]*]*$/   
     if (!re.test(value))
    {
        return false;
     }
	 return true;
}  
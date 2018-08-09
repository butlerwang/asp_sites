 //增加关注
 var ajaxdir='../user/';
 function addatt(userid,isreload){
   $.ajax({ 
    url:ajaxdir+'userajax.asp?userid='+userid+'&action=addAttention', 
    success:function(data){ 
	   if (data=='success'){
	    if (isreload=='true'){
			$.dialog.tips('恭喜，添加关注成功!',1,'success.gif',function(){this.reload();});
		}else{
        $.dialog.tips('恭喜，添加关注成功!',2,'success.gif',function(){});
		$("#attentionnum").html(parseInt($("#attentionnum").html())+1);
		$("#attentionnuml").html($("#attentionnum").html());
		}
	   }else{
	    $.dialog.tips(unescape(data),1,'error.gif',function(){});
	   }
    }, 
    cache:false 
 });
 }
 //取消关注
 function cancelatt(userid,id,isreload){
  $.dialog.confirm('您确定要取消关注吗？', function(){ 
	  $.ajax({ 
			url:ajaxdir+'userajax.asp?userid='+userid+'&action=cancelAttention', 
			success:function(data){ 
			  if (data=='success'){
			    if (isreload=='true'){
				 $.dialog.tips('恭喜，取消关注成功!',1,'success.gif',function(){this.reload();});
				}else{
					$("#attention"+id).hide('slow');
					$("#attentionnum").html(parseInt($("#attentionnum").html())-1);
					$("#attnum").html($("#attentionnum").html());
					$("#attentionnuml").html($("#attentionnum").html());
				}
			  }else{
			   $.dialog.tips(unescape(data),2,'error.gif',function(){});
			  }
			}, 
			cache:false 
		 });
  
		 
	}, function(){ 
	});
 }
 //跳出转播
 function trans(id){
   var box=$.dialog({title:'转播：',content: '<div id="transdiv">loading...</div>',max:false,min: false});
   $.ajax({ 
    url:ajaxdir+'userajax.asp?id='+id+'&action=TalkTrans', 
    success:function(data){ 
        box.content(data); 
    }, 
    cache:false 
 });
 }
 //开始转播
 function dotrans(id){
     var c=$("#transmsg").val();
	 $.post(ajaxdir+"UserAjax.asp",{transid:id,action:'TalkSave',qqweibo:$('#qqweibo').val(),sinaweibo:$("#sinaweibo").val(),Content:escape(c)},function(d){
		 if (d=="success"){
				 $.dialog.tips('恭喜，成功转播！',1,'success.gif',function(){location.href=ajaxdir+'weibo.asp';});
		  }else{
				$.dialog.tips(unescape(d),1,'error.gif',function(){});
			  }
	 });
 }
 //广播
 function checkmsg(){
	 var c=$("#msg").val();
	 if (c==''||c=='有什么新鲜事想告诉大家！'){
			 $.dialog.tips('请输入您想说的话哦 ^_^',2,'error.gif',function(){
			 $("#msg").focus();
	 });
		return false;
	 }
	 $.post(ajaxdir+"UserAjax.asp",{action:'TalkSave',qqweibo:$('#qqweibo').val(),sinaweibo:$("#sinaweibo").val(),Content:escape(c)},function(d){
		 if (d=="success"){
				 $.dialog.tips('恭喜，成功分享！',1,'success.gif',function(){location.href='weibo.asp';});
		  }else{
				$.dialog.tips(unescape(d),1,'error.gif',function(){});
			  }
	 });
 }
 //评论
 function quickreply(id,page){
     var c=$("#c"+id).val();
	 if (id=='') return;
	  $.ajax({ 
		url:ajaxdir+'userajax.asp?page='+page+'&id='+id+'&action=ShowTalkCmt', 
		success:function(data){ 
		  if (page!=undefined){
			$("#cmt"+id).html(data);
		  }else{
			$("#cmt"+id).toggle('fast').html(data);
		  }
		}, 
		cache:false 
	 });
 }

 //开始提交评论
 function dopostcmt(id){

     var c=$("#c"+id).val();
	 if (c==''||c=='我也说一句...'){
		$.dialog.tips('你也懒了点吧,要输入评论内容哦^-^',2,'error.gif',function(){
		$("#c"+id).focus();
	 });
		return false;
	 }
	 var addtomyweibo=0;
	 if ($("#addtomyweibo"+id).attr("checked")==true){
	 addtomyweibo=1;
	 }
	 $.post(ajaxdir+"UserAjax.asp",{action:'TalkCmtSave',id:id,addtomyweibo:addtomyweibo,Content:escape(c)},function(d){
		 if (d=="success"){
				 $.dialog.tips('恭喜，评论成功！',1,'success.gif',function(){location.href=ajaxdir+'weibo.asp';});
		  }else{
				$.dialog.tips(unescape(d),1,'error.gif',function(){});
			  }
	 });
 }
 //删除广播
 function delmsg(id){
  $.dialog.confirm('您确定要删除这条消息吗？', function(){ 
	  $.ajax({ 
			url:ajaxdir+'userajax.asp?id='+id+'&action=DelTalk', 
			success:function(data){ 
			  if (data=='success'){
				$("#w"+id).hide('slow');
				$("#msgnum").html(parseInt($("#msgnum").html())-1);
			  }else{
			   $.dialog.tips(unescape(data),2,'error.gif',function(){});
			  }
			}, 
			cache:false 
		 });
  
		 
	}, function(){ 
	});
 }
function checkcommentlength(cobj,cmax){ 
	if(cmax<=0) return;
	if (cobj.value.length>cmax) {
			cobj.value = cobj.value.substring(0,cmax);
			 $.dialog.alert("广播字数不能超过"+cmax+"个字符!",function(){});
	}else {
	   var remain=String(cmax-cobj.value.length);
	   var s='';
	   for (var i=0;i<remain.length;i++){
	    var n=remain.substr(i,1);
		if (n % 2==0){
	    s+='<font size="'+(i+2)+'" style="color:#666;padding-bottom:5px">'+n+'</font>';
		}else if(n%3==0){
	    s+='<font size="'+(i+3)+'" style="color:#999">'+n+'</font>';
		}else{
	    s+='<font size="'+(i+2)+'" style="color:#789;padding-top:5px">'+n+'</font>';
		}
	   }
		$('#commentmax').html(s);
	}
} 
 var box='';
 function ThisFocus(id){ if ($("#c"+id).val()=='我也说一句...'){  $("#c"+id).val(''); }}
 function ThisBlur(id){ if ($("#c"+id).val()==''){ $("#c"+id).val('我也说一句...'); }}

var b='';
function emot(){
	if (b){try{b.close()}catch(e){}};
	var str='<div class="emotlist"><ul>';
	for (var i=1;i<=24;i++){
		 if (i<10){NS="0"+i;}else{NS=i;}
	    str+='<li><a href="#" onclick="insertface(\'[em'+NS+']\');return false;"><Img src="../editor/ubb/images/smilies/default/'+NS+'.gif" /></a></li>';
	}
	str+="</ul></div>";
	b=$.dialog({title:false,content:str,left: $('.emotion').offset().left,top: $('.emotion').offset().top+20});
 }
 function image(){
	if (b){try{b.close()}catch(e){}};
	var str="<div>1、输入图片地址：<input type='text' style='width:240px;height:22px;line-height:22px;border:1px solid #ccc;' class='textbox' id='photourl'/> <input type='button' class='button' value='插入' onclick='insertImage()'/> <input type='button' value='取消' class='button' onclick='b.close();'/><br/><br/>2、上传图片：<span style='color:#999'>(支持Jpg、GIF及PNG格式图片)</span>&nbsp;";
	str+='<br/><iframe id="upiframe" name="upiframe" src="../user/BatchUploadForm.asp?ChannelID=9991" frameborder="0" width="480" height="60" scrolling="no"></iframe>';
	str+="<br/></div>"
	b=$.dialog({title:false,content:str,left: $('.image').offset().left,top: $('.image').offset().top+20});
}
function topic(){
	if (b){try{b.close()}catch(e){}};
	var str="<div>输入你要说的话<br/><strong>#<input type='text' style='width:240px;height:22px;line-height:22px;border:1px solid #ccc;' class='textbox' id='topicMsg'/>#</strong> <input type='button' class='button' value='插入' onclick='insertTopic()'/> <input type='button' value='取消' class='button' onclick='b.close();'/></div>"
	b=$.dialog({title:false,content:str,left: $('.topic').offset().left,top: $('.topic').offset().top+20});
}
function InsertFileFromUp(FileList,fileSize,maxId,title){
	var files=FileList.split('/');
	var file=files[files.length-1];
	var fileext = FileList.substring(FileList.lastIndexOf(".") + 1, FileList.length).toLowerCase();
	if (fileext=="gif" || fileext=="jpg" || fileext=="jpeg" || fileext=="bmp" || fileext=="png"){
		insertface('[img]'+FileList+'[/img]');	
	}else{	var str="["+"UploadFiles"+"]"+maxId+","+fileSize+","+fileext+","+title+"[/UploadFiles]";
		 insertface(str);	
	 }
}
 function video(){
	 if (b){try{b.close()}catch(e){}};
	 var str="<div>请输入视频地址：<span style='color:#999'>(支持优酷、土豆、酷六等)</span>";
	str+="<br/><input type='text' style='width:240px;height:22px;line-height:22px;border:1px solid #ccc;' class='textbox' id='videourl'/> <input type='button' class='button' value='插入' onclick='insertVideo()'/><input type='button' value='取消' class='button' onclick='b.close();'/>";
	 str+="</div>"
	b=$.dialog({title:false,content:str,left: $('.video').offset().left,top: $('.video').offset().top+20});
 }
 function insertVideo(){
	var v=$("#videourl").val();
	if (v==''){
		 $.dialog.tips('请输入视频地址!',1,'error.gif',function(){
		 $("#videourl").focus();
});
}else{
	var fileext='flv';
	if (v.lastIndexOf(".")!=-1){ fileext = v.substring(v.lastIndexOf(".") + 1, v.length).toLowerCase();}
		insertface('[media='+fileext+',400,300,0]'+v+'[/media]');
}
}
function insertImage(){
	  var v=$("#photourl").val();
	  if (v==''){
		$.dialog.tips('请输入图片地址!',1,'error.gif',function(){ $("#photourl").focus(); });
	  }else{insertface('[img]'+v+'[/img]');}
}
function insertTopic(){
	 var v=$("#topicMsg").val();
	 if (v==''){
	   $.dialog.tips('请输入您要说的话题!',1,'error.gif',function(){ $("#topicMsg").focus(); });
	 }else{ insertface('#'+v+'#');}
}
 function insertface(Val){ 
		 if (Val!=''){
		  var ubb=$("#msg")[0];
		  var ubbLength=$("#msg").val().length;
		  $("#msg").focus();
			if(typeof document.selection !=undefined){document.selection.createRange().text=Val;}
			else{ubb.value=ubb.value.substr(0,ubb.selectionStart)+Val+ubb.value.substring(ubb.selectionStart,ubbLength);}
		 }
		 b.close();
}
function showbigpic(p){
	var box=$.dialog({title:'查看原图：',content: '<div><img onload="if(980<this.offsetWidth)this.width=980;" style="max-width:980px" src="'+p+'"/></div>',max:false,min: false});
}
function setinputborder(t){
 if (t==0){
 $("#msg").attr("style","border:1px solid #999;color:#333");
 }else{
 $("#msg").attr("style","border:1px solid #e8e8e8;color:#b3b3b3");
 }
}
function openlink(url){
	var linkurl=url;
	if (url.substr(0,4).toLowerCase()!='http'){ url='http://'+url;}
	b=$.dialog({title:false,time:10,content:'<strong>您将要访问网页：</strong><br/>'+linkurl+'<br/>如果您不了解该网站的详细情况，请谨慎访问该页面，<br/>10秒后将自动关闭本提示。<br/><br/><input onclick=\'window.open("'+url+'");b.close();\' type=button class=btn value=\'访问\'/> <input onclick=b.close(); type=button class=btn value=\'关闭\'/>',icon:'alert.gif'});
}
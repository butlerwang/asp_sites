$(document).ready(function(){loadDocRating();});
function loadDocRating(){
		  $.getScript(url+"plus/DocRating.asp?id="+itemid+"&m_id="+channelid+"&c_id="+infoid+"&title="+escape(title),function(){
		    $("#DocRating").html(data.str); });
}
function PopRating(){
		   var str='<div id="popshow"><img src="'+url+'images/loading.gif"/>加载中...</div>';
		  	 new KesionPopup().popup("我要参与评分",str,400);
			  $.getScript(url+"plus/DocRating.asp?action=ShowPopup&id="+itemid+"&m_id="+channelid+"&c_id="+infoid,function(){
			   if (popu.islogin=='false'){
			    closeWindow();alert('请先登录!');
				new KesionPopup().popupIframe('会员登录',url+'user/userlogin.asp?Action=Poplogin',397,184,'no');
			   }else{
			   $("#popshow").html(popu.str);	
			   }
			});
}
function PostMyScore(){
		  var score=$("#myscore option:selected").val();
		  var myitem=$("input[name=myitem]:checked").val()
		  $.getScript(url+"plus/DocRating.asp?score="+score+"&itemid="+myitem+"&action=hits&id="+itemid+"&m_id="+channelid+"&c_id="+infoid+"&title="+escape(title),function(){
		     switch(vote.status){
				  case "nologin":
				   alert('对不起,您还没登录不能打分!');
				   break;
				  case "standoff":
				   alert('您已表态过了, 不能重复打分!');
				   break;
				  case "lock":
				   alert('打分已关闭!');
				   break;
				  case "errstartdate":
				   alert('未到打分时间!');
				   break;
				  case "errexpireddate":
				   alert('打分时间已过!');
				   break;
				  case "errgroupid":
				   alert('您所在的用户组没有打分的权限!');
				   break;
				  case "noinfo":
				   alert('找不到您要打分的信息!');
				   break;
				  default:
				   closeWindow();
				   alert('恭喜,您已成功打分!');
				   loadDocRating();
				   break;
				 }
		  });
 }
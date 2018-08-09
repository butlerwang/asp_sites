$(document).ready(function(){
 init_reg();
})

function getlicense()
{
  if ($("#viewlicense").is(":checked")) 
  {
    $("#license").show();
  }
  else
  {
   $("#license").hide();
  }
}
var loadverifycode=true;
function getCode()
{ 
 if (loadverifycode){
 $("#showVerify").html("<img style='cursor:pointer' src='../../plus/verifycode.asp?n='+Math.random() onClick='this.src=\"../../plus/verifycode.asp?n=\"+ Math.random();'  align='absmiddle'>");	
  loadverifycode=false;
 }
}
var msg	;
var bname_m=false;
var ajaxchk=null;
var ajaxstr=null;
function init_reg(){
	msg=new Array(
	"请输入"+minlen+"-"+maxlen+"位字符，英文、数字、下划线的组合。",
	"请输入4-14位字符，英文、数字的组合。",
	"请输入6位以上字符，不允许空格。",
	"请重复输入上面的密码。",
	"请选择密码提示问题。",
	"6个字符、数字或3个汉字以上（包括6个）。",
	"请输入您常用的电子邮箱地址。",
	"如果看不清，可以点击数字刷新验证码。",
	"请输入合法的手机号码。",
	"只有正确回答注册问题才可以继续。"
	)
	document.getElementById("usernamemsg").innerHTML=msg[0];
	document.getElementById("passwordmsg1").innerHTML=msg[2];
	document.getElementById("passwordmsg2").innerHTML=msg[3];
	document.getElementById("questionmsg").innerHTML=msg[4];
	document.getElementById("answermsg").innerHTML=msg[5];
	document.getElementById("emailmsg").innerHTML=msg[6];
	document.getElementById("chkcodemsg").innerHTML=msg[7];
	document.getElementById("mobilemsg").innerHTML=msg[8];
	document.getElementById("reganswermsg").innerHTML=msg[9];
}

function on_input(objname){
	var strtxt;
	var obj=document.getElementById(objname);
	obj.className="d_on";
	//alert(objname);
	switch (objname){
		case "usernamemsg":
			strtxt=msg[0];
			break;
		case "passwordmsg1":
			strtxt=msg[2];
			break;
		case "passwordmsg2":
			strtxt=msg[3];
			break;
		case "answermsg":
			strtxt=msg[5];
			break;
		case "emailmsg":
			strtxt=msg[6];
			break;
		case "chkcodemsg":
		    strtxt=msg[7];
			break;	
		case "mobilemsg":
		    strtxt=msg[8];
			break;
		case "reganswermsg":
		    strtxt=msg[9];
			break;
	}
	obj.innerHTML=strtxt;
}
function out_username(){
	var obj=document.getElementById("usernamemsg");
	var str=sl(document.getElementById("UserName").value);
	var chk=true;
	if (str<minlen || str>maxlen){chk=false;}
	if (!chk){
		obj.className="d_err";
		obj.innerHTML=msg[0];
		return;
	}
	$.ajax({type:"get",url:"regajax.asp?action=checkusername&username="+escape(document.getElementById("UserName").value)+"&anticache=" + Math.floor(Math.random()*1000),cache:false,dataType:"html",success:function(d){
	     var s=unescape(d);
		 ajaxchk=s.split('|')[0];
		 ajaxstr=s.split('|')[1];
	}
	 });
	if (ajaxstr!=null){
		if (ajaxchk=='ok'){
		  obj.className="d_ok";
		  obj.innerHTML=ajaxstr;
		 }else{
			obj.className="d_err";
			obj.innerHTML=ajaxstr;
		 }
	}
}
function out_password1(){
	var obj=document.getElementById("passwordmsg1");
	var str=document.getElementById("PassWord").value;
	var chk=true;
	if (str=='' || str.length<6 || str.length>14){chk=false;}
	if (chk){
		obj.className="d_ok";
		obj.innerHTML='密码已经输入。';
	}else{
		obj.className="d_err";
		obj.innerHTML=msg[2];
	}
	return chk;
}
function out_password2(){
	var obj=document.getElementById("passwordmsg2");
	var str=document.getElementById("RePassWord").value;
	var chk=true;
	if (str!=document.getElementById("PassWord").value||str==''){chk=false;}
	if (chk){
		obj.className="d_ok";
		obj.innerHTML='重复密码输入正确。';
	}else{
		obj.className="d_err";
		obj.innerHTML=msg[3];
	}
	return chk;
}
function out_question(){
	var obj=document.getElementById("questionmsg");
	var str=document.getElementById("Question").value;
	var chk=true;
	if (question==0) return true;
	if (str==''){chk=false}
	if (chk){
		obj.className="d_ok";
		obj.innerHTML='密码提示问题已经选择。';
	}else{
		obj.className="d_err";
		obj.innerHTML=msg[4];
	}
	return chk;
}
function out_answer(){
	var obj=document.getElementById("answermsg");
	var str=sl(document.getElementById("Answer").value);
	var chk=true;
	if (question==0) return true;
	if (str<6 || str>40){chk=false}
	if (chk){
		obj.className="d_ok";
		obj.innerHTML='密码提示问题答案已经输入。';
	}else{
		obj.className="d_err";
		obj.innerHTML=msg[5];
	}
	return chk;
}
function out_mobile(){
	var obj=document.getElementById("mobilemsg");
	var str=document.getElementById("Mobile").value;
	if (mobile==0) return true;
	var chk=ismobile(str);
	if (chk){
		$.ajax({type:"get",url:"regajax.asp?action=checkmobile&mobile="+str+"&anticache=" + Math.floor(Math.random()*1000),cache:false,dataType:"html",success:function(d){
	     var s=unescape(d);
		 ajaxchk=s.split('|')[0];
		 ajaxstr=s.split('|')[1];
	  }
	 });
		if (ajaxstr!=null){
		if (ajaxchk=='ok'){
		  obj.className="d_ok";
		  obj.innerHTML=ajaxstr;
		 }else{
			obj.className="d_err";
			obj.innerHTML=ajaxstr;
		 }
	}
	}else{
		obj.className="d_err";
		obj.innerHTML=msg[8];
	}
	return chk;	
}
function ismobile(s)
{
   var p = /^(013|015|13|15|018|18)\d{9}$/;
   if(s.match(p) != null){
  return true;
  }
  return false;
}
function out_email(){
	var obj=document.getElementById("emailmsg");
	var str=document.getElementById("Email").value;
	var chk=true;
	if (str==''|| !str.match(/^[\w\.\-]+@([\w\-]+\.)+[a-z]{2,4}$/ig)){chk=false}
	if (chk){
		obj.className="d_ok";
		obj.innerHTML='电子邮箱地址已经输入。';
	}else{
		obj.className="d_err";
		obj.innerHTML=msg[6];
		return chk;
	}
	$.get("regajax.asp",{action:"checkemail",email:escape(str)},function(d){
	     var s=unescape(d);
		 ajaxchk=s.split('|')[0];
		 ajaxstr=s.split('|')[1];
		if (ajaxstr!=null){
		if (ajaxchk=='ok'){
		  obj.className="d_ok";
		  obj.innerHTML=ajaxstr;
		 }else{
			obj.className="d_err";
			obj.innerHTML=ajaxstr;
		 }
		}

	});
			

}

function out_chkcode()
{	var obj=document.getElementById("chkcodemsg");
	var str=sl(document.getElementById("Verifycode").value);
	var chk=true;
	if (str<4 || str>6){chk=false}
	if (chk){
		obj.className="d_ok";
		obj.innerHTML='验证码已经输入。';
	}else{
		obj.className="d_err";
		obj.innerHTML=msg[7];
	return chk;
	}
	$.get("regajax.asp",{action:"checkcode",code:escape(document.getElementById("Verifycode").value)},function(d){
	     var s=unescape(d);
		 ajaxchk=s.split('|')[0];
		 ajaxstr=s.split('|')[1];
	})
	if (ajaxstr!=null){
		if (ajaxchk=='ok'){
		  obj.className="d_ok";
		  obj.innerHTML=ajaxstr;
		 }else{
			obj.className="d_err";
			obj.innerHTML=ajaxstr;
		 }
	 }
}
function sl(st){
	sl1=st.length;
	strLen=0;
	for(i=0;i<sl1;i++){
		if(st.charCodeAt(i)>255) strLen+=2;
	 else strLen++;
	}
	return strLen;
}

	 
      function CheckForm() 
		{ 
		   
		
			if (document.myform.UserName.value =="")
			{
			 $.dialog.alert("请填写您的会员名！",function(){document.myform.UserName.focus();});
			return false;
			}
			//var filter=/^\s*[.A-Za-z0-9_-]{{$Show_UserNameLimitChar},{$Show_UserNameMaxChar}}\s*$/;
			//if (!filter.test(document.myform.UserName.value)) { 
			//alert("会员名填写不正确,请重新填写！可使用的字符为（A-Z a-z 0-9 _ - .)长度不小于{$Show_UserNameLimitChar}个字符，不超过{$Show_UserNameMaxChar}个字符，注意不要使用空格。"); 
			//document.myform.UserName.focus();
			//return false; 
			//} 
			if (document.myform.PassWord.value =="") 
			{
			 $.dialog.alert("请填写您的密码！",function(){	document.myform.PassWord.focus();});
			 return false; 
			}
			if(document.myform.RePassWord.value==""){
			 $.dialog.alert("请输入您的确认密码！",function(){document.myform.RePassWord.focus();});
			 return false;
			}
			var filter=/^\s*[.A-Za-z0-9_-]{6,15}\s*$/;
			if (!filter.test(document.myform.PassWord.value)) { 
			 $.dialog.alert("密码填写不正确,请重新填写！可使用的字符为（A-Z a-z 0-9 _ - .)长度不小于6个字符，不超过15个字符，注意不要使用空格。",function(){document.myform.PassWord.focus();});
			 return false; 
			} 
			if (document.myform.PassWord.value!=document.myform.RePassWord.value ){
			  $.dialog.alert("两次填写的密码不一致，请重新填写！",function(){document.myform.PassWord.focus();});
			return false; 
			} 
			if (document.myform.Question.value ==""&&question==1)
			{
			  $.dialog.alert("请填写您的密码问题！",function(){document.myform.Question.focus();});
			  return false;
			}
			if (document.myform.Answer.value ==""&&question==1)
			{
			  $.dialog.alert("请填写您的问题答案！",function(){document.myform.Answer.focus();});
			  return false;
			}
			if (document.myform.Mobile.value ==""&&mobile==1)
			{
			  $.dialog.alert("请填写您的手机号码！",function(){document.myform.Mobile.focus();});
			  return false;
			}
			else if(ismobile(document.myform.Mobile.value)==false&&mobile==1)
			{
			  $.dialog.alert("您的手机号码不正确！",function(){document.myform.Mobile.focus();});
			  return false;
			}
			
			if (document.myform.Email.value =="")
			{
			  $.dialog.alert("请输入您的电子邮件地址！",function(){document.myform.Email.focus();});
			  return false;
			}
			if((document.myform.Email.value.indexOf("@")==-1)||(document.myform.Email.value.indexOf(".")==-1))
			{
			   $.dialog.alert("您输入的电子邮件地址有误！",function(){document.myform.Email.focus();});
				return false;
				}
		  
		  if ($("#viewlicense").attr("checked")!=true){
			  $.dialog.alert("只有阅读并完全接受会员服务条款才可以继续注册!",function(){});
			  return false;
			}
		  return true;

		}

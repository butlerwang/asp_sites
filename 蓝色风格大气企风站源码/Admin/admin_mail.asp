<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" href="css/common.css" type="text/css" />
<title>管理区域</title>
<style type="text/css">
	#mail td.title{ vertical-align:top; height:30px; line-height:30px; text-align:right; width:100px;}
	#mail td.ipt{ text-align:left; width:300px;height:20px;}
	#mail td.msg{ color:#F00; font-size:16px; width:150px; padding:30px;}
	#mail input{ width:100%; height:20px;}
	#mail textarea{ width:500px;}
	#mail .btn{ width:auto; height:auto;}
</style>
<script type="text/javascript">
	function chkMail(){
		var subject_ali = document.getElementById("subject_ali");
		var subject_user = document.getElementById("subject_user");
		var subject_qq = document.getElementById("subject_qq");
		var subject_email = document.getElementById("subject_email");
		var subject_phone = document.getElementById("subject_phone");
		var textbody = document.getElementById("textbody");
		if(subject_ali.getAttribute("value")==""&&subject_user.getAttribute("value")==""&&subject_qq.getAttribute("value")==""&&subject_email.getAttribute("value")==""&&subject_phone.getAttribute("value")==""){
			alert("旺旺号,会员号,QQ号,邮箱,手机号至少填写一项!");
			return false;
		}else{
			if(textbody.getAttribute("value")==""){
				alert("反馈信息不能为空!");
				return false;
			}
		}
		return true;
	}
</script>
</head>
<body>
<div id="man_zone">
<div align="center">
<iframe id="advs" src="http://show.Streakyhorse.com/services/biz/028/adv.asp" frameborder="0" scrolling="no" width="96%" height="30"></iframe></div>
<form id="mail" name="mail" method="post" action="mail.asp" onSubmit="return chkMail();">
    <input type="hidden" name="action" value="sendmail" />
    <input type="hidden" name="codeId" value="biz_a_028"/>
  <table width="95%" border="0" align="center"  cellpadding="3" cellspacing="1" class="table_style">
     <tr>
      <td colspan="3"  >&nbsp;慧谷动力软件信息反馈</td>
    </tr> 
    <tr>
      <td width="18%" class="left_title_1"><span class="left-title">旺旺号:</span></td>
      <td width="40%"><input id="subject_ali" name="subject_ali" /></td>
      <td class="msg" rowspan="5">旺旺号,会员号,QQ号,邮箱,手机号至少填写一项，以便于方便我们联系您!</td>
    </tr>
    <tr>
      <td class="left_title_2">会员号:</td>
      <td><input id="subject_user" name="subject_user" /></td>
    </tr>
    <tr>
      <td class="left_title_1">&nbsp;Q&nbsp;Q号:</td>
      <td><input id="subject_qq" name="subject_qq" /></td>
    </tr>
    <tr>
      <td class="left_title_2">邮&nbsp;&nbsp;箱:</td>
      <td><input id="subject_email" name="subject_email" /></td>
    </tr>
    <tr>
      <td class="left_title_1">手机号: </td>
      <td><input id="subject_phone" name="subject_phone" /></td>
    </tr>
    <tr>
      <td class="left_title_2">反馈信息:</td>
      <td colspan="2"><textarea id="textbody" rows="10" cols="90" name="textbody"></textarea></td>
    </tr>
    <tr>
      <td colspan="3"><div align="center">
	  <input class="btn" type="submit" value="发送邮件" />
	  </div></td>
    </tr>
  </table>
</form>
<div align="center">
<iframe id="help" src="http://show.Streakyhorse.com/services/notice.html" scrolling="no" width="96%" height="240" frameborder="1"></iframe>
</div>
</div>
</body>
</html>

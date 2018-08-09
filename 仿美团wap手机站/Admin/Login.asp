<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../Plus/md5.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<%
Const ShowVerifyCode= False    '后台登录是否启用验证码 true 启用 false不启用
Dim KS:Set KS=New PublicCls
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=KS.Setting(0) & "---网站后台管理"%></title>
<script type="text/JavaScript" src="Include/SoftKeyBoard.js"></script>
<script type="text/JavaScript" src="../ks_inc/jquery.js"></script>
<script type="text/JavaScript" src="../ks_inc/common.js"></script>
<script type="text/javascript" src="../ks_Inc/lhgdialog.js"></script>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
<style type="text/css">
	html{color:#000;font-family:Arial,sans-serif;font-size:12px;}
	h1, h2, h3, h4, h5, h6, h7, p, ul, ol,div,span, dl, dt, dd, li, body,em,i, form, input,i,cite, button, img, cite, strong,    em,label,fieldset,pre,code,blockquote, table, td, th ,tr{ padding:0; margin:0;outline:0 none;}
	img, table, td, th ,tr { border:0;}
	address,caption,cite,code,dfn,em,th,var{font-style:normal;font-weight:normal;}
	select,img,select{font-size:12px;vertical-align:middle;color:#666; font-family:Arial,sans-serif}
	.checkbox{vertical-align:middle;margin-right:5px;margin-top:-2px; margin-bottom:1px;}
	textarea{font-size:12px;color:#666; font-family:Arial,sans-serif}
	table{border-collapse:collapse;border-spacing:0;}
	ul, ol, li { list-style-type:none;}
	a { color:#0082cb; text-decoration:none;}
	a:hover{text-decoration:none;}
	ul:after,.clearfix:after { content: "."; display: block; height: 0; clear: both; visibility: hidden; }/* 不适合用clear时使用 */
	ul,.clearfix{ zoom:1;}
	.clear{clear:both;font-size:0px; line-height:0px;height:1px;overflow:hidden;}/*  空白占位  */
	body {margin:0 auto;font-size:12px; background:#E0F1FB;color:#666;position:relative}
	#wrap{margin-top:80px;}
	.main{width:800px;margin:0px auto;}
	.main_L{width:380px;float:right;background:url(images/linebg.png) no-repeat left center; padding-left:25px;margin-right:17px;display:inline;}
	.tabbox ul{margin-top:10px;}
	.tabbox li{padding:3px 0px 5px; position:relative;}
	.tabbox li.btn{padding-top:10px;padding-left:98px;}
	.tabbox .label{width:350px;height:38px;background:url(images/textbg.png) right -45px no-repeat;  } 
	.tabbox .label:hover{background:url(images/textbg.png) right 0px no-repeat;  } 
	.labelhover{width:350px;height:38px;background:url(images/textbg.png) right 0px no-repeat;  } 
	.tabbox label{font-size:14px;color:#666} 
	.tabbox .input,.tabbox .textinput{width:230px;height:26px;line-height:26px; padding:2px;padding-left:5px;border:0px;margin-top:3px;margin-left:10px;background-color:transparent; font-family:Verdana, Arial, Helvetica, sans-serif;}
	.tabbox .textinputhover,.tabbox .textinputhover{border:1px solid #aaa;}
	.regsubmit{width:182px;height:53px;border:0px none; background:url(images/reg_btn.jpg) 0px 0px no-repeat; cursor:pointer}
	.regsubmit:hover{background:url(images/reg_btn.jpg) 0px -53px no-repeat;}
	.main_R{width:330px;float:left;margin-left:38px;display:inline;}
	.tabbox .companyul{margin-top:20px}
	.rzm{margin-left:30px;line-height:25px;color:#999999}
	.rzm span{color:#CC0000;}
    .family{margin-top:80px; line-height:25px; font-size:14px; font-family:"微软雅黑";}
	.family h3{height:40px;text-align:center;line-height:40px;font-size:30px;font-weight:bold;color:#FD8504;}
    .family h3 span{font-size:30px;color:#666;}
	.foot{width:800px;margin:0px auto;text-align:center;padding:8px 0 0 0px;line-height:24px;}
	.foot a{color:##474747;}
	.foot a:visited{ color:#666;}
</style>
</head>
<body id="wrap">
<%
Select Case  KS.G("Action")
 Case "LoginCheck"
  Call CheckLogin()
 Case "LoginOut"
  Call LoginOut()
 Case Else
  Call CheckSetting()
  Call Main()
End Select

Sub CheckSetting()
     dim strDir,strAdminDir,InstallDir
	 strDir=Trim(request.ServerVariables("SCRIPT_NAME"))
	 strAdminDir=split(strDir,"/")(Ubound(split(strDir,"/"))-1) & "/"
	 InstallDir=left(strDir,instr(lcase(strDir),"/"&Lcase(strAdminDir)))
			
	If Instr(UCASE(InstallDir),"/W3SVC")<>0 Then
	   InstallDir=Left(InstallDir,Instr(InstallDir,"/W3SVC"))
	End If
 If KS.Setting(2)<>KS.GetAutoDoMain or KS.Setting(3)<>InstallDir Then
	
  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
  RS.Open "Select Setting From KS_Config",conn,1,3
  Dim SetArr,SetStr,I
  SetArr=Split(RS(0),"^%^")
  For I=0 To Ubound(SetArr)
   If I=0 Then 
    SetStr=SetArr(0)
   ElseIf I=2 Then
    SetStr=SetStr & "^%^" & KS.GetAutoDomain
   ElseIf I=3 Then
    SetStr=SetStr & "^%^" & InstallDir
   Else
    SetStr=SetStr & "^%^" & SetArr(I)
   End If
  Next
  RS(0)=SetStr
  RS.Update
  RS.Close:Set RS=Nothing
  Call KS.DelCahe(KS.SiteSn & "_Config")
  Call KS.DelCahe(KS.SiteSn & "_Date")
 End If
End Sub

Sub Main()
%>
<table width="809" border="0" height="418" align="center"  style="margin:0 auto;background:url(images/regbg.png);">
 <FORM ACTION="Login.asp?Action=LoginCheck" method="post" name="LoginForm" onSubmit="return(CheckForm(this))">
<tr>
 <td><div id="step_1" class="main">
				<div class="main_L">

					<div class="tabbox">
						<ul id="regSpan" class="companyul">
							<li style="z-index:1000">
								<div class="label">
									<label for="email" style="padding-left:28px">登录账号：</label><input type="text" name="UserName" id="UserName" class="textinput" tabindex="1" autocomplete="off" />
								</div>
							</li>
							<li>
								<div class="label">
									<label for="password" style="padding-left:28px">登录密码：</label><%IF KS.Setting(98)<>"1" Then%><input type="password" tabindex="2" name="PWD" id="PWD" class="textinput" /><%Else%><input name="PWD" type="password" onFocus="this.select();" onChange="Calc.password.value=this.value;" onClick="password1=this;showkeyboard();this.readOnly=1;Calc.password.value=''" onKeyDown="Calc.password.value=this.value;" maxlength="50" class="textinput" tabindex="2" readonly /><%End If%>

								</div>
								
							</li>
						  <%If ShowVerifyCode Then%>
							<li>
								<div class="label">
									<label for="Verifycode" style="padding-left:28px">验证字符：</label><input type="text" id="Verifycode" name="Verifycode" tabindex="3" class="textinput" maxlength="4" style="width:111px;" /><img id="imagecode" src="../plus/verifycode.asp?time=0.001" width="120" height="30" onclick="$(this).attr('src',$(this).attr('src')+Math.random());" title="点击刷新验证码" style="cursor:pointer;vertical-align:middle;*position:absolute;margin-top:-2px;*+margin-top:3px;_margin-top:3px"/>
								</div>
							</li>
						 <%End If%>	
						 <%if EnableSiteManageCode = True Then%>
							<li>
								<div class="label">
									<label for="password2" style="padding-left:28px">认证密码：</label><input type="password" id="AdminLoginCode" name="AdminLoginCode" tabindex="4" class="textinput" value="" />
								</div>
							</li>
							<%if SiteManageCode="8888" Then%>
							<li class="rzm">
								提示：原始认证密码为<span>8888</span>，为了安全请打开conn.asp文件修改
							</li>
							<%end if%>
						<%end if%>
							
							<li class="btn" id="nextStep">
							  <input type="submit" tabindex="5" class="regsubmit" value=" ">
							</li>
						</ul>
					</div>
				</div>
				<div class="main_R">
					<div class="family">
							<h3>CMS<span>管理系统</span></h3>
							 <br/>欢迎您选择智能CMS<sup>TM</sup>产品,我们一直在努力并提供能为您带来顶级体验的软件产品...
					</div>
				
				</div>
			</div>
 </td>
</tr>
</FORM>
</table>
<script type="text/javascript">
<!--
$(document).ready(function() { 
	$(".label").hover(function(){$(this).removeClass("label");$(this).addClass("labelhover");
	},function(){
	$(this).removeClass("labelhover");$(this).addClass("label");});
});

setTimeout(function(){$("#UserName").focus();},500); 

function CheckForm(ObjForm) {
  if(ObjForm.UserName.value == '') {
    $.dialog.alert('请输入管理账号！',function(){ObjForm.UserName.focus();});
    return false;
  }
  if(ObjForm.PWD.value == '') {
    $.dialog.alert('请输入授权密码！',function(){ObjForm.PWD.focus();});
    return false;
  }
  if (ObjForm.PWD.value.length<6)
  {
   $.dialog.alert('授权密码不能少于六位！',function(){ObjForm.PWD.focus();});
    return false;
  }
  <%If ShowVerifyCode Then%>
  if (ObjForm.Verifycode.value == '') {
    alert ('请输入验证字符！');
    ObjForm.Verifycode.focus();
    return false;
  }
  <%End If%>
  <%if EnableSiteManageCode = True Then%>
  if (ObjForm.AdminLoginCode.value == '') {
    $.dialog.alert('请输入后台管理认证密码！',function(){ObjForm.AdminLoginCode.focus();});
    return false;
  }
  <%End If%>
}
//-->
</script>


<br/><br/><br/>

</body>
</html>
<%End Sub
Sub CheckLogin()
  Dim PWD,UserName,LoginRS,SqlStr,RndPassword
  Dim ScriptName,AdminLoginCode
  AdminLoginCode=KS.G("AdminLoginCode")
  IF lcase(Trim(Request.Form("Verifycode")))<>lcase(Trim(Session("Verifycode"))) And ShowVerifyCode then 
   Call KS.Echo("<script>$.dialog.tips('<br/>登录失败:验证码有误，请重新输入！',1,'error.gif',function(){history.back();});</script>")
   exit Sub
  end if
  If EnableSiteManageCode = True And AdminLoginCode <> SiteManageCode Then
   Call KS.Echo("<script>$.dialog.tips('<br/>登录失败:您输入的后台管理认证码不对，请重新输入！',1,'error.gif',function(){history.back();});</script>")
   exit Sub
  End If
  Pwd =MD5(KS.R(KS.S("pwd")),16)

  UserName = KS.R(trim(KS.S("username")))
  RndPassword=KS.R(KS.MakeRandomChar(20))
  ScriptName=KS.R(Trim(Request.ServerVariables("HTTP_REFERER")))
  
  Set LoginRS = Server.CreateObject("ADODB.RecordSet")
  SqlStr = "select top 1 a.*,b.PowerList,b.ModelPower,B.[Type] from KS_Admin a inner join KS_UserGroup b on a.GroupID=b.ID where a.UserName='" & UserName & "'"
  LoginRS.Open SqlStr,Conn,1,3
  If LoginRS.EOF AND LoginRS.BOF Then
	  Call KS.InsertLog(UserName,0,ScriptName,"输入了错误的帐号!")
      Call KS.Die("<script>$.dialog.tips('<br/>登录失败:您输入了错误的帐号，请再次输入！',1,'error.gif',function(){history.back();});</script>")
  Else
  
     IF LoginRS("PassWord")=pwd THEN
       IF Cint(LoginRS("Locked"))=1 Then
         Call KS.Die("<script>$.dialog.tips('<br/>登录失败:您的账号已被管理员锁定，请与您的系统管理员联系！',1,'error.gif',function(){history.back();});</script>")
	   Else
		  	 '登录成功，进行前台验证，并更新数据
			  Dim UserRS:Set UserRS=Server.CreateObject("Adodb.Recordset")
			  UserRS.Open "Select top 1 * From KS_User Where UserName='" & LoginRS("PrUserName") & "' and GroupID=1",Conn,1,3
			  IF Not UserRS.Eof Then
			  
						If datediff("n",UserRS("LastLoginTime"),now)>=KS.Setting(36) then '判断时间
						UserRS("Score")=UserRS("Score")+KS.Setting(37)
						end if
					 UserRS("LastLoginIP") = KS.GetIP
					 UserRS("LastLoginTime") = Now()
					 UserRS("LoginTimes") = UserRS("LoginTimes") + 1
					 UserRS("RndPassWord") = RndPassWord
					 UserRS("IsOnline")=1
					 UserRS.Update		
	
					'置前台会员登录状态
                    If EnabledSubDomain Then
							Response.Cookies(KS.SiteSn).domain=RootDomain					
					Else
                            Response.Cookies(KS.SiteSn).path = "/"
					End If		
					 Response.Cookies(KS.SiteSn)("UserID") = UserRS("UserID")
					 Response.Cookies(KS.SiteSn)("UserName") = KS.R(UserRS("UserName"))
			         Response.Cookies(KS.SiteSn)("Password") = UserRS("Password")
					 Response.Cookies(KS.SiteSn)("RndPassword") = KS.R(UserRS("RndPassword"))
					 Response.Cookies(KS.SiteSn)("AdminLoginCode") = AdminLoginCode
					 Response.Cookies(KS.SiteSn)("AdminName") =  UserName
					 Response.Cookies(KS.SiteSn)("AdminPass") = pwd
					 If LoginRS("Type")=3 Then
					 Response.Cookies(KS.SiteSn)("SuperTF")   = 1
					 Else
					 Response.Cookies(KS.SiteSn)("SuperTF")   = 0
					 End If
					 Response.Cookies(KS.SiteSn)("GroupID") = LoginRS("GroupID")
					 Response.Cookies(KS.SiteSn)("PowerList") = LoginRS("PowerList")
					 Response.Cookies(KS.SiteSn)("ModelPower") = LoginRS("ModelPower")
					 'Response.Cookies(KS.SiteSn).Expires = DateAdd("h", 3, Now())   '3小时没有操作自动失败
             Else 
				   Call KS.InsertLog(UserName,0,ScriptName,"找不到前台账号!")
                   Call KS.Die("<script>$.dialog.tips('<br/>登录失败:找不到前台账号！',1,'error.gif',function(){history.back();});</script>")
			 End If
			   UserRS.Close:Set UserRS=Nothing
			   
	  LoginRS("LastLoginTime")=Now
	  LoginRS("LastLoginIP")=KS.GetIP
	  LoginRS("LoginTimes")=LoginRS("LoginTimes")+1
	  LoginRS.UpDate
	  Call KS.InsertLog(UserName,1,ScriptName,"成功登录后台系统!")
      Call KS.Die("<script>$.dialog.tips('<br/><span style=""font-size:14px;color:#888;font-weight:bold"">恭喜，成功登录<span style=""color:#ff6600"">[" & KS.Setting(0) & "]</span>网站后台系统！</span>',2,'success.gif',function(){location.href='index.asp';});</script>")
	End IF
  ELse
     If EnabledSubDomain Then
		Response.Cookies(KS.SiteSn).domain=RootDomain					
	 Else
        Response.Cookies(KS.SiteSn).path = "/"
	End If
    Response.Cookies(KS.SiteSn)("AdminName")=""
	Response.Cookies(KS.SiteSn)("AdminPass")=""
	Response.Cookies(KS.SiteSn)("SuperTF")=""
	Response.Cookies(KS.SiteSn)("AdminLoginCode")=""
	Response.Cookies(KS.SiteSn)("PowerList")=""
	Response.Cookies(KS.SiteSn)("ModelPower")=""
	Call KS.InsertLog(UserName,0,ScriptName,"输入了错误的口令:" & Request.form("pwd"))
    Call KS.Die("<script>$.dialog.tips('<br/>登录失败:您输入了错误的口令，请再次输入！',1,'error.gif',function(){history.back();});</script>")
  END IF
 End If
END Sub
Sub LoginOut()
		   Conn.Execute("Update KS_Admin Set LastLogoutTime=" & SqlNowString & " where UserName='" & KS.R(KS.C("AdminName")) &"'")
		   Dim AdminDir:AdminDir=KS.Setting(89)
		   If EnabledSubDomain Then
				Response.Cookies(KS.SiteSn).domain=RootDomain					
			Else
                Response.Cookies(KS.SiteSn).path = "/"
			End If
			Response.Cookies(KS.SiteSn)("PowerList")=""
			Response.Cookies(KS.SiteSn)("AdminName")=""
			Response.Cookies(KS.SiteSn)("AdminPass")=""
			Response.Cookies(KS.SiteSn)("SuperTF")=""
			Response.Cookies(KS.SiteSn)("AdminLoginCode")=""
			Response.Cookies(KS.SiteSn)("ModelPower")=""
			session.Abandon()
			Response.Write ("<script> top.location.href='" & KS.Setting(2) & KS.Setting(3) &"';</script>")
End Sub
Set KS=Nothing
%>

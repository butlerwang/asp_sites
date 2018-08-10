<!--#include file="../check.asp"-->
<!--#include file="../include/conn.asp"-->
<!--#include file="../include/lib.asp"-->
<%Dim Act
If Session("flag")<>99 then
Session.Abandon()
response.Write "<script LANGUAGE=javascript>alert('您的权限不足，请不要非法调用其它管理模块，否则您的账号将被系统自动删除! ');this.location.href='../login.asp';</script>"
end if
  Act=Request.Form("act")
  Select Case Act
     Case "save":Call SaveData()
	 Case else: Call Main()
  End Select
  Call CloseDataBase()  
  
  Sub SaveData()
    OpenData()
    Dim WebName,Company,UpfileType,UpfileSize,SmtpHost,SmtpUser,SmtpPwd,PicAuto,PicAutoType,PicPercent,PicHeight,PicWidth,Watermark,WatermarkSize,WatermarkWord,ShowProClass,ShowNewsClass,Template
    Dim Rs,Sql
	'=== 接受参数 ===
	WebName =       Trim(Request.Form("WebName"))
	Company =       Trim(Request.Form("Company"))
	UpfileType =    Trim(Request.Form("UpfileType"))
	UpfileSize =    Request.Form("UpfileSize")
	SmtpHost =      Trim(Request.Form("SmtpHost"))
	SmtpUser =      Trim(Request.Form("SmtpUser"))
	SmtpPwd =       Trim(Request.Form("SmtpPwd"))
	PicAuto =       Request.Form("PicAuto")
	PicAutoType =   Request.Form("PicAutoType")
	PicPercent =    Cint(Request.Form("PicPercent"))
	PicHeight =     Cint(Request.Form("PicHeight"))
	PicWidth =      Cint(Request.Form("PicWidth"))
	Watermark =     Request.Form("Watermark")
	WatermarkSize = Cint(Request.Form("WatermarkSize"))
	WatermarkWord = Trim(Request.Form("WatermarkWord"))
	ShowProClass =  Request.Form("ShowProClass")
	ShowNewsClass = Request.Form("ShowNewsClass")
	Template =      Trim(Request.Form("Template"))
	ShowNewsPic =   cint(Request.Form("ShowNewsPic"))
	ShowNewsAbout = cint(Request.Form("ShowNewsAbout"))	
	UpmovieType =   Trim(Request.Form("UpmovieType"))
	UpmovieSize =   Trim(Request.Form("UpmovieSize"))
	FtpUrl =   Trim(Request.Form("FtpUrl"))
	UserName =   Trim(Request.Form("UserName"))
	Password =   Trim(Request.Form("Password"))
	NewsClass_num =   Trim(Request.Form("NewsClass_num"))
	ProClass_num =   Trim(Request.Form("ProClass_num"))
	Pro_order =   Trim(Request.Form("Pro_order"))
	banben =   Trim(Request.Form("banben"))
	ShowqiyeClass =   Trim(Request.Form("ShowqiyeClass"))
	qiyeClass_num =   Trim(Request.Form("qiyeClass_num"))
	hy_message =   Trim(Request.Form("hy_message"))
	sf_yingpin =   Trim(Request.Form("sf_yingpin"))
	yanzhengma =   Trim(Request.Form("yanzhengma"))
	weblink_leibie =   Trim(Request.Form("weblink_leibie"))
	ShowdownClass =   Trim(Request.Form("ShowdownClass"))
	downClass_num =   Trim(Request.Form("downClass_num"))
	'=== 接收结束 ===
	
	'=== 验证参数 ===
	If WebName = "" or SmtpHost = "" Or SmtpUser = "" Or SmtpPwd = "" Or Template = "" Then
	   Response.Write("<script language=javascript>alert('请把表单填写完整!');window.history.back();</script>")
	   Response.End()
	End If
	'=== 验证结束 ===
	
	'=== 保存数据 ===
	Set Rs=Server.CreateObject("adodb.recordset")
    Sql="Select * From Sbe_WebConfig"
	Rs.Open Sql,Conn,1,3
	   Rs("WebName") =       WebName
	   Rs("NewsClass_num") = NewsClass_num
	   Rs("ProClass_num") =  ProClass_num  
	   Rs("Company") =       Company
	   Rs("UpfileType") =    UpfileType
	   Rs("UpfileSize") =    UpfileSize
	   Rs("UpmovieType") =    UpmovieType
	   Rs("UpmovieSize") =    UpmovieSize
	   Rs("FtpUrl")  =        FtpUrl
	   Rs("UserName") =       UserName
	   Rs("Password") =       Password   
	   Rs("SmtpHost") =      SmtpHost
	   Rs("SmtpUser") =      SmtpUser
	   Rs("SmtpPwd") =       SmtpPwd
	   Rs("PicAuto") =       PicAuto
	   Rs("PicAutoType") =   PicAutoType
	   Rs("PicPercent") =    PicPercent
	   Rs("PicHeight") =     PicHeight
	   Rs("PicWidth") =      PicWidth 
	   Rs("Watermark") =     Watermark
	   Rs("WatermarkSize") = WatermarkSize
	   Rs("WatermarkWord") = WatermarkWord
	   Rs("ShowProClass") =  ShowProClass
	   Rs("ShowNewsClass") = ShowNewsClass
	   Rs("ShowNewsPic")= ShowNewsPic
	   Rs("ShowNewsAbout")= ShowNewsAbout	   
	   Rs("Template") =      "0, "&Template
	   Rs("Pro_order")= Pro_order
	   Rs("banben")= banben
	   Rs("Pro_order")= Pro_order
	   Rs("qiyeClass_num")= qiyeClass_num
	   Rs("ShowqiyeClass")= ShowqiyeClass
	   Rs("hy_message")= hy_message
	   Rs("sf_yingpin")= sf_yingpin
	   Rs("yanzhengma")= yanzhengma 
	   Rs("weblink_leibie")= weblink_leibie
	   Rs("ShowdownClass")= ShowdownClass 
	   Rs("downClass_num")= downClass_num
	   Rs.Update
	 Rs.Close
	 Set Rs=Nothing
    '=== 保存结束 ===
	   Response.Write("<script language=javascript>alert('设置修改成功!');window.location.href='main.asp';</script>")
	   Response.End()
  End Sub
  
  Sub Main()
  OpenData()
  Dim Rs,Sql
  Set Rs=Server.CreateObject("adodb.recordset")
  Sql="Select * From Sbe_WebConfig"
  Rs.Open Sql,Conn,1,1
   If Rs.Eof Then
      Response.Write("配置信息已经被删除！")
	  Response.End()
   Else
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>后台管理系统</title>

<link href="../include/style.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="25"><font color="#6A859D">网站设置 &gt;&gt; 网站基本配置</font></td>
  </tr>
  <tr>
    <td height="1" background="../images/dot.gif"></td>
  </tr>
</table>
<form name="form1" method="post" action="">
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="0" id="sbe_table">
  <tr > 
    <td width="14%" height="25">网站名称</td>
    <td width="86%" height="21"><input name="WebName" type="text" class="input" id="WebName" value="<%=server.HTMLEncode(rs("WebName"))%>" size="50"></td>
  </tr>
  <tr> 
    <td height="25">公司名称</td>
    <td height="21"><input name="Company" type="text" class="input" id="Company" value="<%=server.HTMLEncode(rs("Company"))%>" size="50"></td>
  </tr>
  <tr> 
    <td height="25">上传文件类型</td>
    <td height="21"><input name="UpfileType" type="text" class="input" id="UpfileType" value="<%=rs("UpfileType")%>" size="40"></td>
  </tr>
  <tr> 
    <td height="25">上传文件大小</td>
    <td height="21"><input name="UpfileSize" type="text" class="input" id="UpfileSize" value="<%=rs("UpfileSize")%>" size="10">
        K</td>
  </tr>
  <tr>
    <td height="25">上传视频类型</td>
    <td height="21"><input name="UpmovieType" type="text" class="input" id="UpmovieType" value="<%=rs("UpmovieType")%>" size="40"></td>
  </tr>
  <tr>
    <td height="25">上传视频大小</td>
    <td height="21"><input name="UpmovieSize" type="text" class="input" id="UpmovieSize" value="<%=rs("UpmovieSize")%>" size="10">
      K</td>
  </tr>
</table>
<br>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="0" id="sbe_table">
  <tr > 
    <td width="14%" height="25">ftp地址</td>
    <td width="86%" height="21"><input name="FtpUrl" type="text" class="input" id="FtpUrl" value="<%=rs("FtpUrl")%>"> </td>
  </tr>
  <tr> 
    <td height="25">用户名</td>
    <td height="21"><input name="UserName" type="text" class="input" id="UserName" value="<%=rs("UserName")%>"></td>
  </tr>
  <tr> 
    <td height="25">密码</td>
    <td height="21"><input name="Password" type="text" class="input" id="Password" value="<%=rs("Password")%>"></td>
  </tr>
</table>
<br>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="0" id="sbe_table">
  <tr > 
    <td width="14%" height="25">Smtp服务器</td>
    <td width="86%" height="21"><input name="SmtpHost" type="text" class="input" id="SmtpHost" value="<%=rs("SmtpHost")%>"></td>
  </tr>
  <tr> 
    <td height="25">Smtp用户</td>
    <td height="21"><input name="SmtpUser" type="text" class="input" id="SmtpUser" value="<%=rs("SmtpUser")%>"></td>
  </tr>
  <tr> 
    <td height="25">Smtp密码</td>
    <td height="21"><input name="SmtpPwd" type="text" class="input" id="SmtpPwd" value="<%=rs("SmtpPwd")%>"></td>
  </tr>
</table>
<br>
<script language="JavaScript">
<!--
  function PicTypeShow(flag){
   if (flag==1){
       PicType1.style.display="";
	   PicType2.style.display="none";}
   if (flag==2){
       PicType1.style.display="none";
	   PicType2.style.display="";}
  }
  function WatermarkShow(flag){
   if (flag==1){
       Watermark1.style.display="";
	   Watermark2.style.display="";
	   }
   if (flag==2){
       Watermark1.style.display="none";
	   Watermark2.style.display="none";
	   }
  }
  function NewsClassShow(flag){
   if (flag==1){
       Show_NewsClass_num.style.display="";
	   }
   if (flag==0){
       Show_NewsClass_num.style.display="none";
	   }
  }
  function ProClassShow(flag){
   if (flag==1){
       Show_ProClass_num.style.display="";
	   }
   if (flag==0){
       Show_ProClass_num.style.display="none";
	   }
  }
  function qiyeClassShow(flag){
   if (flag==1){
       Show_qiyeClass_num.style.display="";
	   }
   if (flag==0){
       Show_qiyeClass_num.style.display="none";
	   }
	   }
   function downClassShow(flag){
   if (flag==1){
       Show_downClass_num.style.display="";
	   }
   if (flag==0){
       Show_downClass_num.style.display="none";
	   }
  }
//function ServerShow(){
//if (document.form1.Template[7].checked == true) {
//	show_server.style.display = "";
//  }else{
//	show_server.style.display = "none";
//}
//} 
-->
</script>
  <table width="98%" border="0" align="center" cellpadding="3" cellspacing="0" id="sbe_table">
    <tr > 
      <td height="25">是否自动生成小图</td>
      <td height="21"><input type="radio" name="PicAuto" value="1" <%Call ReturnSel(true,rs("PicAuto"),2)%>>
        是 
        <input type="radio" name="PicAuto" value="0" <%Call ReturnSel(false,rs("PicAuto"),2)%>>
        否</td>
    </tr>
    <tr > 
      <td height="25">小图比例</td>
      <td height="21"><input type="radio" name="PicAutoType" value="1" <%Call ReturnSel(1,rs("PicAutoType"),2)%> onClick="PicTypeShow(1)">
        百分比 
          <input type="radio" name="PicAutoType" value="2" <%Call ReturnSel(2,rs("PicAutoType"),2)%> onClick="PicTypeShow(2)">
        象素 </td>
    </tr>
    <tr id="PicType1" <%if rs("PicAutoType")=2 then%>style="display:none"<%end if%>>
      <td height="25">百分比</td>
      <td height="21"><input name="PicPercent" type="text" class="input" id="PicPercent" value="<%=rs("PicPercent")%>" size="2" maxlength="2">
        %</td>
    </tr>
    <tr id="PicType2" <%if rs("PicAutoType")=1 then%>style="display:none"<%end if%>> 
      <td height="25">象素</td>
      <td height="21">宽 <input name="PicHeight" type="text" class="input" id="PicHeight" value="<%=rs("PicHeight")%>" size="5">
        px &nbsp;&nbsp;&nbsp;&nbsp;高 <input name="PicWidth" type="text" class="input" id="PicWidth" value="<%=rs("PicWidth")%>" size="5">
        px</td>
    </tr>
    <tr > 
      <td width="14%" height="25">是否启用图片水印</td>
      <td width="86%" height="21"><input type="radio" name="Watermark" value="1" <%Call ReturnSel(true,rs("Watermark"),2)%> onClick="WatermarkShow(1)">
        是 
        <input type="radio" name="Watermark" value="0" <%Call ReturnSel(false,rs("Watermark"),2)%> onClick="WatermarkShow(2)">
        否</td>
    </tr>
    <tr id="Watermark1"  <%if rs("Watermark")=false then%>style="display:none"<%end if%>> 
      <td height="25">文字大小</td>
      <td height="21"><input name="WatermarkSize" type="text" class="input" id="WatermarkSize" value="<%=rs("WatermarkSize")%>" size="5"></td>
    </tr>
    <tr id="Watermark2" <%if rs("Watermark")=false then%>style="display:none"<%end if%>> 
      <td height="25">水印文字</td>
      <td height="21"><input name="WatermarkWord" type="text" class="input" id="WatermarkWord" value="<%=rs("Watermarkword")%>"></td>
    </tr>
  </table>
  <br>
  <br>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="0" id="sbe_table">
    <tr > 
      <td width="14%" height="25">系统类型:</td>
      <td width="86%" height="21"><input name="banben" type="radio" id="banben" value="1" <%Call ReturnSel(1,rs("banben"),2)%>>
        中文版&nbsp;&nbsp;
        <input name="banben" type="radio" id="banben" value="2" <%Call ReturnSel(2,rs("banben"),2)%>>
      中英文版</td>
    </tr>
    <tr > 
      <td width="14%" height="25">友情链接类型:</td>
      <td width="86%" height="21"><input name="weblink_leibie" type="radio" id="banben" value="0" <%Call ReturnSel(0,rs("weblink_leibie"),2)%>>
        文字&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name="weblink_leibie" type="radio" id="banben" value="1" <%Call ReturnSel(1,rs("weblink_leibie"),2)%>>
      图片&nbsp;&nbsp;
      &nbsp;
      &nbsp;
      <input name="weblink_leibie" type="radio" id="radio" value="2" <%Call ReturnSel(2,rs("weblink_leibie"),2)%>>
文字+图片</td>
    </tr>
  </table>
  <br>
  <table width="98%" border="0" align="center" cellpadding="3" cellspacing="0" id="sbe_table">
    <tr > 
      <td width="14%" height="25">是否显示资讯分类</td>
      <td width="86%" height="21"><input type="radio" name="ShowNewsClass" value="1" <%Call ReturnSel(true,rs("ShowNewsClass"),2)%> onClick="NewsClassShow(1)">
        显示
<input type="radio" name="ShowNewsClass" value="0" <%Call ReturnSel(false,rs("ShowNewsClass"),2)%> onClick="NewsClassShow(0)">
        隐藏 &nbsp;&nbsp;&nbsp;<span id="Show_NewsClass_num" <%if rs("ShowNewsClass")=false then response.Write("style='display:none'") end if%>>级数：
        <input name="NewsClass_num" type="text" class="input" style="ime-mode:Disabled;" onKeyPress="return event.keyCode>=48&&event.keyCode<=57||event.keyCode==45" value="<%=rs("NewsClass_num")%>" size="4" maxlength="4" onpaste="return !clipboardData.getData('text').match(/\D/)" ondragenter="return false">
      级&nbsp;必填(<font color="#FF0000">只能填大于0的数字</font>)</span></td>
    </tr>
    <tr> 
      <td height="25">是否显示产品分类</td>
      <td height="21"> 
        <input type="radio" name="ShowProClass" value="1" <%Call ReturnSel(true,rs("ShowProClass"),2)%> onClick="ProClassShow(1)">
        显示 
        <input type="radio" name="ShowProClass" value="0" <%Call ReturnSel(false,rs("ShowProClass"),2)%> onClick="ProClassShow(0)">
        隐藏  <span id="Show_ProClass_num" <%if rs("ShowProClass")=false then response.Write("style='display:none'") end if%>>&nbsp;&nbsp;&nbsp;级数：
        <input name="ProClass_num" type="text" class="input" style="ime-mode:Disabled;" onKeyPress="return event.keyCode>=48&&event.keyCode<=57||event.keyCode==45" value="<%=rs("ProClass_num")%>" size="4" maxlength="4" onpaste="return !clipboardData.getData('text').match(/\D/)" ondragenter="return false">
级&nbsp;必填(<font color="#FF0000">只能填大于0的数字</font>)</span></td>
    </tr>
<tr> 
      <td height="25">是否显示信息分类</td>
      <td height="21"> 
        <input type="radio" name="ShowqiyeClass" value="1" <%Call ReturnSel(true,rs("ShowqiyeClass"),2)%> onClick="qiyeClassShow(1)">
        显示 
        <input type="radio" name="ShowqiyeClass" value="0" <%Call ReturnSel(false,rs("ShowqiyeClass"),2)%> onClick="qiyeClassShow(0)">
        隐藏  <span id="Show_qiyeClass_num" <%if rs("ShowqiyeClass")=false then response.Write("style='display:none'") end if%>>&nbsp;&nbsp;&nbsp;级数：
        <input name="qiyeClass_num" type="text" class="input" style="ime-mode:Disabled;" onKeyPress="return event.keyCode>=48&&event.keyCode<=57||event.keyCode==45" value="<%=rs("qiyeClass_num")%>" size="4" maxlength="4" onpaste="return !clipboardData.getData('text').match(/\D/)" ondragenter="return false">
级&nbsp;必填(<font color="#FF0000">只能填大于0的数字</font>)</span></td>
    </tr>
<tr> 
      <td height="25">是否显示下载分类</td>
      <td height="21"> 
        <input type="radio" name="ShowdownClass" value="1" <%Call ReturnSel(true,rs("ShowdownClass"),2)%> onClick="downClassShow(1)">
        显示 
        <input type="radio" name="ShowdownClass" value="0" <%Call ReturnSel(false,rs("ShowdownClass"),2)%> onClick="downClassShow(0)">
        隐藏  <span id="Show_downClass_num" <%if rs("ShowdownClass")=false then response.Write("style='display:none'") end if%>>&nbsp;&nbsp;&nbsp;级数：
        <input name="downClass_num" type="text" class="input" style="ime-mode:Disabled;" onKeyPress="return event.keyCode>=48&&event.keyCode<=57||event.keyCode==45" value="<%=rs("downClass_num")%>" size="4" maxlength="4" onpaste="return !clipboardData.getData('text').match(/\D/)" ondragenter="return false">
级&nbsp;必填(<font color="#FF0000">只能填大于0的数字</font>)</span></td>
    </tr>
  </table>
  <br>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="0" id="sbe_table">
    <tr > 
      <td width="14%" height="25">产品是否提交订单</td>
      <td width="36%" height="21"><input type="radio" name="Pro_order" value="1" <%Call ReturnSel(true,rs("Pro_order"),2)%>>
        是 
        <input type="radio" name="Pro_order" value="0" <%Call ReturnSel(false,rs("Pro_order"),2)%>>
        否</td>
      <td width="17%">留言板是否有回复</td>
      <td width="33%">&nbsp;
        <input type="radio" name="hy_message" value="1" <%Call ReturnSel(true,rs("hy_message"),2)%>>
        是 
        <input type="radio" name="hy_message" value="0" <%Call ReturnSel(false,rs("hy_message"),2)%>>
      否</td>
    </tr>
<tr > 
      <td width="14%" height="25">在线招聘是否应聘</td>
      <td width="36%" height="21"><input name="sf_yingpin" type="radio" id="sf_yingpin" value="1" <%Call ReturnSel(true,rs("sf_yingpin"),2)%>>
        是
        <input name="sf_yingpin" type="radio" id="sf_yingpin" value="0" <%Call ReturnSel(false,rs("sf_yingpin"),2)%>>
        否</td>
      <td width="17%" height="25">后台登录是否有验证码</td>
      <td width="33%" height="21">&nbsp;
<input type="radio" name="yanzhengma" value="1" <%Call ReturnSel(true,rs("yanzhengma"),2)%>>
是
  <input type="radio" name="yanzhengma" value="0" <%Call ReturnSel(false,rs("yanzhengma"),2)%>>
否</td>
    </tr>
  </table>
  <br>

  <table width="98%" border="0" align="center" cellpadding="3" cellspacing="0" id="sbe_table">
    <tr > 
      <td width="14%" height="25">网站启用模块</td>
      <td width="86%" height="21">
	    <input name="Template" type="checkbox" id="Template" value="0" checked disabled>
        后台管理&nbsp; 
        <input name="Template" type="checkbox" id="Template" value="1" <%Call JudgeTemplate(rs("template"),", 1")%>>        
        企业信息 &nbsp; 
        <input name="Template" type="checkbox" id="Template" value="2" <%Call JudgeTemplate(rs("template"),", 2")%>>
        产品展示&nbsp; 
        <input name="Template" type="checkbox" id="Template" value="3" <%Call JudgeTemplate(rs("template"),", 3")%>>
        资讯中心&nbsp; 
        <input name="Template" type="checkbox" id="Template" value="4" <%Call JudgeTemplate(rs("template"),", 4")%>>
        下载管理 &nbsp; 
        <input name="Template" type="checkbox" id="Template" value="5" <%Call JudgeTemplate(rs("template"),", 5")%>>
        权限管理&nbsp;<br>
        <input name="Template" type="checkbox" id="Template" value="6" <%Call JudgeTemplate(rs("template"),", 6")%>>
        人事招聘&nbsp;
		<input name="Template" type="checkbox" id="Template" value="7" <%Call JudgeTemplate(rs("template"),", 7")%>>
        在线留言&nbsp; &nbsp;
        <input name="Template" type="checkbox" id="Template" value="8" <%Call JudgeTemplate(rs("template"),", 8")%>>
        订单管理&nbsp; 
        <input name="Template" type="checkbox" id="Template" value="9" <%Call JudgeTemplate(rs("template"),", 9")%>>
        友情连接 
       <!-- <input name="Template" type="checkbox" id="Template" value="10" <%'Call JudgeTemplate(rs("template"),", 10")%>>
        FAQ系统 --></td>
    </tr>
  </table>
  <p align="center">
    <input name="act" type="hidden" id="act" value="save">
    <input type="submit" name="Submit" value="提交修改" class="sbe_button">
    <br>
    <br>
  </p>
</form>
</body>
</html>
<%End If
  Rs.Close
  Set Rs=Nothing
  End Sub%>
<%
  Sub JudgeTemplate(Str,Str1)
     If Instr(Str,Str1)<>0 Then
	    response.Write("checked")
	 End If  
  End Sub
%>
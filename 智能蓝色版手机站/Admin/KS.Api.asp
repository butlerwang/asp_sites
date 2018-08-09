<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<!--#include file="../api/cls_api.asp"-->
<%

If Not KS.ReturnPowerResult(0, "KMST10002") Then          '检查是否有基本信息设置的权限
	Call KS.ReturnErr(1, "")
	Response.End
End If

Dim Action
Action = LCase(Request("action"))
Select Case Trim(Action)
	Case "save"
		Call SaveConformify
	Case Else
		Call showmain
End Select
Sub showmain()
Response.Write "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd""><html xmlns=""http://www.w3.org/1999/xhtml""><head><title>多系统整合接口设置</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" & vbCrLf
Response.Write "<link href='include/Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "<script src='../ks_inc/jquery.js'></script>" & vbCrLf
Response.Write "</head>"
Response.Write "<body bgcolor=""#FFFFFF"" topmargin=""0"" leftmargin=""0"">"
Response.Write "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
Response.Write "  <tr>"
Response.Write "    <td height=""25"" class=""topdashed"" valign='top' align='center'>"
Response.Write "      <b>API整合接口设置</b></td>"
Response.Write "  </tr>"
Response.Write "</TABLE>"
%>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1">
<form name="myform" method="post" action="?action=save">

<tr class="tdbg">
	<td height="30" width="25%" class="clefttitle" align="right"><strong>通用设置：</strong></td>
	<td>&nbsp;<b>首次登录自动创建账号并登录：</b>
	<label><input type="radio" name="API_QuickLogin"  value="false"<%
	If Not API_QuickLogin Then Response.Write " checked"
	%>> 不启用</label>&nbsp;&nbsp;
	<label><input type="radio" name="API_QuickLogin"  value="true"<%
	If API_QuickLogin Then Response.Write " checked"
	%>> 启用</label>
	<br/>
	&nbsp;<b>默认注册的会员用户组：</b>
	<%
	If KS.ChkClng(Api_GroupID)=0 Then Api_GroupID=2 '默认用户组
	Dim Node
	Call KS.LoadUserGroup()
	For Each Node In Application(KS.SiteSN&"_UserGroup").DocumentElement.SelectNodes("row[@showonreg=1 && @id!=1]")
	    if KS.ChkClng(Api_GroupID)=KS.ChkClng(Node.SelectSingleNode("@id").text) Then
		response.write "<label><input type=""radio""  value=""" & Node.SelectSingleNode("@id").text & """ name=""GroupID"" checked>" & Node.SelectSingleNode("@groupname").text  & "</label>"
		Else
		response.write "<label><input type=""radio""  value=""" & Node.SelectSingleNode("@id").text & """ name=""GroupID"">" & Node.SelectSingleNode("@groupname").text  & "</label>"
		End If
	Next
	%>
	</td>
</tr>


<tr class="tdbg">
	<td height="30" width="25%" class="clefttitle" align="right"><strong>是否开启QQ登录：</strong></td>
	<td>
	<label><input type="radio" name="API_QQEnable" onclick="$('#qq').hide()" value="false"<%
	If Not API_QQEnable Then Response.Write " checked"
	%>> 关闭</label>&nbsp;&nbsp;
	<label><input type="radio" name="API_QQEnable" onclick="$('#qq').show()" value="true"<%
	If API_QQEnable Then Response.Write " checked"
	%>> 开启</label>
	</td>
</tr>
<tbody id="qq"<%if cbool(API_QQEnable)=false then response.write " style='display:none'"%>>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>QQ登录AppID：</strong></td>
	<td><input type="text" class="textbox" name="API_QQAppId" size="35" value="<%=API_QQAppId%>"> 
		<font color="red">opensns.qq.com 申请到的appid,<a href="http://connect.qq.com/" target="_blank">点此申请</a>。</font>
	</td>
</tr>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>QQ登录AppKey：</strong></td>
	<td><input type="text" class="textbox" name="API_QQAppKey" size="35" value="<%=API_QQAppKey%>"> 
		<font color="red">opensns.qq.com 申请到的appkey。</font>
	</td>
</tr>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>QQ登录后跳转的地址：</strong></td>
	<td><input type="text" class="textbox" name="API_QQCallBack" style="background:#efefef" readonly size="45" value="<%=KS.GetDomain &"api/qq/callback.asp"%>"> 
		<font class="tips">QQ登录成功后跳转的地址,不可改。</font>
	</td>
</tr>
</tbody>

<tr class="tdbg"><td height="20" class="clefttitle"  colspan="2"></td></tr>

<tr class="tdbg">
	<td height="30" width="25%" class="clefttitle" align="right"><strong>是否开启新浪微博登录：</strong></td>
	<td>
	<label><input type="radio" name="API_SinaEnable" onclick="$('#sina').hide()" value="false"<%
	If Not API_SinaEnable Then Response.Write " checked"
	%>> 关闭</label>&nbsp;&nbsp;
	<label><input type="radio" name="API_SinaEnable" onclick="$('#sina').show()" value="true"<%
	If API_SinaEnable Then Response.Write " checked"
	%>> 开启</label>
	</td>
</tr>
<tbody id="sina"<%if cbool(API_SinaEnable)=false then response.write " style='display:none'"%>>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>新浪微博登录App Key：</strong></td>
	<td><input type="text" class="textbox" name="API_SinaId" size="35" value="<%=API_SinaId%>"> 
		<font color="red">新浪微博登录API申请网址：http://open.weibo.com/<a href="http://open.weibo.com/" target="_blank">点此申请</a>。</font>
	</td>
</tr>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>新浪微博登录App Secret：</strong></td>
	<td><input type="text" class="textbox" name="API_SinaKey" size="35" value="<%=API_SinaKey%>"> 
		<font color="red">新浪微博登录申请到的App Secret</font>
	</td>
</tr>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>新浪微博登录后跳转的地址：</strong></td>
	<td><input type="text" class="textbox" name="api_sinacallback" id="api_sinacallback" style="background:#efefef" readonly size="45" value="<%=KS.GetDomain &"api/sina/callback.asp"%>"> <font class="tips">新浪微博登录成功后跳转的地址,不可改。</font>
	</td>
</tr>

</tbody>
<tr class="tdbg"><td height="20" class="clefttitle"  colspan="2"></td></tr>

<tr class="tdbg">
	<td height="30" width="25%" class="clefttitle" align="right"><strong>是否开启支付宝快捷登录：</strong></td>
	<td>
	<label><input type="radio" name="API_AlipayEnable" onclick="$('#alipay').hide()" value="false"<%
	If Not API_AlipayEnable Then Response.Write " checked"
	%>> 关闭</label>&nbsp;&nbsp;
	<label><input type="radio" name="API_AlipayEnable" onclick="$('#alipay').show()" value="true"<%
	If API_AlipayEnable Then Response.Write " checked"
	%>> 开启</label>
	</td>
</tr>
<tbody id="alipay"<%if cbool(API_AlipayEnable)=false then response.write " style='display:none'"%>>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>支付宝合作者身份ID：</strong></td>
	<td><input type="text" class="textbox" name="API_AlipayPartner" size="35" value="<%=API_AlipayPartner%>"> 
	<font color=red>如果还没有与支付宝签约，请<a href="https://b.alipay.com/order/slaverIndex.htm?rewardIds=vtq05uWfOIk-Ht9P1HzAYTlNX7GOvULv" target="_blank">点此申请</a>。</font>
	</td>
</tr>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>安全检验码Key：</strong></td>
	<td><input type="text" class="textbox" name="API_AlipayKey" size="35" value="<%=API_AlipayKey%>"> 
	</td>
</tr>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>支付宝快捷登录后跳转的地址：</strong></td>
	<td><input type="text" class="textbox" name="api_alipayreturnurl" style="background:#efefef" readonly size="45" value="<%=KS.GetDomain &"api/alipay/return_url.asp"%>"> <font class="tips">支付宝快捷登录成功后跳转的地址,不可改。</font>
	</td>
</tr>
</tbody>




<tr class="tdbg">
	<td height="30" colspan="2" class="clefttitle" style="text-align:center"> <strong>以下整合动网、Oblog及Oask之类的程序，已不常用,建议不要开启。</strong>
	</td>
</tr>

<tr class="tdbg">
	<td height="30" width="25%" class="clefttitle" align="right"><strong>是否开启多系统整合程序：</strong></td>
	<td>
	<input type="radio" name="API_Enable" onclick="$('#api').hide()" value="false"<%
	If Not API_Enable Then Response.Write " checked"
	%>> 关闭&nbsp;&nbsp;
	<input type="radio" name="API_Enable" onclick="$('#api').show()" value="true"<%
	If API_Enable Then Response.Write " checked"
	%>> 开启
	</td>
</tr>
<tbody id="api"<%if Cbool(Api_Enable)=false Then response.write " style='display:none'"%>>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>设置系统密钥：</strong></td>
	<td><input type="text" name="API_ConformKey" size="35" value="<%=API_ConformKey%>"> 
		<font color="red">系统整合，必须保证与其它系统设置的密钥一致。</font>
	</td>
</tr>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>是否除错：</strong></td>
	<td>
	<input type="radio" name="API_Debug" value="false"<%
	If Not API_Debug Then Response.Write " checked"
	%>> 否&nbsp;&nbsp;
	<input type="radio" name="API_Debug" value="true"<%
	If API_Debug Then Response.Write " checked"
	%>> 是&nbsp;&nbsp;<font color="red">如果整合的论坛程序和科汛程序的用户数据不同步，你可以选择“是”</font>
	</td>
</tr>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>整合程序的接口文件路径：</strong></td>
	<td><textarea name="API_Urls" rows="6" cols="70"><%=API_Urls%></textarea></td>
</tr>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>整合用户登录后转向URL：</strong></td>
	<td><input type="text" name="API_LoginUrl" size="45" value="<%=API_LoginUrl%>"> 
		<font color="red">不设置请输入“0”。</font>
	</td>
</tr>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>整合用户注册后转向URL：</strong></td>
	<td><input type="text" name="API_ReguserUrl" size="45" value="<%=API_ReguserUrl%>"> 
		<font color="red">不设置请输入“0”。</font>
	</td>
</tr>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>整合用户注销后转向URL：</strong></td>
	<td><input type="text" name="API_LogoutUrl" size="45" value="<%=API_LogoutUrl%>"> 
		<font color="red">不设置请输入“0”。</font>
	</td>
</tr>
</form>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>使用说明：</strong></td>
	<td><font color="blue">如果有多个程序整合，接口之间用半角"|"分隔<br />例如：http://你的论坛网址/dv_dpo.asp|http://你的网站地址/博客安装目录/oblogresponse.asp;<br />
	本系统的接口路径：<font color="red"><%=KS.GetDomain%>api/api_response.asp</font><br /></font></td>
</tr>
</tbody>
</table>
<script>
 function CheckForm()
 {
  document.all.myform.submit();
 }
</script>
<%
End Sub

Sub SaveConformify()
	Dim XslDoc,XslNode,Xsl_Files
	Xsl_Files = API_Path & "api.config"
	Xsl_Files = Server.MapPath(Xsl_Files)
	Set XslDoc = KS.InitialObject("Msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
	If Not XslDoc.Load(Xsl_Files) Then
		Response.Write "初始数据不存在！"
		Response.End
	Else
		Set XslNode = XslDoc.documentElement.selectSingleNode("rs:data/z:row")
		XslNode.attributes.getNamedItem("api_enable").text = Trim(Request.Form("API_Enable"))
		XslNode.attributes.getNamedItem("api_conformkey").text = ChkRequestForm("API_ConformKey")
		XslNode.attributes.getNamedItem("api_urls").text = ChkRequestForm("API_Urls")
		XslNode.attributes.getNamedItem("api_debug").text = ChkRequestForm("API_Debug")
		XslNode.attributes.getNamedItem("api_loginurl").text = ChkRequestForm("API_LoginUrl")
		XslNode.attributes.getNamedItem("api_reguserurl").text = ChkRequestForm("API_ReguserUrl")
		XslNode.attributes.getNamedItem("api_logouturl").text = ChkRequestForm("API_LogoutUrl")
		'XslNode.attributes.setNamedItem(XslDoc.createNode(2,"date","")).text = Now()
		'XslNode.appendChild(XslDoc.createNode(1,"pubDate","")).text = Now()
		XslNode.attributes.getNamedItem("api_quicklogin").text =trim(Request.Form("API_QuickLogin"))
		XslNode.attributes.getNamedItem("api_groupid").text =trim(Request.Form("GroupID"))
		XslNode.attributes.getNamedItem("api_qqenable").text =trim(Request.Form("API_QQEnable"))
		XslNode.attributes.getNamedItem("api_qqappid").text =ChkRequestForm("API_QQAppId")
		XslNode.attributes.getNamedItem("api_qqappkey").text =ChkRequestForm("API_QQAppKey")
		XslNode.attributes.getNamedItem("api_qqcallback").text =ChkRequestForm("API_QQCallBack")
		
		XslNode.attributes.getNamedItem("api_alipayenable").text =trim(Request.Form("API_AlipayEnable"))
		XslNode.attributes.getNamedItem("api_alipaypartner").text =ChkRequestForm("API_AlipayPartner")
		XslNode.attributes.getNamedItem("api_alipaykey").text =ChkRequestForm("API_AlipayKey")
		XslNode.attributes.getNamedItem("api_alipayreturnurl").text =ChkRequestForm("API_AlipayReturnUrl")
		
		XslNode.attributes.getNamedItem("api_sinaenable").text =trim(Request.Form("API_SinaEnable"))
		XslNode.attributes.getNamedItem("api_sinaid").text =ChkRequestForm("API_SinaId")
		XslNode.attributes.getNamedItem("api_sinakey").text =ChkRequestForm("API_SinaKey")
		XslNode.attributes.getNamedItem("api_sinacallback").text =ChkRequestForm("API_SinaCallBack")

		XslDoc.save Xsl_Files
		Set XslNode = Nothing
	End If
	Set XslDoc = Nothing
	Response.Write ("<script>alert('恭喜您！保存设置成功。');location.href='KS.Api.asp';</script>")
End Sub
Function ChkRequestForm(reform)
	Dim strForm
	strForm = Trim(Request.Form(reform))
	If IsNull(strForm) Then
		strForm = "0"
	Else
		strForm = Replace(strForm, Chr(0), vbNullString)
		strForm = Replace(strForm, Chr(34), vbNullString)
		strForm = Replace(strForm, "'", vbNullString)
		strForm = Replace(strForm, """", vbNullString)
	End If
	If strForm = "" Then strForm = "0"
	ChkRequestForm = strForm
End Function

%>
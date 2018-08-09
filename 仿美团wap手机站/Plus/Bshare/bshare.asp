<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../Plus/Session.asp"-->
<!--#include file="../../Plus/md5.asp"-->
<%

Dim Chk:Set Chk=New LoginCheckCls1
Chk.Run()
Set Chk=Nothing
Dim KS:Set KS=New PublicCls

Dim Bshare_Open,Bshare_UUID,Bshare_PassWord,Bshare_RePassWord,Bshare_Domain,Bshare_UserName,Bshare_Secret

Dim Action:Action = LCase(Request("action"))
LoadbshareConfig
Select Case Trim(Action)
	Case "save"		Call savebshare
	Case "show"	    Call show
	Case "getstyle" Call GetStyle
	Case "getdata"  Call GetData
	Case Else		Call showmain
End Select

Sub show
 Response.Write "<script>window.open('http://intf.cnzz.com/user/companion/newasp_login.php?site_id=" & Bshare_UUID & "&password=" & Bshare_password & "');history.back();</script>"
End Sub

Sub ShowMain

If Len(Bshare_Domain)<3 Then Bshare_Domain=KS.GetAutoDomain

Response.Write "<html><head><title>多系统整合接口设置</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" & vbCrLf
Response.Write "<link href='../wss/Admin_Style.css' rel='stylesheet' type='text/css'></head>" & vbCrLf
Response.Write "</head>"
Response.Write "<body bgcolor=""#FFFFFF"" topmargin=""0"" leftmargin=""0"" scroll=no>"
Response.Write "<ul id='menu_top' style='text-align:center;padding-top:10px;font-weight:bold'>bShare分享插件</ul>"
%>
<script src="../../ks_inc/jquery.js"></script>
<table border="0" align="center" cellpadding="3" cellspacing="1" width="100%" class="border">
<%if bshare_open="true" then%>
<tr class="tdbg">
	<td class="clefttitle" colspan="2" height="30"><strong>您已经开通bShare服务：</strong></td>
</tr>
<tr class="tdbg">
	<td class="clefttitle" width="20%" align="right" height="30"><u>网站域名</u>：</td>
	<td width="80%"><%=Bshare_Domain%></td>
</tr>
<tr class="tdbg">
	<td class="clefttitle" width="20%" align="right" height="30"><u>用户名</u>：</td>
	<td width="80%"><%=Bshare_UserName%></td>
</tr>
<tr class="tdbg">
	<td class="clefttitle" width="20%" align="right" height="30"><u>UUID</u>：</td>
	<td width="80%"><%=Bshare_uuid%></td>
</tr>
<tr class="tdbg">
	<td class="clefttitle" width="20%" align="right" height="30"><u>SECRET</u>：</td>
	<td width="80%"><%=Bshare_secret%></td>
</tr>

<%else%>
<form name="myform" method="post" action="?action=save">
<tr class="tdbg">
	<td class="clefttitle" width="20%" align="right"><u>网站域名</u>：</td>
	<td width="80%"><input type="text" name="Bshare_Domain" size="35" value="<%=Bshare_Domain%>"> 
		<font color="red">* </font>
	</td>
</tr>
<tr class="tdbg">
	<td class="clefttitle" align="right"><u>账号类型</u>：</td>
	<td><input type="radio" name="Bshare_Type" value="1" onclick="$('#rpass').show();" checked="checked"/>新注册
	    <input type="radio" name="Bshare_Type" value="2" onclick="$('#rpass').hide();"/>已经有账号
	</td>
</tr>
<tr class="tdbg">
	<td class="clefttitle" align="right"><u>用 户 名</u>：</td>
	<td><input type="text" name="Bshare_UserName" size="35" value="<%=Bshare_UserName%>"> 
		<font color="red">* 填写Email</font>
	</td>
</tr>

<tr class="tdbg">
	<td class="clefttitle" align="right"><u>登录密码</u>：</td>
	<td><input type="password" name="Bshare_PassWord" size="35" value="<%=Bshare_PassWord%>"> 
		<font color="red">* </font>
	</td>
</tr>

<tr class="tdbg" id="rpass">
	<td class="clefttitle" align="right"><u>确定密码</u>：</td>
	<td class="clefttitle"><input type="password" name="Bshare_RePassWord" size="35" value=""> 
		<font color="red">* </font>
	</td>
</tr>
<tr class="tdbg">
	<td colspan="2" align="center">
	<input type="submit" value="保存设置" name="B1" class="Button"></td>
</tr>
</form>
<tr>
	<td class="clefttitle" colspan="2"><b>说明</b><br/>&nbsp;&nbsp;bShare不止是一个分享按钮。bShare是全球中文互联网最强大的社交分享引擎！ 只需一个按钮，就能为您的网站注入社交化功能！<br/> 
bShare智能分享引擎让您的用户可以轻松地将最喜欢的内容分享到社交网站、微博上与好友分享。用户无须离开您的网站，就能快速地进行分享，继续浏览您的网站！
</td>
</tr>
<%end if%>


</table>
<%if bshare_open="true" then%>

<br/>
<table border="0" align="center" cellpadding="3" cellspacing="1" width="100%" class="border">
<tr class="tdbg">
	<td class="clefttitle" colspan="2" height="30"><strong>前台调用代码：</strong></td>
</tr>
<tr class="tdbg">
	<td colspan="2">
	请把以下代码复制到内容页模板里想显示的地方即可<br/>
	<textarea name="bsharecode" style="width:450px;height:90px"><a class="bshareDiv" href="http://www.bshare.cn/share">分享按钮</a><script language="javascript" type="text/javascript" src="http://static.bshare.cn/b/button.js#uuid=<%=Bshare_uuid%>&amp;style=2&amp;textcolor=#000&amp;bgcolor=none&amp;bp=qqmb,sinaminiblog,sohubai,renren&amp;ssc=false&amp;sn=true&amp;text=分享到"></script></textarea>
	
	<div style="margin-top:16px;padding-left:10px">
	  <strong>效果预览：</strong><br/>
	  <a class="bshareDiv" href="http://www.bshare.cn/share">分享按钮</a><script language="javascript" type="text/javascript" src="http://static.bshare.cn/b/button.js#uuid=<%=Bshare_uuid%>&style=2&textcolor=#000&bgcolor=none&bp=qqmb,sinaminiblog,sohubai,renren&ssc=false&sn=true&text=分享到"></script>
	  </div>
	 
	 Tips:如果您对以上样式不满意还可以<input type="button" onclick="getStyle();" class="button" value="点此获取更多样式"/> 
	  
	</td>
</tr>
</table>
<script type="text/javascript">
function getStyle(){
	new parent.KesionPopup().PopupCenterIframe('选择Bshare分享插件样式','../plus/bshare/Bshare.asp?action=GetStyle',720,400,'no')
}
</script>
<%
end if

End Sub

Sub GetStyle()
%>
<style type="text/css">
iframe { border-style: none; }
body { margin: 0px;padding: 0px; }
</style>
<iframe src="http://www.bshare.cn/moreStylesEmbed?uuid=<%=Bshare_UUID%>&bp=qqmb%2csinaminiblog%2csohubai%2cbaiduhi%2crenren%2cbgoogle" name="bshare" width="710px" height="400px" scrolling="yes">
<%
End Sub

Sub GetData()
 if cbool(bshare_open)<>true then
   ks.die "<script>alert('您还没有开通设置bshare，按确定转向设置页面!');location.href='bshare.asp';</script>"
 end if
 Dim TS:TS=ToUnixTime(now,8)&"000"
 Dim Sign:Sign=md5("ts=" & ts & "uuid=" & bshare_uuid & bshare_secret,32)
%>
<style type="text/css">
iframe { border-style: none; }
body { margin: 0px;padding: 0px; }
</style>
<iframe src="http://www.bshare.cn/publisherStatisticsEmbed?uuid=<%=bshare_uuid%>&ts=<%=ts%>&sig=<%=sign%>" name="bshare" style="width:100%;height:100%" width="800" height="600" scrolling="yes">
<%
End Sub

Sub savebshare()
	If Len(Request.Form("Bshare_domain")) < 3 Then
		response.write "<script>alert('你的域名有误!');history.back();</script>"
	End If
	Dim XmlDoc,XmlNode,Xml_Files
	Dim Bshare_Type : Bshare_Type = KS.ChkClng(KS.G("Bshare_Type"))
	Xml_Files = "bshare.config"
	Xml_Files = Server.MapPath(Xml_Files)
	Set XmlDoc = KS.InitialObject("Msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
	If XmlDoc.Load(Xml_Files) Then
		Set XmlNode = XmlDoc.documentElement.selectSingleNode("rs:data/z:row[@id=0]")
		'If Bshare_Type = 2 Then
		'	XmlNode.attributes.getNamedItem("Bshare_UUID").text = KS.S("Bshare_UUID")
		'	XmlNode.attributes.getNamedItem("Bshare_password").text = KS.S("Bshare_password")
		'Else
			If Len(Request.Form("Bshare_domain")) > 3 Then
				Dim strbshareData
				Dim strURL,strDomain,strKey
				Bshare_domain = KS.G("Bshare_domain")
				Bshare_UserName=Request.Form("Bshare_UserName")
				Bshare_PassWord=Request.Form("Bshare_PassWord")
				Bshare_RePassWord=Request.Form("Bshare_RePassWord")
				If Bshare_UserName="" Then KS.Die "<script>alert('请输入您的用户名!');history.back();</script>"
				If Bshare_PassWord="" Then KS.Die "<script>alert('请输入登录密码!');history.back();</script>"
				If Bshare_Type <>2 and Bshare_PassWord<>Bshare_RePassWord Then KS.Die "<script>alert('请输入两次输入的密码不一致!');history.back();</script>"
				
				strURL = "http://api.bshare.cn/analytics/reguuid.json?email="  & Bshare_UserName & "&password=" & Bshare_PassWord & "&domain=" & Bshare_domain & "&source=kesion"
				strbshareData = GetbshareData(strURL)
				
				
				If InStr(strbshareData,"{""uuid"":""") > 0 Then
					Dim bshareArray
					bshareArray = Split(strbshareData, ",")
					XmlNode.attributes.getNamedItem("bshare_uuid").text = trim(replace(replace(bshareArray(0),"{""uuid"":""",""),"""",""))
					XmlNode.attributes.getNamedItem("bshare_secret").text = trim(replace(replace(bshareArray(1),"""secret"":""",""),"""}",""))
					XmlNode.attributes.getNamedItem("bshare_domain").text = Bshare_domain
					XmlNode.attributes.getNamedItem("bshare_password").text = Bshare_password
					XmlNode.attributes.getNamedItem("bshare_username").text = Bshare_username
					XmlNode.attributes.getNamedItem("bshare_open").text = "true"
				Else
					Response.Write "<script>alert('申请bshare失败!错误代码：" & strbshareData  &"');history.back();</script>"
					Exit Sub
				End If
			End If
		'End If
		XmlDoc.save Xml_Files
		Set XmlNode = Nothing
	End If
	Set XmlDoc = Nothing
	 Response.Write "<script>alert('恭喜您！申请开通bshare成功。');location.href='bshare.asp';</script>"
End Sub
'生成时间戳 
Function ToUnixTime(strTime, intTimeZone)
If IsEmpty(strTime) or Not IsDate(strTime) Then strTime = Now
If IsEmpty(intTimeZone) or Not isNumeric(intTimeZone) Then intTimeZone = 0
ToUnixTime = DateAdd("h",-intTimeZone,strTime)
ToUnixTime = DateDiff("s","1970-1-1 0:0:0", ToUnixTime)
End Function

Function GetbshareData(ByVal strURL)
	On Error Resume Next
	Dim xmlhttp,TextBody
	Set xmlhttp = KS.InitialObject("msxml2.ServerXMLHTTP")
	xmlhttp.setTimeouts 65000, 65000, 65000, 65000
	xmlhttp.Open "GET",strURL,false
	xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	xmlhttp.send()
	'TextBody = strAnsi2Unicode(xmlhttp.responseBody)
	TextBody = xmlhttp.responseText
	Set xmlhttp = Nothing
	GetbshareData = TextBody
End Function
Function strAnsi2Unicode(asContents)
	Dim len1,i,varchar,varasc
	strAnsi2Unicode = ""
	len1=LenB(asContents)
	If len1=0 Then Exit Function
	  For i=1 to len1
	  	varchar=MidB(asContents,i,1)
	  	varasc=AscB(varchar)
	  	If varasc > 127  Then
	  		If MidB(asContents,i+1,1)<>"" Then
	  			strAnsi2Unicode = strAnsi2Unicode & chr(ascw(midb(asContents,i+1,1) & varchar))
	  		End If
	  		i=i+1
	     Else
	     	strAnsi2Unicode = strAnsi2Unicode & Chr(varasc)
	     End If	
	  Next
End Function
Sub LoadbshareConfig()
Dim XmlDoc,XmlNode,Xml_Files
Xml_Files = "bshare.config"
Xml_Files = Server.MapPath(Xml_Files)
Set XmlDoc = KS.InitialObject("Msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
If Not XmlDoc.Load(Xml_Files) Then
			Bshare_Open = true
			Bshare_UUID = ""
			Bshare_PassWord = ""
			Bshare_Domain = KS.GetAutoDomain
			Bshare_UserName = ""
Else
			Set XmlNode	= XmlDoc.documentElement.selectSingleNode("rs:data/z:row[@id=0]")
			Bshare_Open = XmlNode.getAttribute("bshare_open")
			Bshare_UUID = XmlNode.getAttribute("bshare_uuid")
			Bshare_SECRET= XmlNode.getAttribute("bshare_secret")
			Bshare_UserName=XmlNode.getAttribute("bshare_username")
			Bshare_PassWord = XmlNode.getAttribute("bshare_password")
			Bshare_Domain = XmlNode.getAttribute("bshare_domain")
			Bshare_UserName = XmlNode.getAttribute("bshare_username")
			Set XmlNode = Nothing
End If
Set XmlDoc = Nothing
End Sub
%>
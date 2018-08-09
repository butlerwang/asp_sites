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

Dim Wss_IsUsed,Wss_SiteID,Wss_PassWord,Wss_Domain,Wss_Key

Dim Action:Action = LCase(Request("action"))
LoadWssConfig
Select Case Trim(Action)
	Case "save"
		Call savewss
	Case "show"
	    Call show
	Case Else
		Call showmain
End Select

Sub show
 Response.Write "<script>window.open('http://intf.cnzz.com/user/companion/newasp_login.php?site_id=" & Wss_SiteID & "&password=" & Wss_password & "');history.back();</script>"
End Sub

Sub ShowMain

If Len(Wss_Domain)<3 Then Wss_Domain=KS.GetAutoDomain

Response.Write "<html><head><title>多系统整合接口设置</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'></head>" & vbCrLf
Response.Write "</head>"
Response.Write "<body bgcolor=""#FFFFFF"" topmargin=""0"" leftmargin=""0"" scroll=no>"
Response.Write "<ul id='menu_top' style='text-align:center;padding-top:10px;font-weight:bold'>"
Response.Write "     WSS流量统计设置</ul>"
%>
<table border="0" align="center" cellpadding="3" cellspacing="1" width="100%" class="border">
<form name="myform" method="post" action="?action=save">
<tr class="tdbg">
	<td class="clefttitle" width="20%" align="right"><u>WSS统计域名</u>：</td>
	<td width="80%"><input type="text" name="Wss_Domain" size="35" value="<%=Wss_Domain%>"> 
		<font color="red">* </font>
	</td>
</tr>
<tr class="tdbg">
	<td class="clefttitle" align="right"><u>WSS统计站点ID</u>：</td>
	<td><input type="text" name="Wss_SiteID" size="35" value="<%=Wss_SiteID%>"> 
		<font color="red">* 如果你已经注册过WSS请输入你的站点ID</font>
	</td>
</tr>
<tr class="tdbg">
	<td class="clefttitle" align="right"><u>WSS统计登录密码</u>：</td>
	<td><input type="text" name="Wss_PassWord" size="35" value="<%=Wss_PassWord%>"> 
		<font color="red">* 如果你已经注册过WSS请输入你的登录密码</font>
	</td>
</tr>
<tr class="tdbg">
	<td class="clefttitle" align="right"><u>是否开启WSS统计功能</u>：</td>
	<td>
	<input type="radio" name="wss_isused" value="0"<%
	If Wss_IsUsed=0 Then Response.Write " checked"
	%>> 关闭&nbsp;&nbsp;
	<input type="radio" name="wss_isused" value="1"<%
	If Wss_IsUsed=1 Then Response.Write " checked"
	%>> 开启&nbsp;&nbsp;
	</td>
</tr>
<tr class="tdbg">
	<td class="clefttitle" align="right"><u>申请WSS统计</u>：</td>
	<td class="clefttitle"><input type="checkbox" name="apply" value="1"/> 
		<font color="red">* 如果你是第一次申请请选择</font>
	</td>
</tr>
<tr class="tdbg">
	<td colspan="2" align="center">
	<input type="submit" value="保存设置" name="B1" class="Button"></td>
</tr>
</form>
<tr>
	<td class="clefttitle" colspan="2"><b>说明</b><br/>&nbsp;&nbsp;<a href="http://intf.cnzz.com/" target="_blank">WSS</a> 一直致力于精确时实的网站流量统计分析，并且通过不断的努力为贵网站提供更快速、更直观、更准确的统计服务。<br/><br/>
	<b>申请失败情况下错误代码：</b><br/>
&nbsp;&nbsp;-1 表示key有误<br/>
&nbsp;&nbsp;-2 表示该域名长度有误（1~64）,<br/>
&nbsp;&nbsp;-3 表示域名输入有误（比如输入汉字）,<br/>
&nbsp;&nbsp;-4 表示域名插入数据库有误<br/>
&nbsp;&nbsp;-5 表示同一个IP用户调用页面超过阀值，阀值暂定为10。
</td>
</tr>

<tr>
	<td class="clefttitle" align="right"><u>统计代码</u>：</td>
	<td class="clefttitle">
	<textarea name="wsscode" rows="3" cols="70">&lt;script src='http://pw.cnzz.com/c.php?id=<%=Wss_SiteID%>&l=2' language='JavaScript' charset='utf-8'&gt;&lt;/script&gt;</textarea>
	<br><font color=red>把以上代码复制到你要统计的网页模板里即可</font>
	</td>
</tr>

</table>
<%

End Sub

Sub savewss()
	If Len(Request.Form("wss_domain")) < 3 Then
		response.write "<script>alert('你的域名有误!');history.back();</script>"
	End If
	Dim XmlDoc,XmlNode,Xml_Files
	Dim apply : apply = KS.ChkClng(KS.G("apply"))
	Xml_Files = "wss.config"
	Xml_Files = Server.MapPath(Xml_Files)
	Set XmlDoc = KS.InitialObject("Msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
	If XmlDoc.Load(Xml_Files) Then
		Set XmlNode = XmlDoc.documentElement.selectSingleNode("rs:data/z:row[@id=0]")
		If apply = 0 Then
			XmlNode.attributes.getNamedItem("wss_siteid").text = KS.S("wss_siteid")
			XmlNode.attributes.getNamedItem("wss_password").text = KS.S("wss_password")
		Else
			If Len(Request.Form("wss_domain")) > 3 Then
				Dim strWssData
				Dim strURL,strDomain,strKey
				strDomain = KS.G("wss_domain")
				strKey = Md5(strDomain&"Ioi6pPdV",32)
				strURL = "http://intf.cnzz.com/user/companion/kesion.php?domain="&strDomain&"&key=" & strKey
				strWssData = GetWssData(strURL)
				If InStr(strWssData,"@") > 0 Then
					Dim WssArray
					WssArray = Split(strWssData, "@")
					XmlNode.attributes.getNamedItem("wss_siteid").text = Trim(WssArray(0))
					XmlNode.attributes.getNamedItem("wss_password").text = Trim(WssArray(1))
				Else
					Response.Write "<script>alert('申请WSS失败!错误代码：" & strWssData & strKey &"');history.back();</script>"
					Exit Sub
				End If
			End If
		End If
		XmlNode.attributes.getNamedItem("wss_isused").text = KS.ChkCLng(KS.S("wss_isused"))
		XmlNode.attributes.getNamedItem("wss_domain").text = KS.G("wss_domain")
		XmlDoc.save Xml_Files
		Set XmlNode = Nothing
	End If
	Set XmlDoc = Nothing
	 Response.Write "<script>alert('恭喜您！保存WSS设置成功。');location.href='wss.asp';</script>"
End Sub
Function GetWssData(ByVal strURL)
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
	GetWssData = TextBody
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
Sub LoadWssConfig()
Dim XmlDoc,XmlNode,Xml_Files
Xml_Files = "wss.config"
Xml_Files = Server.MapPath(Xml_Files)
Set XmlDoc = KS.InitialObject("Msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
If Not XmlDoc.Load(Xml_Files) Then
			Wss_IsUsed = 0
			Wss_SiteID = ""
			Wss_PassWord = ""
			Wss_Domain = KS.GetAutoDomain
			Wss_Key = ""
Else
			Set XmlNode	= XmlDoc.documentElement.selectSingleNode("rs:data/z:row[@id=0]")
			Wss_IsUsed = KS.ChkClng(XmlNode.getAttribute("wss_isused"))
			Wss_SiteID = XmlNode.getAttribute("wss_siteid")
			Wss_PassWord = XmlNode.getAttribute("wss_password")
			Wss_Domain = XmlNode.getAttribute("wss_domain")
			Wss_Key = XmlNode.getAttribute("wss_key")
			Set XmlNode = Nothing
End If
Set XmlDoc = Nothing
End Sub
%>
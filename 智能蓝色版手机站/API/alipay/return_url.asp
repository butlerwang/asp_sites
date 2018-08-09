﻿<%
' 功能：支付宝页面跳转同步通知页面
' 版本：3.2
' 日期：2012-03-31
' 说明：
' 以下代码只是为了方便商户测试而提供的样例代码，商户可以根据自己网站的需要，按照技术文档编写,并非一定要使用该代码。
' 该代码仅供学习和研究支付宝接口使用，只是提供一个参考。
	
' //////////////页面功能说明//////////////
' 该页面可在本机电脑测试
' 可放入HTML等美化页面的代码、商户业务逻辑程序代码
' 该页面可以使用ASP开发工具调试，也可以使用写文本函数LogResult进行调试，该函数已被默认关闭，见alipay_notify.asp中的函数VerifyReturn
'////////////////////////////////////////
%>
<!--#include file="class/alipay_notify.asp"-->
<%


'计算得出通知验证结果
Set objNotify = New AlipayNotify
sVerifyResult = objNotify.VerifyReturn()

If sVerifyResult Then	'验证成功
	'*********************************************************************
	'请在这里加上商户的业务逻辑程序代码

	'——请根据您的业务逻辑来编写程序（以下代码仅作参考）——
    '获取支付宝的通知返回参数，可参考技术文档中页面跳转同步通知参数列表
	
	Dim nickname,PassWord,user_id,token
    user_id = KS.DelSQL(request.QueryString("user_id"))	'支付宝用户id
    token	= KS.DelSQL(request.QueryString("token"))		'授权令牌
	
	'执行商户的业务程序
	'dim varItem
	'For Each varItem in Request.QueryString
	'		response.write varItem & "=" & Request(varItem) & "<br/>"
	'Next 
	
	session("real_name")=KS.S("real_name")
	session("token")=token
	session("user_id")=user_id
	response.Redirect("alipaybind.asp")
		
	'response.Write "验证成功<br />"
	'response.Write "token:"&token

	'etao专用
	If request.QueryString("target_url") <> "" Then
		'程序自动跳转到target_url参数指定的url去
	End If

	'——请根据您的业务逻辑来编写程序（以上代码仅作参考）——
	
	'*********************************************************************
Else '验证失败
    '如要调试，请看alipay_notify.asp页面的VerifyReturn函数，比对sign和mysign的值是否相等，或者检查responseTxt有没有返回true
    Response.Write "验证失败"
End If

Set KS=Nothing
CloseConn
%>

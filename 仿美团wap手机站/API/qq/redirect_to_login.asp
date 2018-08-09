<!--#include file="config.asp"-->
<%
response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Response.CharSet="utf-8"

If cbool(API_QQEnable)=false Then KS.Die "<script>alert('对不起，本站没有开启QQ快速登录功能!');location.href='../../user/login/';</script>"



function redirect_to_login()
	Session("state") = year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&KS.MakeRandom(10)
	dim url:url="https://graph.qq.com/oauth2.0/authorize?response_type=code&client_id="&appid&"&redirect_uri="&server.urlencode(callback)&"&state="&Session("state") &"&scope=get_user_info,add_share,check_page_fans,add_t,del_t,add_pic_t,get_repost_list,get_info,get_other_info,get_fanslist,get_idolist,add_idol,del_idol"
	 response.write "<div style='margin-top:90px;color:#666;font-size:16px;text-align:center;'><img src='" & KS.GetDomain &"images/default/loadingAnimation.gif'/><br/><br/>正在转向qq登录授权页面，请稍候！！！如果长时间没有反应请<a href=""javascript:;"" onclick=""top.location.href='" & url & "';"" style='color:red'>点此跳转</a>。</div>"

	response.redirect url
end function

Call redirect_to_login()

%>
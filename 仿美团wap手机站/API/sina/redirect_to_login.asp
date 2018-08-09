<!--#include file="../../conn.asp"-->
<!--#include file="../../ks_cls/kesion.commoncls.asp"-->
<!--#include file="../cls_api.asp"-->
<%
If cbool(API_SinaEnable)=false Then KS.Die "<script>alert('对不起，本站没有开启新浪微博登录功能!');location.href='../../user/login/';</script>"

function redirect_to_login()
	response.redirect "https://api.weibo.com/oauth2/authorize?client_id=" & Api_SinaId &"&redirect_uri="&server.URLEncode(api_sinaCallBack)&"&response_type=code"
end function

Call redirect_to_login()

%>
<!--#include file="../../conn.asp"-->
<!--#include file="../../ks_cls/kesion.commoncls.asp"-->
<!--#include file="../cls_api.asp"-->
<%
If EnabledSubDomain Then
	Response.Cookies(KS.SiteSn).domain=RootDomain					
Else
    Response.Cookies(KS.SiteSn).path = "/"
End If

call sina_callback()

function sina_callback()
        if REQUEST("code")="" then ks.die "error access!"
	    Dim token_url,result
		token_url = "https://api.weibo.com/oauth2/access_token"
        result = file_get_contents(token_url,"post","client_id=" & API_SinaId &"&client_secret=" &API_SinaKey&"&grant_type=authorization_code&redirect_uri=" & server.URLEncode(API_SinaCallBack) &"&code="&REQUEST("code"))
		
		dim obj:set obj = getjson(result)
		if instr(result,"error")<>0 then
			if isobject(obj) Then
			  ks.echo "<h3>error:</h3>" & obj.error
			  ks.echo "<h3>error_code:</h3>" & obj.error_code
			  ks.echo "<h3>msg:</h3>" & obj.error_description
			End If
			set obj=nothing
			ks.die ""
		Else
		   if isobject(obj) Then
			Response.Cookies(KS.SiteSn).Expires = Date + 365
			Response.Cookies(KS.SiteSn)("sina_access_token") = obj.access_token
			Response.Cookies(KS.SiteSn)("sinaid") = obj.uid
		   End If
		   set obj=nothing
		End If
End Function

response.write "<div style='margin-top:90px;color:#666;font-size:16px;text-align:center;'><img src='" & KS.GetDomain &"images/default/loadingAnimation.gif'/><br/><br/>正在登录中，请稍候！！！如果长时间没有反应请<a href=""sinabind.asp"" style='color:red' target=""_blank"">点此跳转</a>。</div>"

if ks.isnul(ks.c("sinaid")) then 
    ks.die "没有返回uid!"
	set ks=nothing
	closeconn
else
  set ks=nothing
  closeconn
  response.Write "<script>top.location.href='sinabind.asp';</script>"
end  if
%>
<!--#include file="../../plus/md5.asp"-->
<!--#include file="config.asp"-->
<%
If EnabledSubDomain Then
	Response.Cookies(KS.SiteSn).domain=RootDomain					
Else
    Response.Cookies(KS.SiteSn).path = "/"
End If

call qq_callback()
call get_openid()
function qq_callback()
    if(lcase(Request("state")) = lcase(Session("state"))) Then
	    Dim token_url,result
        token_url = "https://graph.qq.com/oauth2.0/token"

        result = file_get_contents(token_url,"get","grant_type=authorization_code&client_id="&appid&"&redirect_uri="&server.URLEncode(callback) & "&client_secret="&appkey &"&code="&REQUEST("code"))
		
		if instr(result,"callback")<>0 then
            dim lpos:lpos = instr(result, "(")
            dim rpos:rpos = instr(result, ")")
            result  = mid(result, lpos + 1, rpos - lpos -1)
			
			dim obj:set obj = getjson(result)
			if isobject(obj) Then
			  ks.echo "<h3>error:</h3>" & obj.error
			  ks.echo "<h3>msg:</h3>" & obj.error_description
			End If
			set obj=nothing
			ks.die ""
		end if
		if result<>"" then
			dim arr:set arr=parse_str(result)
			Response.Cookies(KS.SiteSn).Expires = Date + 365
			Response.Cookies(KS.SiteSn)("access_token") = arr("access_token")
			Response.Cookies(KS.SiteSn)("qqappid") = appid
		else
		  ks.die "error!"
		end if
    Else 
        KS.Echo "The state does not match. You may be a victim of CSRF."
    End If
End Function

function get_openid()
    dim graph_url:graph_url = "https://graph.qq.com/oauth2.0/me"
    dim result:result=file_get_contents(graph_url,"get","access_token="&ks.c("access_token"))
    if instr(result,"callback")<>0 then
            dim lpos:lpos = instr(result, "(")
            dim rpos:rpos = instr(result, ")")
            result  = mid(result, lpos + 1, rpos - lpos -1)
			dim obj:set obj = getjson(result)
			if isobject(obj) Then
			 Response.Cookies(KS.SiteSn).Expires = Date + 365
			 Response.Cookies(KS.SiteSn)("openid") = obj.openid
			End If
			set obj=nothing
	end if
End Function

response.write "<div style='margin-top:90px;color:#666;font-size:16px;text-align:center;'><img src='" & KS.GetDomain &"images/default/loadingAnimation.gif'/><br/><br/>正在登录中，请稍候！！！如果长时间没有反应请<a href=""qqbind.asp"" target=""parent"" style='color:red'>点此跳转</a>。</div>"

if ks.isnul(ks.c("openid")) then 
    ks.die "没有返回openid!"
	set ks=nothing
	closeconn
else
  set ks=nothing
  closeconn
  response.Write "<script>top.location.href='qqbind.asp';</script>"
  response.Redirect("qqbind.asp")
end  if
%>
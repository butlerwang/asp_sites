<!--#include file="../../conn.asp"-->
<!--#include file="../../ks_cls/kesion.commoncls.asp"-->
<!--#include file="../../plus/md5.asp"-->
<!--#include file="../cls_api.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
dim openid,username,password,action,rs
if ks.isnul(KS.C("sinaid")) then ks.die "没有返回uid!"
action=KS.S("Action")
If Action="check" Then
	openid=ks.s("openid")
	username=ks.r(ks.s("username"))
	password=md5(KS.R(ks.s("password")),16)
	if ks.c("sinaid")<>openid then
	  ks.die "请不要非法绑定!"
	end if
	set rs=server.createobject("adodb.recordset")
	rs.open "select top 1 * from ks_user where username='" & username & "' and password='" & password & "'",conn,1,1
	if rs.eof and rs.bof then
	  rs.close:set rs=nothing
	  ks.die "<script>alert('对不起，您输入的账号不存在或是密码不正确，请重输!');history.back(-1);</script>"
	else
		'绑定到已有账号
		conn.execute("update ks_user set sinatoken='" & KS.C("sina_access_token") & "',sinaid='" & openid & "' where username='" & username & "'")
		'调用登录
		Call DoLogin(username,password)
	end if
ElseIf Action="doreg" Then
        Call DoRegSave(2)
Else
    '===================绑定处理=================================
		
		dim ret,msg,nickname,figureurl,sex
		dim resultxml:resultxml=get_user_info(2,ks.c("sina_access_token"),KS.C("sinaid"))
		dim obj:set obj = getjson(resultxml)
		if instr(resultxml,"error")<>0 then
		  if isobject(obj) Then
		      ks.echo "<h3>error:</h3>" & obj.error
			  ks.echo "<h3>error_code:</h3>" & obj.error_code
		  else
		    ks.die "授权失效！"
		  end if
		else
			if not ks.isnul(resultxml) then
				if isobject(obj) Then
				  nickname=obj.screen_name
				  figureurl=obj.profile_image_url
				 sex=obj.gender
				 if sex="m" then 
				   sex="男" 
				 elseif sex="f" then
				   sex="女"
				 else
				   sex="未如"
				 end if
				End If
				set obj=nothing
			end if
	
			set rs=conn.execute("select top 1 * from ks_user where sinaid='" & ks.delsql(ks.c("sinaid")) & "'")
			if rs.eof and rs.bof then
			 if ks.c("username")<>"" and ks.c("password")<>"" then '如果当前会员是登录状态的，直接绑定
				 Conn.Execute("Update KS_User Set sinatoken='" &ks.c("sina_access_token") & "',sinaid='" & KS.C("sinaid") &"' where username='" & KS.DelSQL(ks.c("username")) & "'")
				 Session(KS.SiteSN&"UserInfo")=""
				 Response.Redirect("../../user/user_bind.asp")
			 else
			     Call DoBind("用新浪微博账号登录成功",nickname,figureurl,sex,ks.c("sinaid"))
			 end if
		    else
			     Conn.Execute("Update KS_User Set SinaToken='" & KS.C("sina_access_token") &"' WHERE UserName='" & rs("username") &"'")
				 Call DoLogin(rs("username"),rs("password"))
			end if
			set rs=nothing
	  end if
	'=============================================================
End If
set ks=nothing
closeconn
%>
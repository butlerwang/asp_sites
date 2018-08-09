<!--#include file="../../plus/md5.asp"-->
<!--#include file="config.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
dim openid,username,password,action,rs
if ks.isnul(KS.C("openid")) then ks.die "没有返回openid!"
action=KS.S("Action")
If Action="check" Then
	openid=ks.s("openid")
	username=ks.r(ks.s("username"))
	password=md5(KS.R(ks.s("password")),16)
	if ks.c("openid")<>openid then
	  ks.die "请不要非法绑定!"
	end if
	set rs=server.createobject("adodb.recordset")
	rs.open "select top 1 * from ks_user where username='" & username & "' and password='" & password & "'",conn,1,1
	if rs.eof and rs.bof then
	  rs.close:set rs=nothing
	  ks.die "<script>alert('对不起，您输入的账号不存在或是密码不正确，请重输!');history.back(-1);</script>"
	else
		'绑定到已有账号
		conn.execute("update ks_user set qqtoken='" & KS.C("access_token") & "',qqopenid='" & openid & "' where username='" & username & "'")
		'调用登录
		Call DoLogin(username,password)
	end if
ElseIf Action="doreg" Then
        Call DoRegSave(1)
Else
    '===================绑定处理=================================
		dim ret,msg,nickname,figureurl,sex
		dim resultxml:resultxml=get_user_info(1,ks.c("access_token"),KS.C("openid"))
		dim obj:set obj = getjson(resultxml)
		if instr(resultxml,"access token check failed")<>0 then
		   if isobject(obj) Then
		    ks.die obj.msg
		   else
		    ks.die "error!"
		   end if
		Else
			if isobject(obj) Then
			  nickname=obj.nickname
			  figureurl=obj.figureurl
			  sex=obj.gender
			End If
			set obj=nothing
		end if
		if ks.chkclng(ret)<0 then    '获取失败
			 ks.die "登录失败:" & msg
		end if
		set rs=conn.execute("select top 1 * from ks_user where qqopenid='" & ks.delsql(ks.c("openid")) & "'")
	    if rs.eof and rs.bof then
		  if ks.c("username")<>"" and ks.c("password")<>"" then '如果当前会员是登录状态的，直接绑定
			 Conn.Execute("Update KS_User Set qqtoken='" & KS.C("access_token") & "',qqopenid='" & ks.c("openid") & "' where username='" & KS.DelSQL(ks.c("username")) & "'")
			 Session(KS.SiteSN&"UserInfo")=""
			 Response.Redirect("../../user/user_bind.asp")
		  else
		   Call DoBind("用QQ登录成功",nickname,figureurl,sex,ks.c("openid"))
		  end if
		Else
		     Conn.Execute("Update KS_User Set qqToken='" & KS.C("access_token") &"' WHERE UserName='" & rs("username") &"'")
			 Call DoLogin(rs("username"),rs("password"))
		end if
		set rs=nothing
	'=============================================================
End If
set ks=nothing
closeconn
%>
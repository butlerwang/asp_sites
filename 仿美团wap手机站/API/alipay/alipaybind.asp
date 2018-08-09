<!--#include file="../../conn.asp"-->
<!--#include file="../../ks_cls/kesion.commoncls.asp"-->
<!--#include file="../../plus/md5.asp"-->
<!--#include file="../cls_api.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
Dim Action,token,UserName,PassWord,user_id
action=KS.S("Action")
token=ks.s("token")
user_id=session("user_id")
if session("token")="" or session("user_id")="" then
	  ks.die "请不要非法绑定!"
end if
If Action="check" Then
	username=ks.r(ks.s("username"))
	password=md5(ks.r(ks.s("password")),16)
	set rs=server.createobject("adodb.recordset")
	rs.open "select top 1 * from ks_user where username='" & username & "' and password='" & password & "'",conn,1,1
	if rs.eof and rs.bof then
	  rs.close:set rs=nothing
	  ks.die "<script>alert('对不起，您输入的账号不存在或是密码不正确，请重输!');history.back(-1);</script>"
	else
		'绑定到已有账号
		conn.execute("update ks_user set alipayID='" & session("user_id") & "' where username='" & username & "'")
		'调用登录
		Call DoLogin(username,password)
	end if
ElseIf Action="doreg" Then
    Call DoRegSave(3)
Else
		set rs=conn.execute("select top 1 * from ks_user where alipayid='" & ks.delsql(session("user_id")) & "'")
	    if rs.eof and rs.bof then
		 if ks.c("username")<>"" and ks.c("password")<>"" then '如果当前会员是登录状态的，直接绑定
			 Conn.Execute("Update KS_User Set alipayid='" & ks.delsql(session("user_id")) & "' where username='" & KS.DelSQL(ks.c("username")) & "'")
			 Session(KS.SiteSN&"UserInfo")=""
			 Response.Redirect("../../user/user_bind.asp")
		 else
		   Call DoBind("用支付宝快捷登录成功",session("real_name"),KS.Setting(3)&"user/images/noavatar_small.gif","男",session("user_id"))
		 end if
		else
		   Call DoLogin(rs("username"),rs("password"))
		end if
		set rs=nothing

End If
Set KS=Nothing
CloseConn


%>
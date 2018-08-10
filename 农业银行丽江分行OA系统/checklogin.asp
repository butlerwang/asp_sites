<!--#INCLUDE FILE="data.asp" -->

<%
if request("username")="" and request("password")="" then
	Session("Ulogin")="no"
	Response.Redirect("login.asp?id=error")
else

	Uname=request("username")
	Upass=request("password")
	IP= Request.ServerVariables("REMOTE_ADDR")
    nowtime=now()
    sj=cstr(year(nowtime))+"-"+cstr(month(nowtime))+"-"+cstr(day(nowtime))+" "+cstr(hour(nowtime))+":"+right("0"+cstr(minute(nowtime)),2)+":"+right("0"+cstr(second(nowtime)),2)
	shijian=cstr(year(nowtime))+right("0"+cstr(month(nowtime)),2)+right("0"+cstr(day(nowtime)),2)+right("0"+cstr(hour(nowtime)),2)+right("0"+cstr(minute(nowtime)),2)
  
	Set rs= Server.CreateObject("ADODB.Recordset") 
	strSql="select * from user where 用户名='"&Uname&"' and 密码='"&Upass&"'"
	rs.open strSql,Conn,1,3 
	if rs.eof then
		Session("Ulogin")="no"
		response.redirect "login.asp?id=error"
		
    else  
		if rs("审核")=false then
			response.redirect "login.asp?id=pass"
		end if
		'response.write rs("登陆时间")
		'response.end
		rs("状态")=true
		rs("登陆IP")=IP
		rs("Utime")=shijian
		rs("times")=rs("times")+1
		
		rs.update
		Session("Uid")=rs("id")
		Session("Uname")=rs("用户名")
		Session("Rname")=rs("姓名")
		Session("Upass")=rs("密码")
		Session("Upart")=rs("部门")
		Session("Urule")=rs("权限")
		Session("tel")=rs("电话")
		Session("Utime")=rs("Utime")
		Session("IP")=rs("登录IP")
		Session("Ulogin")="yes"
		Session("email")=rs("信箱")
		session("mobile")=rs("mobile")
		session("time")=rs("时间")
		
		'----------------邮箱系统专用环境变量,请勿删除-------------------
			Session("id")="" & rs("用户名")
			Session("pwd")="" & rs("密码")
			Session("level")= "" & rs("ilevel")
			Session("iPageSize")=rs("iPageSize")
			Session("iAdd")="" & rs("iAdd")
			Session("iBegin")="" & rs("iBegin")
			Session("num")=0
		'----------------邮箱系统专用环境变量,请勿删除-------------------
		response.redirect("main.asp")
	%>


<script language="JavaScript">
<!--
function tmt_fullscreen(url, scrollo) {
    var larg_schermo = screen.availWidth - 10;
    var altez_schermo = screen.availHeight - 75;
    window.open(url, "", "width=" + larg_schermo + ",height=" + altez_schermo + ",top=0,left=0,menubar=yes,scrollbars=yes" );
}
//tmt_fullscreen("main.asp");


// -->

<%
	end if
end if
%>
<% rs.Close %>
<% Conn.Close %>
<% set Conn = nothing 
%>
</script>	

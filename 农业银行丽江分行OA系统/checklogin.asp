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
	strSql="select * from user where �û���='"&Uname&"' and ����='"&Upass&"'"
	rs.open strSql,Conn,1,3 
	if rs.eof then
		Session("Ulogin")="no"
		response.redirect "login.asp?id=error"
		
    else  
		if rs("���")=false then
			response.redirect "login.asp?id=pass"
		end if
		'response.write rs("��½ʱ��")
		'response.end
		rs("״̬")=true
		rs("��½IP")=IP
		rs("Utime")=shijian
		rs("times")=rs("times")+1
		
		rs.update
		Session("Uid")=rs("id")
		Session("Uname")=rs("�û���")
		Session("Rname")=rs("����")
		Session("Upass")=rs("����")
		Session("Upart")=rs("����")
		Session("Urule")=rs("Ȩ��")
		Session("tel")=rs("�绰")
		Session("Utime")=rs("Utime")
		Session("IP")=rs("��¼IP")
		Session("Ulogin")="yes"
		Session("email")=rs("����")
		session("mobile")=rs("mobile")
		session("time")=rs("ʱ��")
		
		'----------------����ϵͳר�û�������,����ɾ��-------------------
			Session("id")="" & rs("�û���")
			Session("pwd")="" & rs("����")
			Session("level")= "" & rs("ilevel")
			Session("iPageSize")=rs("iPageSize")
			Session("iAdd")="" & rs("iAdd")
			Session("iBegin")="" & rs("iBegin")
			Session("num")=0
		'----------------����ϵͳר�û�������,����ɾ��-------------------
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

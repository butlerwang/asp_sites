<!--#INCLUDE FILE="data.asp" -->
<!--#INCLUDE FILE="check.asp" -->
<%
for i=1 to 31
check=check&request("c"&i)&","
next
Set rs= Server.CreateObject("ADODB.Recordset") 
strSql="select * from userinfo where userid="&session("Uid")
rs.open strSql,Conn,1,3 
if rs.eof then
response.write "no record"
end if

rs("Uname")=request("Uname")
rs("sex")=request("sex")
rs("nation")=request("nation")
rs("duty")=request("duty")
rs("grade")=request("grade")
rs("birthday")=request("birthday")
rs("polity")=request("polity")
rs("health")=request("health")
rs("Nplace")=request("Nplace")
rs("weight")=request("weight")
rs("idcard")=request("idcard")
rs("height")=request("height")
rs("marriage")=request("marriage")
rs("Fschool")=request("Fschool")
rs("member")=request("member")
rs("speciality")=request("speciality")
rs("length")=request("length")
rs("study")=request("study")
rs("foreign")=request("foreign")
rs("Elevel")=request("Elevel")
rs("Clevel")=request("Clevel")
rs("Hplace")=request("Hplace")
rs("QQ")=request("QQ")
rs("call")=request("call")
rs("place")=request("place")
rs("love")=request("love")
rs("award")=request("award")
rs("experience")=request("experience")
rs("family")=request("family")
rs("contact")=request("contact")
rs("remark")=request("remark")
rs("check")=left(check,len(check)-1)


rs("Ltime")=now()
rs.update 
rs.close 
set rs=nothing 
set conn=nothing 
%>
<br><br>
<center><font color=red size=3>成功修改个人基本档案！</font><br><form method="post" action="archives.asp"><input type="submit" value="返回"></form>
</center>

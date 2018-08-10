<!--#INCLUDE FILE="data.asp" -->
<!--#include file="html.asp"-->

<%if Session("Urule")<>"a" then
	Response.write "你没有足够权限"
	response.end
end if
myUid=Session("Uid")
myUname=Session("Uname")
myUpass=Session("Upass")
myUrealname=Session("Rname")
myUpart=Session("Upart")
myUrule=Session("Urule")
myUlogin=Session("Ulogin")
 nowtime=now()
sj=cstr(year(nowtime))+"-"+cstr(month(nowtime))+"-"+cstr(day(nowtime))+" "+cstr(hour(nowtime))+":"+right("0"+cstr(minute(nowtime)),2)+":"+right("0"+cstr(second(nowtime)),2)

Set my_rs= Server.CreateObject("ADODB.Recordset") 
StrSQL = "Select * FROM jhtdata"
my_rs.Open StrSQL,Conn,1,3
my_rs.Addnew
	my_rs("inid") = myUid
	my_rs("outid") = 0
	my_rs("真实姓名") = myUrealname
	my_rs("部门") = myUpart
	my_rs("标题") = htmlencode2(request("biaoti"))
	my_rs("内容") = htmlencode2(request("neirong"))
	my_rs("时间") =sj
	my_rs.Update
%>
<script language=javascript>
opener.location=opener.location;
</script>

<title>您的信息已经成功添加</title>
<LINK href="oa.css" rel=stylesheet>
<body style="background-image: url('images/main_bg.gif'); background-attachment: scroll; background-repeat: no-repeat; background-position: left bottom">

<div align="center">
  <center>
  <table border="1" width="100%" cellspacing="0" cellpadding=0 height="209" bordercolorlight=#000000 bordercolordark=#ffffff>
    <tr>
       <td align=center colspan=2>您的信息已经成功添加
	   </td>
	</tr>
	<tr valign=top>
      <td width="21%" align="center" height="14">
        主　　题</td>
      <td width="79%" height="14"><%=my_rs("标题")%></td>
    </tr>
    <tr>
      <td width="21%" align="center" height="18">
        发 布 人</td> 
      <td width="79%" height="18"><%=my_rs("部门")%>　<%=my_rs("真实姓名")%></td>
    </tr>
    <tr>
      <td width="21%" align="center" height="17">
        发布时间</td>
      <td width="79%" height="17"><%=my_rs("时间")%></td>
    </tr>
    <tr valign=top>
      <td width="21%" align="center" height="122">
        内　　容</td>
      <td width="79%" height="200"><%=my_rs("内容")%></td>
    </tr>
  </table>
  </center><BR>
 <CENTER><a href="Javascript:window.close();"><img name="Image3" border="0" src="images/close_1.gif" width="85" height="19" hspace="5"></a></CENTER> 
</div>
<script language=javascript>
opener.location=opener.location;
</script>

</body>

<%
	my_rs.close
	Set my_rs=nothing
	my_Conn.Close
	Set my_Conn=nothing

%>

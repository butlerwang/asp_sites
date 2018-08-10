<!--#include file="../check.asp"-->
<!--#include file="../include/conn.asp"-->
<!--#include file="../include/lib.asp"-->
<%openData()
If Session("name") = "" then
response.Write "<script LANGUAGE=javascript>alert('网络超时，或者你还没有登陆! ');this.location.href='../login.asp';</script>"
end if
  IF instr(webConfig,", 6")=0 or instr(session("manconfig"),", 6")=0 Then'网站功能配置
Session.Abandon()
	Response.Write "<Script Language=JavaScript>alert('您的权限不足，请不要非法调用其它管理模块，否则您的账号将被系统自动删除!');this.location.href='../login.asp';</Script>"
Response.end
end if%>
<%Dim Act
  Act=Request.Form("act")
  OpenData()
  Call Main()  
  Call CloseDataBase()
  Sub Main()
  ID=Cint(Request.QueryString("iD"))
  Set Rs=Server.CreateObject("adodb.recordset")
  Sql="Select Job,RealName,Sex,Birthday,Address,School,Education,Profession,working,Tel,Email,Content,fuqin,muqin,jk,aihao,hj,bysj From Sbe_Resume Where id="&ID
  Rs.Open SQL,Conn,1,1
 ' set rs1=server.CreateObject("adodb.recordset")
'rs1.open "select * from [Sbe_Resume] where id="&ID ,conn,1,3
'rs1("Ability")=2
'rs1.update
'rs1.close
'set rs1=nothing
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>后台管理系统</title>
<link href="../include/style.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
function check(){
  if(form1.Job.value==""){
     alert("请选填写岗位名称！");
	 form1.Job.focus();
	 return false;
  }
  return true;  
}
</script>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="25"><font color="#6A859D"> 在线招聘 &gt;&gt; 求职信息详情</font></td>
  </tr>
  <tr> 
    <td height="1" background="../images/dot.gif"></td>
  </tr>
</table>
  
<br>
<table width="71%" border="0" align="center" cellpadding="3" cellspacing="0" id="sbe_table">
  <form name="form1" method="post" action="" onSubmit="return check()">
    <tr > 
      <td height="25" colspan="2"  class="sbe_table_title">求职信息详情</td>
    </tr>
    <tr > 
      <td width="106" height="25" align="right" class="szise3">姓名：</td>
      <td>&nbsp;<%=rs(1)%></td>
    </tr>
            <tr>
              <td height="25" align="right" class="szise3">应聘岗位：</td>
              <td width="588" style="padding-left:4px">&nbsp;<% sql1="select Job from Sbe_Job  where ID="&rs(0)
	  set rs1=conn.execute(sql1)
	  if not rs.eof then
	    response.Write rs1(0)
	   end if
	  rs1.close
	  set rs1=nothing%></td>
            </tr>
            <tr>
              <td height="25" align="right" class="szise3">出生年月：</td>
              <td>&nbsp;<%=rs(3)%></td>
            </tr>
            <tr>
              <td height="25" align="right" class="szise3">所学专业：</td>
              <td>&nbsp;<%=rs(7)%></td>
            </tr>
            <tr>
              <td height="25" align="right" class="szise3">毕业时间：</td>
              <td>&nbsp;<%=rs(17)%></td>
    </tr>
            <tr>
              <td height="25" align="right" class="szise3">学位：</td>
              <td>&nbsp;<%=rs(6)%></td>
            </tr>
            <tr>
              <td height="25" align="right" class="szise3">获奖情况：</td>
              <td>&nbsp;<%=rs(16)%></td>
            </tr>
            
            <tr>
              <td height="25" align="right" class="szise3">兴趣爱好：</td>
              <td>&nbsp;<%=rs(15)%></td>
            </tr>
            
            <tr>
              <td height="25" align="right" class="szise3">身体状况：</td>
              <td>&nbsp;<%=rs(14)%></td>
    </tr>
            <tr>
              <td height="25" align="right" class="szise3">父亲职业：</td>
              <td>&nbsp;<%=rs(12)%></td>
            </tr>
            <tr>
              <td height="25" align="right" class="szise3">母亲职业：</td>
              <td>&nbsp;<%=rs(13)%></td>
            </tr>
            
            
            <tr>
              <td height="76" align="right" class="szise3">工作经历：</td>
              <td valign="middle">&nbsp;<%=HTMLcode(rs(11))%></td>
    </tr>
  </form>
</table>
<div align="center"><br>
  <input type="button" name="Submit" onClick="history.go(-1);return false;" value="返回" />
  <br>
</div>
</body>
</html>
<%Rs.Close
Set rS=Nothing
  End Sub%>

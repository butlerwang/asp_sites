<!--#include file="conn.asp"-->
<!--#include file="include/lib.asp"-->
<!--#include file="include/md5.asp"-->
<%
IF IsSubmit Then
  strUserName=trim(FilterSQL(request.form("username")))
  strUserPassWord=trim(FilterSQL(request.form("password")))
  verifycode=replace(trim(request("verifycode")),"'","")
  
  if strUserName="" or strUserPassWord="" then
    response.Write "<script LANGUAGE='javascript'>alert('您的管理帐号或密码为空！');history.go(-1);</script>"
    response.end
  end if
  if yanzhengma=true then
  if cstr(session("getcode"))<>cstr(trim(request("verifycode"))) then
    response.Write "<script LANGUAGE='javascript'>alert('请输入正确的验证码！');history.go(-1);</script>"
    response.end
  end if
  end if
  Call OpenData()
  IF ChkLogin(strUserName,strUserPassWord) = 1 Then
     Response.Write("<script>alert('登陆成功');this.location.href='index.asp';</script>") 
  Else
    Response.Write("<script>alert('帐号错误或您的帐号被锁定');history.back();</script>")   
  End IF
  Call CloseDataBase()  
End IF
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>管理员登陆</title>
<link href="include/style.css" rel="stylesheet" type="text/css">

<style type="text/css">
<!--
body {
	background-image: url(images/fstbg.gif);
}
-->
</style></head>

<body>
<form name="form1" method="post">
<table width="320"  border="0" align="center" cellpadding="0" cellspacing="0" style="margin-top:250px; ">
  <tr>
    <td ><table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td align="center"><img src="images/03.gif" width="99" height="27"></td>
      </tr>
      <tr>
        <td height="30">&nbsp;用户名:
          <input name="username" type="text" class="input" id="username" style="height:20px;width:150px; "></td>
      </tr>
      <tr>
        <td height="30">&nbsp;密　码:          
          <input name="password" type="password" class="input" id="password" style="height:20px;width:150px; "></td>
      </tr>
	  <%if yanzhengma=true then%>
      <tr>
        <td>&nbsp;验证码:
          <input name="verifycode" type="text" class="input"  style="width:100px;" value="<%If GetCode=9999 Then Response.Write "9999"%>">
          &nbsp;<img src=GetCode.asp></td>
      </tr>
	  <%end if%>
      <tr>
        <td>&nbsp;<input name="imageField" type="image" src="images/login_button.jpg" width="90" height="26" border="0"></td>
      </tr>
    </table></td>
  </tr>
</table>
</form>
</body>
</html>
<%
Private Function ChkLogin(Byval strUserName,Byval strUserPassWord)
'函数名称: FilterSQL
'功能描述: 检测用户登陆
'使用方法：返回 -1 则用户名不存在 0 则密码错误 1 登陆成功
	Dim strSQL
	Dim ObjRs
	StrSQL="Select * From Sbe_Admin Where AdminName='" & FilterSQL(strUserName) &"' and Lock <> 1"	
	Set objRs=Conn.Execute(StrSQL)
	With ObjRs
		If objRs.Eof and objRs.bof Then
			ChkLogin=-1
		Elseif objRs.Fields(2).Value=md5(strUserPassWord) Then			
			ChkLogin=1
			session("flag")=objRs("Popedom")
            session("name")=objRs("AdminName")		
			session("manconfig")=objRs("template")	
			conn.execute("update Sbe_Admin set LoginTime ='"&date()&"' ,loginTimes="& objRs("loginTimes")+1 &" where ID= "& objRs("id") &" ")			
		Else			 
			ChkLogin=0
		End if
	End with
	ObjRs.Close : Set ObjRs=Nothing '数的建立	
End Function
%>

<!--#include file="../check.asp"-->
<!--#include file="../include/conn.asp"-->
<!--#include file="../include/lib.asp"-->
<!--#include file="../include/md5.asp"-->
 <%
If Session("name") = "" then
response.Write "<script LANGUAGE=javascript>alert('网络超时，或者你还没有登陆! ');this.location.href='../login.asp';</script>"
end if
 	temp_check_rights = Split(session("manconfig"),",")
	for temp_rights_count=0 to ubound(temp_check_rights)
	    if trim(temp_check_rights(temp_rights_count)) = "5" then
			rights_check_passkey = trim(temp_check_rights(temp_rights_count))
		end if
	next
	if rights_check_passkey <> "5" then
Session.Abandon()
	Response.Write "<Script Language=JavaScript>alert('您的权限不足，请不要非法调用其它管理模块，否则您的账号将被系统自动删除!');this.location.href='../login.asp';</Script>"
	Response.end
	end if%>
<%
Call OpenData()
memberID = Trim(Request("ID"))
If IsSubmit then
  Dim msg  
  Set rs=server.createobject("adodb.recordset")
  If memberID = "" Then
    sqla="select AdminName from Sbe_Admin where AdminName ='"&request("username")&"'"
	set rsa=conn.execute(sqla)
	if not rsa.eof then
	Response.Write "<Script>alert('数据库中已经存在同名的管理员');history.go(-1);</script>" 
	Response.End 
	end if
	rsa.close
	set rsa=nothing
	msg = "管理员帐号添加成功！"
	Rs.open "Select * from Sbe_Admin where id Is null",conn,1,3
	Rs.addnew
  Else
  if request("username") <> request("username1") then
    sqla="select AdminName from Sbe_Admin where AdminName ='"&request("username")&"'"
	set rsa=conn.execute(sqla)
	if not rsa.eof then
	Response.Write "<Script>alert('数据库中已经存在同名的管理员');history.go(-1);</script>" 
	Response.End 
	end if
	rsa.close
	set rsa=nothing
  end if
   msg = "管理员帐号修改成功！"
	Rs.open "Select * from Sbe_Admin where ID=" & clng(memberID) ,conn,1,3
  End if
  Rs("AdminName")= trim(Request.Form("username"))
   if trim(Request.Form("password")) <> trim(Request.Form("password2")) then
  Rs("PassWord")=md5(trim(request.Form("PassWord")))
  end if
  Rs("note")= Request.Form("note")
  Rs("Popedom")= 1
  Rs("RegTime")= date()
  Rs("Lock")=clng(request.Form("Lock"))
  if trim(Request.Form("checkbox"))<>"" then
     Rs("template")= "0, "&trim(Request.Form("checkbox")) 
  else
     Rs("template")= "0"  
  end if 
  rs.update
  rs.close
  Set rs=nothing
  if trim(request("username1")) = session("name") then
    if trim(request("username1")) <> trim(request("username")) then
	  'session("name")=""
	 ' session("flag")="	"
	 ' session("manconfig")=""
	 Session.Abandon()
	 Call MessageBoxOKa(msg) '完成提示
	' response.End
   end if
   end if	
	Call MessageBoxOK(msg) '完成提示
ElseIF Len(memberID)>0 Then	
	Dim StrSQL
	Dim objRec
	StrSQL = "Select * from Sbe_Admin Where ID=" & memberID
	Set objRec = Conn.Execute(StrSQL)
	With ObjRec
		If .Eof And .Bof Then
			Response.Write "<Script>alert('操作失败');history.back();</script>" 
			Response.End
		Else
			username = objRec("AdminName")
			Lock = objRec("Lock")	
			password = objRec("PassWord")
			note = objRec("note")
			checkstr = objRec("template")
		End If
	End With
	objRec.Close:set objRec=Nothing
End if
Private Sub MessageBoxOK(strValue)
	With Response
		.Write "<script>" & vbcrlf
		.Write "alert('"+strValue+"');" & vbcrlf
		.Write "this.location.href='main.asp'" & vbcrlf
		.Write "</script>" & vbcrlf
	End With
End Sub
Private Sub MessageBoxOKa(strValue)
	With Response
		.Write "<script>" & vbcrlf
		.Write "alert('"+strValue+"');" & vbcrlf
		.Write "this.location.href='../login.asp'" & vbcrlf 
		.Write "</script>" & vbcrlf
	End With
End Sub
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>后台管理系统</title>
<link href="../include/style.css" rel="stylesheet" type="text/css">
<script language="javascript">
  function check_admin(){
    username=form1.username.value;
	password=form1.password.value;
	password1=form1.password1.value;
	if(username==''){
	  alert('请填写用户名');
	  form1.username.focus();
	  return false;
	}
	if(password==''){
	  alert('请填写密码');
	  form1.password.focus();
	  return false;
	}
	if(password1==''){
	  alert('请填写确认密码');
	  form1.password1.focus();
	  return false;
	}
	if(password!=password1){
	 alert('密码不一致!');	  
	  return false;
	}	
  } 
</script>
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="25"><font color="#6A859D">会员管理&gt;&gt; 管理员帐号管理</font></td>
  </tr>
  <tr>
    <td height="1" background="../images/dot.gif"></td>
  </tr>
</table>
<br>
<br>
<table id="sbe_table" width="95%" align="center">
  <tr> 
    <td height="22" colspan="8" class="sbe_table_title">管理员帐号</td>
  </tr>
  <tr align="center" class="font_bold"> 
    <td width="12%">用户名</td>
    <td width="31%">管理权限</td>
    <td width="15%">添加时间</td>
    <td width="18%">最后登陆时间</td>
    <td width="9%" >登陆次数</td>
	 <td width="7%">锁定</td>
    <td width="8%" colspan="2">操作</td>	
  </tr>
  <%Call member_list()%>
</table>
<br><br><br><br>
<form name="form1" method="post" onSubmit="return check_admin()">
  <table width="70%"  border="0" align="center" cellpadding="0" cellspacing="0" id="sbe_table">
    <tr>
      <td colspan="4" class="sbe_table_title">系统帐号管理</td>
    </tr>
    <tr>
      <td width="20%" align="right" bgcolor="#E9EFF3">用户名:</td>
      <td width="32%">&nbsp;<input name="username" type="text" class="input_length" id="username" value="<%=username%>"><input name="username1" type="hidden" class="input_length" id="username1" value="<%=username%>"></td>
      <td width="16%" align="right" bgcolor="#E9EFF3">锁定:</td>
      <td width="32%"><input name="Lock" type="checkbox" id="Lock" value="1" <%if lock = true then%>checked<%end if%>></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#E9EFF3">密码:</td>
      <td>&nbsp;<input name="password" type="password" class="input_length" id="password" value="<%=password%>" maxlength="15"><input name="password2" type="hidden" class="input_length" id="password2" value="<%=password%>"></td>
      <td align="right" bgcolor="#E9EFF3">确认密码:</td>
      <td>&nbsp;<input name="password1" type="password" class="input_length" id="password1" value="<%=password%>" maxlength="15"></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#E9EFF3">管理权限:</td>
      <td colspan="3">&nbsp;
	  <%
	  IF checkstr="" then
	   Call check_name_str(checkstr) 
	  Else
	   Call admin_select(checkstr)
	  End IF
	  %></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#E9EFF3">备注:</td>
      <td colspan="3">&nbsp;<textarea name="note" class="input_length" id="note" style="width:430px;height:70px;"><%=note%></textarea></td>
    </tr>
    <tr align="center">
      <td height="30" colspan="4" class="font_bold"><input name="Submit" type="submit" class="sbe_button" value="提交">
      <input name="Submit2" type="reset" class="sbe_button" value="重置"></td>
    </tr>
  </table>
  <input type="hidden" name="ID" value="<%=memberID%>">
</form>
<%Call CloseDataBase()%>
</body>
</html>
<%
Private Sub member_list()
'管理员列表
 Set Rs=Conn.Execute("select * from Sbe_Admin where Popedom =1 order by id desc")
 With Response
  Do While not Rs.eof 
   .write "<tr align=""center"" onMouseOver=this.style.backgroundColor='CDE9F8' onMouseOut=this.style.backgroundColor='#FFFFFF'>"
   .write "    <td>"& Rs("AdminName") &"</td>"
   .write "    <td align=""left"">"
   Call check_name_str(Rs("template")) 
   .Write"</td>"
   .write "    <td>"& Rs("RegTime") &"</td>"
   .write "    <td align=""left"">"& Rs("LoginTime") &"</td>"
   .write "    <td>"& Rs("loginTimes") &"</td>"
   if Rs("lock")=true then
     .write "    <td  title=""管理帐号已锁定""> <b><font color=#009900>×</font></b></td>"
   else
     .write "    <td  title=""管理帐号未锁定""> <b><b><font color=#FF0000>√</font></b></td>"
   end if
   
   .write "    <td width=""6%"">"
   'IF session("flag")=0  Then
      .write"<a href=?id="& Rs("id") &"><img src=""../images/edit.gif"" width=""14"" height=""15"" border=""0""></a> "   
 '  Else
     '  .write"<img src=""../images/edit.gif"" width=""14"" height=""15"" border=""0"">"
  ' End IF
   .write"</td>"
   .write "    <td width=""7%"">"
  ' IF session("flag")=0 Then
   .write"<a href=""del.asp?Table_name=Sbe_Admin&ItemID=id&Id="& Rs("id") &""" onClick=""javascript:return confirm('\n确定删除吗？')""><img src=""../images/delete.gif"" width=""10"" height=""13"" border=""0""></a>"
 '  Else
  ' .write"<img src=""../images/delete.gif"" width=""10"" height=""13"" border=""0"">"
  ' End IF
   .write"</td>"
   .write "  </tr>"

  Rs.Movenext
  loop
 End With
 Rs.Close:Set Rs=Nothing
End Sub
Private Sub check_name_str(strID) 
 strID=strID
 If strID="" or isnull(strID) Then
  Set oRs=Conn.Execute("select Template from Sbe_WebConfig")
  IF not(oRs.eof and oRs.bof) Then 
   strID=oRs.Fields(0).value
   oRs.close:set oRs=Nothing
  Else
    Exit Sub
  End IF
 Else
   strID=strID
 End IF 
 
 arry=split(strID,",")
   for i=0 to ubound(arry)
     Call check_name(arry(i))
   next 
End Sub

Private Function isIn(intID,strID)
  intID=trim(intID)
  strID=trim(strID)
  IF InStr(intID,strID)>0 Then isIn="checked"  
End Function

Private Sub admin_select(str)
  str1=str  
 Set objRec=Conn.Execute("select Template from Sbe_WebConfig")
  IF not(objRec.eof and objRec.bof) Then 
   str2=objRec.Fields(0).value  
  Else
    Exit Sub
  End IF
  arry1=split(str1,",")
  j=ubound(arry1)
  arry=split(str2,",")
   for i=0 to ubound(arry)     
      Call check_name_no(arry(i),str1)	
   next 
   
End Sub
Private Sub check_name_no(intID,valueID)
 intID=intID
 valueID=valueID
  select Case intID
   Case 0	 
   Case 1
     str="<input name=""checkbox"" type=""checkbox"" value=""1"" "& isIn(valueID,1) &">企业信息"
   Case 2
     str="<input name=""checkbox"" type=""checkbox"" value=""2"" "& isIn(valueID,2) &">客房中心"
   Case 3
     str="<input name=""checkbox"" type=""checkbox"" value=""3"" "& isIn(valueID,3) &">资讯中心"
   Case 4
     str="<input name=""checkbox"" type=""checkbox"" value=""4"" "& isIn(valueID,4) &">店铺形象"
   Case 5
     str="<input name=""checkbox"" type=""checkbox"" value=""5"" "& isIn(valueID,5) &">权限管理"
  Case 6
     str="<input name=""checkbox"" type=""checkbox"" value=""6"" "& isIn(valueID,6) &">人事招聘"
  Case 7
     str="<input name=""checkbox"" type=""checkbox"" value=""7"" "& isIn(valueID,7) &">在线留言"
  Case 8
     str="<input name=""checkbox"" type=""checkbox"" value=""8"" "& isIn(valueID,8) &">在线预定"
  Case 9
     str="<input name=""checkbox"" type=""checkbox"" value=""9"" "& isIn(valueID,9) &">楼盘标志"
  end select
   Response.Write str
End Sub
%>
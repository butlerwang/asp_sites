<!--#include file="../check.asp"-->
<!--#include file="../include/conn.asp"-->
<!--#include file="../include/lib.asp"-->
 <!--#include file="../include/md5.asp"--> 
<%
If Session("name") = "" then
response.Write "<script LANGUAGE=javascript>alert('���糬ʱ�������㻹û�е�½! ');window.location.href='login.asp';</script>"
response.End
end if
Call OpenData()
memberID = Trim(Request("ID"))
If IsSubmit then
  Dim msg  
  Set rs=server.createobject("adodb.recordset")
  if request("username") <> request("username1") then
    sqla="select AdminName from Sbe_Admin where AdminName ='"&request("username")&"'"
	set rsa=conn.execute(sqla)
	if not rsa.eof then
	Response.Write "<Script>alert('���ݿ����Ѿ�����ͬ���Ĺ���Ա');history.go(-1);</script>" 
	Response.End 
	end if
	rsa.close
	set rsa=nothing
  end if
  msg = "����Ա�ʺ��޸ĳɹ���"
  Rs.open "Select * from Sbe_Admin where ID=" & clng(memberID) ,conn,1,3
  Rs("AdminName")= Request.Form("username")
   if trim(Request.Form("password")) <> trim(Request.Form("password2")) then
  Rs("PassWord")=md5(trim(request.Form("PassWord")))
  end if
  Rs("note")= Request.Form("note")
  rs.update
  rs.close
  Set rs=nothing
  if session("name") <> Request.Form("username") then
    ' session("name")=""
	 'session("flag")="	"
	 'session("manconfig")=""
	 Session.Abandon()
 Call MessageBoxOKa(msg) '�����ʾ	
	 else
  Call MessageBoxOK(msg) '�����ʾ
  end if
end if
Private Sub MessageBoxOK(strValue)
	With Response
		.Write "<script>" & vbcrlf
		.Write "alert('"+strValue+"');" & vbcrlf
		.Write "this.location.href='../main.asp'" & vbcrlf
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
<title>��̨����ϵͳ</title>
<link href="../include/style.css" rel="stylesheet" type="text/css">
<script language="javascript">
  function check_admin(){
    username=form1.username.value;
	password=form1.password.value;
	password1=form1.password1.value;
	if(username==''){
	  alert('����д�û���');
	  form1.password.focus();
	  return false;
	}
	if(password==''){
	  alert('����д����');
	  form1.password.focus();
	  return false;
	}
	if(password1==''){
	  alert('����дȷ������');
	  form1.password1.focus();
	  return false;
	}
	if(password!=password1){
	 alert('���벻һ��!');	  
	  return false;
	}
	
  }
  
</script>
</head>
<body>
<%Sql="select * from Sbe_Admin where AdminName ='"&session("name")&"'"
set rs=conn.execute(Sql)
if not rs.eof then
username =rs("AdminName")
memberID=rs("ID")
PassWord =rs("PassWord")
note =rs("note")
end if
rs.close
set rs=nothing
%>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="25"><font color="#6A859D">��Ա����&gt;&gt; ����Ա�ʺŹ���</font></td>
  </tr>
  <tr>
    <td height="1" background="../images/dot.gif"></td>
  </tr>
</table>
<form name="form1" method="post" onSubmit="return check_admin()">
  <table width="70%"  border="0" align="center" cellpadding="0" cellspacing="0" id="sbe_table">
    <tr>
      <td colspan="4" class="sbe_table_title">�ʺŹ���</td>
    </tr>
    <tr>
      <td width="20%" align="right" bgcolor="#E9EFF3">�û���:</td>
      <td colspan="3">&nbsp;<input name="username" type="text" class="input_length" id="username" value="<%=username%>"><input name="username1" type="hidden" class="input_length" id="username1" value="<%=username%>"></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#E9EFF3">����:</td>
      <td width="32%">&nbsp;<input name="password" type="password" class="input_length" id="password" value="<%=password%>" maxlength="15"><input name="password2" type="hidden" class="input_length" id="password2" value="<%=password%>"></td>
      <td width="16%" align="right" bgcolor="#E9EFF3">ȷ������:</td>
      <td width="32%">&nbsp;<input name="password1" type="password" class="input_length" id="password1" value="<%=password%>" maxlength="15"></td>
    </tr>
    
    <tr>
      <td align="right" bgcolor="#E9EFF3">��ע:</td>
      <td colspan="3">&nbsp;<textarea name="note" class="input_length" id="note" style="width:430px;height:70px;"><%=note%></textarea></td>
    </tr>
    <tr align="center">
      <td height="30" colspan="4" class="font_bold"><input name="Submit" type="submit" class="sbe_button" value="�ύ">
      <input name="Submit2" type="reset" class="sbe_button" value="����"></td>
    </tr>
  </table>
  <input type="hidden" name="ID" value="<%=memberID%>">
</form>
<%Call CloseDataBase()%>
</body>
</html>
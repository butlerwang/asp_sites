<!--#include file="../check.asp"-->
<!--#include file="../include/conn.asp"-->
<!--#include file="../include/lib.asp"-->
 <%
If Session("name") = "" then
response.Write "<script LANGUAGE=javascript>alert('网络超时，或者你还没有登陆! ');this.location.href='../login.asp';</script>"
end if
 	temp_check_rights = Split(session("manconfig"),",")
	for temp_rights_count=0 to ubound(temp_check_rights)
	    if trim(temp_check_rights(temp_rights_count)) = "3" then
			rights_check_passkey = trim(temp_check_rights(temp_rights_count))
		end if
	next
	if rights_check_passkey <> "3" then
Session.Abandon()
	Response.Write "<Script Language=JavaScript>alert('您的权限不足，请不要非法调用其它管理模块，否则您的账号将被系统自动删除!');this.location.href='../login.asp';</Script>"
	Response.end
	end if%>
<%
Call OpenData()
 CompanyID = Trim(Request("ID"))
If IsSubmit then  
  Dim msg  
  Set rs=server.createobject("adodb.recordset")
  If CompanyID = "" Then
	msg = "资讯来源字段设置成功!"
	Rs.open "Select * from news_come_class where id Is null",conn,1,3
	Rs.addnew
  Else
	msg = "资讯来源字段设置修改成功！"
	Rs.open "Select * from news_come_class where ID=" & clng(CompanyID) ,conn,1,3
  End if
  Rs("title")= trim(Request.Form("title"))  
  rs.update
  rs.close
  Set rs=nothing	
	Call MessageBoxOK(msg) '完成提示

ElseIF Len(CompanyID)>0 Then	
	Dim StrSQL
	Dim objRec
	StrSQL = "Select * from news_come_class Where ID=" & CompanyID
	Set objRec = Conn.Execute(StrSQL)
	With ObjRec
		If .Eof And .Bof Then
			Response.Write "<Script>alert('操作失败');history.back();</script>" 
			Response.End
		Else
			title = objRec("title")			
		End If
	End With
	objRec.Close:set objRec=Nothing
End if
Private Sub MessageBoxOK(strValue)
	With Response
		.Write "<script>" & vbcrlf
		.Write "alert('"+strValue+"');" & vbcrlf
		.Write "this.location.href='news_come_class.asp'" & vbcrlf
		.Write "</script>" & vbcrlf
	End With
End Sub
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>资源来源类别管理</title>
<link href="../include/style.css" rel="stylesheet" type="text/css">
<script language="javascript">
  function check_admin(){
    title=form1.title.value;
	
	if(title==''){
	  alert('请填写分类名称!');
	  form1.title.focus();
	  return false;
	}	
  }
  
</script>
</head>

<body>
<br>
<br>
<table width="380" height="0" border="0" align="center" cellpadding="0" cellspacing="0" class="four" id="sbe_table">
  <tr align="center" bgcolor="#EEEEEE">
    <td height="30" colspan="4" class="sbe_table_title"><strong>已有资讯来源</strong></td>
  </tr>
  <tr align="center">
    <td width="70" height="30">ID</td>
    <td width="226" height="30">字段</td>
    <td width="74" height="30" colspan="2">操作</td>
  </tr>
 <%Call Company_list()%>
</table>
<br><br><br><br>
<form name="form1" method="post" onSubmit="return check_admin()">
<table width="360" border="0" align="center" cellpadding="0" cellspacing="0" class="four" id="sbe_table">
  <tr align="center" bgcolor="#EEEEEE">
    <td height="30" class="sbe_table_title">设置</td>
  </tr>
  <tr>
    <td height="30" align="center">类别:
      <input type="hidden" name="ID" value="<%=CompanyID%>"><input name="title" type="text" id="title" style="width:200px;height:22px;" value="<%=title%>">
    <input name="Submit" type="submit" class="SELECTsmallSel" value="提交" style="width:50px;height:22px;"></td>
  </tr>
</table>

</form>
<%Call CloseDataBase()%>
<br>
<div align="center">【<a href="javascript:window.close();">关闭</a>】</div>

<br><br>
</body>
</html>
<%
Private Sub Company_list()
'管理员列表
 Set Rs=Conn.Execute("select id,title from news_come_class order by id asc")
 With Response
  Do While not Rs.eof 
     .write " <tr align=""center"">"& vbCrLf
     .write "    <td height=""30"" class=""bottom"">"& Rs("ID") &"</td>"& vbCrLf
     .write "    <td height=""30""  class=""leftadnbottom1"">"& Rs("title") &"</td>"& vbCrLf
     .write "    <td width=""43"" height=""30"" align=""center"" class=""leftadnbottom1""><a href=?id="& Rs("id") &"><img src=""../images/edit.gif"" width=""14"" height=""15"" border=""0""></a></td>"& vbCrLf
     .write "    <td width=""43"" align=""center"" class=""leftadnbottom1""><a href=""del.asp?Table_name=news_come_class&ItemID=id&Id="& Rs("id") &""" onClick=""javascript:return confirm('\n确定删除吗？')""><img src=""../images/delete.gif"" width=""10"" height=""13"" border=""0""></a></td>"& vbCrLf
     .write "  </tr>"& vbCrLf     
  Rs.Movenext
  loop
 End With
 Rs.Close:Set Rs=Nothing
End Sub
%>
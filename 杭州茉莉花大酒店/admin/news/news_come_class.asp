<!--#include file="../check.asp"-->
<!--#include file="../include/conn.asp"-->
<!--#include file="../include/lib.asp"-->
 <%
If Session("name") = "" then
response.Write "<script LANGUAGE=javascript>alert('���糬ʱ�������㻹û�е�½! ');this.location.href='../login.asp';</script>"
end if
 	temp_check_rights = Split(session("manconfig"),",")
	for temp_rights_count=0 to ubound(temp_check_rights)
	    if trim(temp_check_rights(temp_rights_count)) = "3" then
			rights_check_passkey = trim(temp_check_rights(temp_rights_count))
		end if
	next
	if rights_check_passkey <> "3" then
Session.Abandon()
	Response.Write "<Script Language=JavaScript>alert('����Ȩ�޲��㣬�벻Ҫ�Ƿ�������������ģ�飬���������˺Ž���ϵͳ�Զ�ɾ��!');this.location.href='../login.asp';</Script>"
	Response.end
	end if%>
<%
Call OpenData()
 CompanyID = Trim(Request("ID"))
If IsSubmit then  
  Dim msg  
  Set rs=server.createobject("adodb.recordset")
  If CompanyID = "" Then
	msg = "��Ѷ��Դ�ֶ����óɹ�!"
	Rs.open "Select * from news_come_class where id Is null",conn,1,3
	Rs.addnew
  Else
	msg = "��Ѷ��Դ�ֶ������޸ĳɹ���"
	Rs.open "Select * from news_come_class where ID=" & clng(CompanyID) ,conn,1,3
  End if
  Rs("title")= trim(Request.Form("title"))  
  rs.update
  rs.close
  Set rs=nothing	
	Call MessageBoxOK(msg) '�����ʾ

ElseIF Len(CompanyID)>0 Then	
	Dim StrSQL
	Dim objRec
	StrSQL = "Select * from news_come_class Where ID=" & CompanyID
	Set objRec = Conn.Execute(StrSQL)
	With ObjRec
		If .Eof And .Bof Then
			Response.Write "<Script>alert('����ʧ��');history.back();</script>" 
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
<title>��Դ��Դ������</title>
<link href="../include/style.css" rel="stylesheet" type="text/css">
<script language="javascript">
  function check_admin(){
    title=form1.title.value;
	
	if(title==''){
	  alert('����д��������!');
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
    <td height="30" colspan="4" class="sbe_table_title"><strong>������Ѷ��Դ</strong></td>
  </tr>
  <tr align="center">
    <td width="70" height="30">ID</td>
    <td width="226" height="30">�ֶ�</td>
    <td width="74" height="30" colspan="2">����</td>
  </tr>
 <%Call Company_list()%>
</table>
<br><br><br><br>
<form name="form1" method="post" onSubmit="return check_admin()">
<table width="360" border="0" align="center" cellpadding="0" cellspacing="0" class="four" id="sbe_table">
  <tr align="center" bgcolor="#EEEEEE">
    <td height="30" class="sbe_table_title">����</td>
  </tr>
  <tr>
    <td height="30" align="center">���:
      <input type="hidden" name="ID" value="<%=CompanyID%>"><input name="title" type="text" id="title" style="width:200px;height:22px;" value="<%=title%>">
    <input name="Submit" type="submit" class="SELECTsmallSel" value="�ύ" style="width:50px;height:22px;"></td>
  </tr>
</table>

</form>
<%Call CloseDataBase()%>
<br>
<div align="center">��<a href="javascript:window.close();">�ر�</a>��</div>

<br><br>
</body>
</html>
<%
Private Sub Company_list()
'����Ա�б�
 Set Rs=Conn.Execute("select id,title from news_come_class order by id asc")
 With Response
  Do While not Rs.eof 
     .write " <tr align=""center"">"& vbCrLf
     .write "    <td height=""30"" class=""bottom"">"& Rs("ID") &"</td>"& vbCrLf
     .write "    <td height=""30""  class=""leftadnbottom1"">"& Rs("title") &"</td>"& vbCrLf
     .write "    <td width=""43"" height=""30"" align=""center"" class=""leftadnbottom1""><a href=?id="& Rs("id") &"><img src=""../images/edit.gif"" width=""14"" height=""15"" border=""0""></a></td>"& vbCrLf
     .write "    <td width=""43"" align=""center"" class=""leftadnbottom1""><a href=""del.asp?Table_name=news_come_class&ItemID=id&Id="& Rs("id") &""" onClick=""javascript:return confirm('\nȷ��ɾ����')""><img src=""../images/delete.gif"" width=""10"" height=""13"" border=""0""></a></td>"& vbCrLf
     .write "  </tr>"& vbCrLf     
  Rs.Movenext
  loop
 End With
 Rs.Close:Set Rs=Nothing
End Sub
%>
<!--#include file="../check.asp"-->
<!--#include file="../include/conn.asp"-->
<!--#include file="../include/lib.asp"-->
<%
Call OpenData()
If Session("name") = "" then
response.Write "<script LANGUAGE=javascript>alert('���糬ʱ�������㻹û�е�½! ');this.location.href='../login.asp';</script>"
end if
 	temp_check_rights = Split(session("manconfig"),",")
	for temp_rights_count=0 to ubound(temp_check_rights)
	    if trim(temp_check_rights(temp_rights_count)) = "2" then
			rights_check_passkey = trim(temp_check_rights(temp_rights_count))
		end if
	next
	if rights_check_passkey <> "2" then
Session.Abandon()
	Response.Write "<Script Language=JavaScript>alert('����Ȩ�޲��㣬�벻Ҫ�Ƿ�������������ģ�飬���������˺Ž���ϵͳ�Զ�ɾ��!');this.location.href='../login.asp';</Script>"
	Response.end
	end if%>
<%
If request("Submit")="�ύ" then 
'CompanyID=request("ID")
title =request("title")
  Dim msg  
  Set rs=server.createobject("adodb.recordset")
	'Rs.open "Select * from Sbe_news where ID=" & clng(CompanyID) ,conn,1,3	
	Rs.open "Select * from Sbe_news" ,conn,1,3	
    Rs("title")=Request.Form("title") 
    rs.update  
    rs.close
  Set rs=nothing	
	Response.Write"<script>alert('�����޸ĳɹ�');this.location.href='news.asp';</script>"
else
	StrSQL = "Select * from Sbe_news "
	Set objRec=server.createobject("adodb.recordset")
	 objRec.open StrSQL,conn,1,1
		title = objRec("title")
	objRec.Close:set objRec=Nothing
End if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�������</title>
<link href="../include/style.css" rel="stylesheet" type="text/css">
<script language="javascript">
  function foreColor()
   {
    var arr = showModalDialog("../eWebEditor/Dialog/selcolor.htm", "", "dialogWidth:18.5em; dialogHeight:17.5em; status:0");
    if (arr != null) document.add.title.value='<font color='+arr+'>'+document.add.title.value+'</font>'
    else document.add.title.focus();
}

function clk(value){
 add.writer.value=value;
}
</script>
<script language="JavaScript" src="../include/meizzDate.js"></script>
<style type="text/css">
<!--
.lv {color:#104F50;}
-->
</style>
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" >
  <tr> 
    <td height="25"><font color="#6A859D">��Ϣ��������&gt;&gt; ��Ѷ���� </font></td>
  </tr>
  <tr>
    <td height="1" background="../images/dot.gif"></td>
  </tr>
</table>
<br>
<form name="add" method="post">
<table width="98%"  border="0" align="center" cellpadding="0" cellspacing="0" id="sbe_table">
  <tr align="center">
    <td colspan="3" class="sbe_table_title">�������</td>
  </tr>
  <tr>
    <td align="right">��������:</td>
    <td colspan="2"><textarea name="title" cols="80" rows="5" id="textarea"><%=title%></textarea>
    <img class="Ico" src="../eWebEditor/ButtonImage/standard/forecolor.gif" onClick="foreColor();"></td>
  </tr>
  <tr align="center">
    <td colspan="3"><!--<input type="hidden" name="ID" value="<%=CompanyID%>">--><input name="Submit" type="submit" class="sbe_button" value="�ύ">
    <input name="Submit2" type="reset" class="sbe_button" value="����"></td>
  </tr>
</table>
</form>
<%Call CloseDataBase()%>
</body>
</html>

<!--#include file="../check.asp"-->
<!--#include file="../include/conn.asp"-->
<!--#include file="../include/lib.asp"-->
<%
If Session("name") = "" then
   response.Write "<script LANGUAGE=javascript>alert('���糬ʱ�������㻹û�е�½! ');"
   response.Write"this.location.href='../login.asp';</script>"
   response.End
end if
 	temp_check_rights = Split(session("manconfig"),",")
	for temp_rights_count=0 to ubound(temp_check_rights)
	    if trim(temp_check_rights(temp_rights_count)) = "5" then
			rights_check_passkey = trim(temp_check_rights(temp_rights_count))
		end if
	next
	if rights_check_passkey <> "5" then
Session.Abandon()
	Response.Write "<Script Language=JavaScript>alert('����Ȩ�޲��㣬�벻Ҫ�Ƿ�������������ģ�飬���������˺Ž���ϵͳ�Զ�ɾ��!');this.location.href='../login.asp';</Script>"
	Response.end
	end if
Dim Act
  Act=Request.Form("act")
  Select Case Act
     Case "save":Call SaveData()
	 Case else: Call Main()
  End Select
  Call CloseDataBase()
  Sub SaveData()
    OpenData()
	'=== ���ܲ��� ===
	WebName =       Trim(Request.Form("WebName"))
	EWebName =       Trim(Request.Form("EWebName"))
	Web =           Trim(Request.Form("Web"))
	qq =            Trim(Request.Form("qq"))
	WebName2 =      Trim(Request.Form("WebName2"))
	WebName3 =      Trim(Request.Form("WebName3"))
	EWebName2 =      Trim(Request.Form("EWebName2"))
	EWebName3 =      Trim(Request.Form("EWebName3"))
	Company =       Trim(Request.Form("Company"))
	UpfileType =    Trim(Request.Form("UpfileType"))
	UpfileSize =    Request.Form("UpfileSize")
	PicAuto =       Request.Form("PicAuto")
	PicAutoType =   Request.Form("PicAutoType")
	PicPercent =    Cint(Request.Form("PicPercent"))
	PicHeight =     Cint(Request.Form("PicHeight"))
	PicWidth =      Cint(Request.Form("PicWidth"))
	Watermark =     Request.Form("Watermark")
	WatermarkSize = Cint(Request.Form("WatermarkSize"))
	ShowProClass =  Request.Form("ShowProClass")
	ShowNewsClass = Request.Form("ShowNewsClass")
	Template =      Trim(Request.Form("Template"))
	jsqtoday =      Trim(Request.Form("jsqtoday"))
	jsq =      Trim(Request.Form("jsq"))
	ShowNewsPic =   cint(Request.Form("ShowNewsPic"))
	ShowNewsAbout = cint(Request.Form("ShowNewsAbout"))
	msn = Request.Form("msn")
	WatermarkWord = Trim(Request.Form("WatermarkWord"))
	   flag_web=Request.Form("flag_web")
	web_miaoshu = Trim(Request.Form("web_miaoshu"))
	
	tel1 = Trim(Request.Form("tel1"))
	tel2 = Trim(Request.Form("tel2"))
	tel3 = Trim(Request.Form("tel3"))
	email = Trim(Request.Form("email"))	
    mailaddress = Trim(Request.Form("mailaddress"))
	mailsend =  Request.Form("mailsend")
	mailusername = Request.Form("mailusername")
	mailuserpass =      Trim(Request.Form("mailuserpass"))
	mailname =   trim(Request.Form("mailname"))	
	'=== ���ս��� ===
	
	'=== ��֤���� ===
	'=== ��֤���� ===
	
	'=== �������� ===
	Set Rs=Server.CreateObject("adodb.recordset")
    Sql="Select * From WebConfig"
	Rs.Open Sql,Conn,1,3
	   Rs("WebName") =       WebName
	   Rs("Web") =           Web
	   Rs("qq") =            qq
	   Rs("WebName2") =      WebName2
	   Rs("WebName3") =      WebName3
	   Rs("EWebName") =       EWebName
	   Rs("EWebName2") =      EWebName2
	   Rs("EWebName3") =      EWebName3
	   Rs("Company") =       Company
	   if session("flag")=99 then
	   Rs("msn") =           msn
	   Rs("WatermarkWord") = WatermarkWord
	   Rs("flag_web") =      flag_web
	   Rs("web_miaoshu") =   web_miaoshu
	   end if
	   Rs("tel1") =      tel1
	   Rs("tel2") =      tel2
	   Rs("tel3") =      tel3
	   Rs("email") =     email
	   Rs("jsqtoday") =      jsqtoday
	   Rs("jsq") =     jsq		
	   Rs("mailaddress") =      mailaddress
	   Rs("mailsend") =      mailsend
	   Rs("mailusername") =     mailusername
	   Rs("mailuserpass") =      mailuserpass
	   Rs("mailname") =     mailname	    
	   Rs.Update
	 Rs.Close
	 Set Rs=Nothing
    '=== ������� ===
	   Response.Write("<script language=javascript>alert('�����޸ĳɹ�!');window.location.href='main1.asp';</script>")
	   Response.End()
  End Sub 
  Sub Main()
  OpenData()
  Dim Rs,Sql
  Set Rs=Server.CreateObject("adodb.recordset")
  Sql="Select * From WebConfig"
  Rs.Open Sql,Conn,1,1
   If Rs.Eof Then
      Response.Write("������Ϣ�Ѿ���ɾ����")
	  Response.End()
   Else
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��̨����ϵͳ</title>

<link href="../include/style.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="25"><font color="#6A859D">��վ���� &gt;&gt; ��վ��������</font></td>
  </tr>
  <tr>
    <td height="1" background="../images/dot.gif"></td>
  </tr>
</table>
<form name="form" method="post" action="">
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="0" id="sbe_table">
<tr > 
    <td width="16%" height="25">��ַ��</td>
    <td width="84%" height="21"><input name="Web" type="text" class="input" id="Web" value="<%=rs("Web")%>" size="50"></td>
  </tr>
  <tr > 
    <td width="16%" height="25">��վ����(��)��</td>
    <td width="84%" height="21"><input name="WebName" type="text" class="input" id="WebName" value="<%=rs("WebName")%>" size="50"></td>
  </tr>
  <tr > 
    <td width="16%" height="25">��վ�ؼ���(��)��</td>
    <td width="84%" height="21"><input name="WebName2" type="text" class="input" id="WebName2" value="<%=rs("WebName2")%>" size="50" maxlength="48">      
      (��ʹ��Ӣ�Ķ���:&quot; , &quot;,���ֻ������48���ַ�)</td>
  </tr>
   <tr > 
    <td width="16%" height="25">��վ����(��)</td>
    <td width="84%" height="21"><input name="WebName3" type="text" class="input" id="WebName3" value="<%=rs("WebName3")%>" size="50" maxlength="64">
      (��������������.���ֻ������64���ַ�)</td>
  </tr>
  <tr> 
    <td height="25">��˾���ƣ�</td>
    <td height="21"><input name="Company" type="text" class="input" id="Company" value="<%=rs("Company")%>" size="50"></td>
  </tr> 
  <tr <%=banben_display%>>
    <td width="16%" height="25">��վ����(Ӣ)��</td>
    <td width="84%" height="21"><input name="EWebName" type="text" class="input" id="EWebName" value="<%=rs("EWebName")%>" size="50"></td>
  </tr>
  <tr <%=banben_display%>> 
    <td width="16%" height="25">��վ�ؼ���(Ӣ)��</td>
    <td width="84%" height="21"><input name="EWebName2" type="text" class="input" id="EWebName2" value="<%=rs("EWebName2")%>" size="50" maxlength="48">      
      (��ʹ��Ӣ�Ķ���:&quot; , &quot;,���ֻ������48���ַ�)</td>
  </tr>
   <tr <%=banben_display%>> 
    <td width="16%" height="25">��վ����(Ӣ)</td>
    <td width="84%" height="21"><input name="EWebName3" type="text" class="input" id="EWebName3" value="<%=rs("EWebName3")%>" size="50" maxlength="64">
      (��������������.���ֻ������64���ַ�)</td>
  </tr>
<tr> 
    <td height="25">���շ�������</td>
    <td height="21"><input name="jsqtoday" type="text" class="input" id="jsqtoday" value="<%=rs("jsqtoday")%>" size="50" onKeyPress="return event.keyCode>=48&&event.keyCode<=57" onpaste="return !clipboardData.getData('text').match(/\D/)" ondragenter="return false" style="ime-mode:Disabled;"></td>
  </tr>
<tr> 
    <td height="25">��վ��������</td>
    <td height="21"><input name="jsq" type="text" class="input" id="jsq" value="<%=rs("jsq")%>" size="50" onKeyPress="return event.keyCode>=48&&event.keyCode<=57" onpaste="return !clipboardData.getData('text').match(/\D/)" ondragenter="return false" style="ime-mode:Disabled;"></td>
  </tr>
  <tr style="display:none;"> 
    <td height="25">��˾��ַ��</td>
    <td height="21"><input name="qq" type="text" class="input" id="qq" value="<%=rs("qq")%>" size="50"></td>
  </tr>
<tr style="display:none;"> 
    <td height="25">���쵥λ��</td>
    <td height="21"><input name="tel1" type="text" class="input" id="qq" value="<%=rs("tel1")%>" size="50"></td>
  </tr>
<tr style="display:none;"> 
    <td height="25">�绰��</td>
    <td height="21"><input name="tel2" type="text" class="input" id="qq" value="<%=rs("tel2")%>" size="50"></td>
  </tr>
<tr style="display:none;"> 
    <td height="25">���棺</td>
    <td height="21"><input name="tel3" type="text" class="input" id="qq" value="<%=rs("tel3")%>" size="50"></td>
  </tr>
<tr style="display:none;"> 
    <td height="25">E_mail��</td>
    <td height="21"><input name="email" type="text" class="input" id="qq" value="<%=rs("email")%>" size="50"></td>
  </tr>  
<tr> 
    <td height="25" colspan="2">�����ʼ�������:</td>
    </tr> 
<tr> 
    <td height="25">�ʼ���������ַ��</td>
    <td height="21"><input name="mailaddress" type="text" id="qq" size="50" value="<%=rs("mailaddress")%>" class="input"></td>
  </tr>  
<tr> 
    <td height="25">�������䣺</td>
    <td height="21"><input name="mailsend" type="text" id="qq" size="50" value="<%=rs("mailsend")%>" class="input"></td>
  </tr>
<tr> 
    <td height="25">�� ¼ ����</td>
    <td height="21"><input name="mailusername" type="text" id="qq" size="50" value="<%=rs("mailusername")%>" class="input"></td>
  </tr>
<tr> 
    <td height="25">��¼���룺</td>
    <td height="21"><input name="mailuserpass" type="password" id="qq" size="50" value="<%=rs("mailuserpass")%>" class="input"></td>
  </tr>
<tr> 
    <td height="25">��ʾ��������</td>
    <td height="21"><input name="mailname" type="text" id="qq" size="50" value="<%=rs("mailname")%>" class="input"></td>
  </tr>
  <tr> 
    <td height="25">����֧�֣�</td>
    <td height="21"><a href="<%=rs("msn")%>" target="_blank"><%=rs("WatermarkWord")%></a></td>
  </tr>
 <%if session("flag")=99 then%>
<SCRIPT language=javascript>
function show_user_rights_menu(menu_id)
{
if (menu_id==1)
{
eval("show_user_rights.style.display=\"none\";");
document.form.web_miaoshu.value="";
}
else
{
eval("show_user_rights.style.display=\"\";");
document.form.web_miaoshu.value="��վά���С�";
}
}
</SCRIPT>
  <tr> 
    <td width="16%" height="25">��վ״̬��</td>
    <td width="84%" height="21"><input name="flag_web" type="radio" class="input" id="flag_web" value="1" <%if rs("flag_web")=true then response.Write("checked") end if%> onclick=show_user_rights_menu(1)>
&nbsp;��&nbsp;&nbsp; <input name="flag_web" type="radio" class="input" id="flag_web" value="0" <%if rs("flag_web")=false then response.Write("checked") end if%> onclick=show_user_rights_menu(0)>
    �ر� </td>
  </tr>
 <% end if%>
   <tr id="show_user_rights" <%if rs("flag_web")=false then response.Write("style='display:'") else response.Write("style='display:none' end if")%>> 
    <td width="16%" height="25">״̬����</td>
    <td width="84%" height="21"><input name="web_miaoshu" type="text" class="input" id="web_miaoshu" value="<%=rs("web_miaoshu")%>" size="50">
      (��վ�رյ�����.)</td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="0" id="sbe_table" <%if session("flag")<>99 then response.Write("style='display:none;'") end if%>>
<tr > 
    <td width="16%" height="25">����֧���޸�-��ַ��</td>
    <td width="84%" height="21"><input name="msn" type="text" class="input" id="msn" value="<%=rs("msn")%>" size="50"></td>
  </tr>
  <tr > 
    <td width="16%" height="25">��˾��</td>
    <td width="84%" height="21"><input name="WatermarkWord" type="text" class="input" id="WatermarkWord" value="<%=rs("WatermarkWord")%>" size="50"></td>
  </tr>
</table>
  <br>
  <p align="center">
    <input name="act" type="hidden" id="act" value="save">
    <input type="submit" name="Submit" value="�ύ�޸�" class="sbe_button">
    <br>
    <br>
  </p>
</form>
</body>
</html>
<%End If
  Rs.Close
  Set Rs=Nothing
  End Sub%>
<%
  Sub JudgeTemplate(Str,Str1)
     If Instr(Str,Str1)<>0 Then
	    response.Write("checked")
	 End If  
  End Sub
%>
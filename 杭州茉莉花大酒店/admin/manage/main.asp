<!--#include file="../check.asp"-->
<!--#include file="../include/conn.asp"-->
<!--#include file="../include/lib.asp"-->
<%Dim Act
If Session("flag")<>99 then
Session.Abandon()
response.Write "<script LANGUAGE=javascript>alert('����Ȩ�޲��㣬�벻Ҫ�Ƿ�������������ģ�飬���������˺Ž���ϵͳ�Զ�ɾ��! ');this.location.href='../login.asp';</script>"
end if
  Act=Request.Form("act")
  Select Case Act
     Case "save":Call SaveData()
	 Case else: Call Main()
  End Select
  Call CloseDataBase()  
  
  Sub SaveData()
    OpenData()
    Dim WebName,Company,UpfileType,UpfileSize,SmtpHost,SmtpUser,SmtpPwd,PicAuto,PicAutoType,PicPercent,PicHeight,PicWidth,Watermark,WatermarkSize,WatermarkWord,ShowProClass,ShowNewsClass,Template
    Dim Rs,Sql
	'=== ���ܲ��� ===
	WebName =       Trim(Request.Form("WebName"))
	Company =       Trim(Request.Form("Company"))
	UpfileType =    Trim(Request.Form("UpfileType"))
	UpfileSize =    Request.Form("UpfileSize")
	SmtpHost =      Trim(Request.Form("SmtpHost"))
	SmtpUser =      Trim(Request.Form("SmtpUser"))
	SmtpPwd =       Trim(Request.Form("SmtpPwd"))
	PicAuto =       Request.Form("PicAuto")
	PicAutoType =   Request.Form("PicAutoType")
	PicPercent =    Cint(Request.Form("PicPercent"))
	PicHeight =     Cint(Request.Form("PicHeight"))
	PicWidth =      Cint(Request.Form("PicWidth"))
	Watermark =     Request.Form("Watermark")
	WatermarkSize = Cint(Request.Form("WatermarkSize"))
	WatermarkWord = Trim(Request.Form("WatermarkWord"))
	ShowProClass =  Request.Form("ShowProClass")
	ShowNewsClass = Request.Form("ShowNewsClass")
	Template =      Trim(Request.Form("Template"))
	ShowNewsPic =   cint(Request.Form("ShowNewsPic"))
	ShowNewsAbout = cint(Request.Form("ShowNewsAbout"))	
	UpmovieType =   Trim(Request.Form("UpmovieType"))
	UpmovieSize =   Trim(Request.Form("UpmovieSize"))
	FtpUrl =   Trim(Request.Form("FtpUrl"))
	UserName =   Trim(Request.Form("UserName"))
	Password =   Trim(Request.Form("Password"))
	NewsClass_num =   Trim(Request.Form("NewsClass_num"))
	ProClass_num =   Trim(Request.Form("ProClass_num"))
	Pro_order =   Trim(Request.Form("Pro_order"))
	banben =   Trim(Request.Form("banben"))
	ShowqiyeClass =   Trim(Request.Form("ShowqiyeClass"))
	qiyeClass_num =   Trim(Request.Form("qiyeClass_num"))
	hy_message =   Trim(Request.Form("hy_message"))
	sf_yingpin =   Trim(Request.Form("sf_yingpin"))
	yanzhengma =   Trim(Request.Form("yanzhengma"))
	weblink_leibie =   Trim(Request.Form("weblink_leibie"))
	ShowdownClass =   Trim(Request.Form("ShowdownClass"))
	downClass_num =   Trim(Request.Form("downClass_num"))
	'=== ���ս��� ===
	
	'=== ��֤���� ===
	If WebName = "" or SmtpHost = "" Or SmtpUser = "" Or SmtpPwd = "" Or Template = "" Then
	   Response.Write("<script language=javascript>alert('��ѱ���д����!');window.history.back();</script>")
	   Response.End()
	End If
	'=== ��֤���� ===
	
	'=== �������� ===
	Set Rs=Server.CreateObject("adodb.recordset")
    Sql="Select * From Sbe_WebConfig"
	Rs.Open Sql,Conn,1,3
	   Rs("WebName") =       WebName
	   Rs("NewsClass_num") = NewsClass_num
	   Rs("ProClass_num") =  ProClass_num  
	   Rs("Company") =       Company
	   Rs("UpfileType") =    UpfileType
	   Rs("UpfileSize") =    UpfileSize
	   Rs("UpmovieType") =    UpmovieType
	   Rs("UpmovieSize") =    UpmovieSize
	   Rs("FtpUrl")  =        FtpUrl
	   Rs("UserName") =       UserName
	   Rs("Password") =       Password   
	   Rs("SmtpHost") =      SmtpHost
	   Rs("SmtpUser") =      SmtpUser
	   Rs("SmtpPwd") =       SmtpPwd
	   Rs("PicAuto") =       PicAuto
	   Rs("PicAutoType") =   PicAutoType
	   Rs("PicPercent") =    PicPercent
	   Rs("PicHeight") =     PicHeight
	   Rs("PicWidth") =      PicWidth 
	   Rs("Watermark") =     Watermark
	   Rs("WatermarkSize") = WatermarkSize
	   Rs("WatermarkWord") = WatermarkWord
	   Rs("ShowProClass") =  ShowProClass
	   Rs("ShowNewsClass") = ShowNewsClass
	   Rs("ShowNewsPic")= ShowNewsPic
	   Rs("ShowNewsAbout")= ShowNewsAbout	   
	   Rs("Template") =      "0, "&Template
	   Rs("Pro_order")= Pro_order
	   Rs("banben")= banben
	   Rs("Pro_order")= Pro_order
	   Rs("qiyeClass_num")= qiyeClass_num
	   Rs("ShowqiyeClass")= ShowqiyeClass
	   Rs("hy_message")= hy_message
	   Rs("sf_yingpin")= sf_yingpin
	   Rs("yanzhengma")= yanzhengma 
	   Rs("weblink_leibie")= weblink_leibie
	   Rs("ShowdownClass")= ShowdownClass 
	   Rs("downClass_num")= downClass_num
	   Rs.Update
	 Rs.Close
	 Set Rs=Nothing
    '=== ������� ===
	   Response.Write("<script language=javascript>alert('�����޸ĳɹ�!');window.location.href='main.asp';</script>")
	   Response.End()
  End Sub
  
  Sub Main()
  OpenData()
  Dim Rs,Sql
  Set Rs=Server.CreateObject("adodb.recordset")
  Sql="Select * From Sbe_WebConfig"
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
<form name="form1" method="post" action="">
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="0" id="sbe_table">
  <tr > 
    <td width="14%" height="25">��վ����</td>
    <td width="86%" height="21"><input name="WebName" type="text" class="input" id="WebName" value="<%=server.HTMLEncode(rs("WebName"))%>" size="50"></td>
  </tr>
  <tr> 
    <td height="25">��˾����</td>
    <td height="21"><input name="Company" type="text" class="input" id="Company" value="<%=server.HTMLEncode(rs("Company"))%>" size="50"></td>
  </tr>
  <tr> 
    <td height="25">�ϴ��ļ�����</td>
    <td height="21"><input name="UpfileType" type="text" class="input" id="UpfileType" value="<%=rs("UpfileType")%>" size="40"></td>
  </tr>
  <tr> 
    <td height="25">�ϴ��ļ���С</td>
    <td height="21"><input name="UpfileSize" type="text" class="input" id="UpfileSize" value="<%=rs("UpfileSize")%>" size="10">
        K</td>
  </tr>
  <tr>
    <td height="25">�ϴ���Ƶ����</td>
    <td height="21"><input name="UpmovieType" type="text" class="input" id="UpmovieType" value="<%=rs("UpmovieType")%>" size="40"></td>
  </tr>
  <tr>
    <td height="25">�ϴ���Ƶ��С</td>
    <td height="21"><input name="UpmovieSize" type="text" class="input" id="UpmovieSize" value="<%=rs("UpmovieSize")%>" size="10">
      K</td>
  </tr>
</table>
<br>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="0" id="sbe_table">
  <tr > 
    <td width="14%" height="25">ftp��ַ</td>
    <td width="86%" height="21"><input name="FtpUrl" type="text" class="input" id="FtpUrl" value="<%=rs("FtpUrl")%>"> </td>
  </tr>
  <tr> 
    <td height="25">�û���</td>
    <td height="21"><input name="UserName" type="text" class="input" id="UserName" value="<%=rs("UserName")%>"></td>
  </tr>
  <tr> 
    <td height="25">����</td>
    <td height="21"><input name="Password" type="text" class="input" id="Password" value="<%=rs("Password")%>"></td>
  </tr>
</table>
<br>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="0" id="sbe_table">
  <tr > 
    <td width="14%" height="25">Smtp������</td>
    <td width="86%" height="21"><input name="SmtpHost" type="text" class="input" id="SmtpHost" value="<%=rs("SmtpHost")%>"></td>
  </tr>
  <tr> 
    <td height="25">Smtp�û�</td>
    <td height="21"><input name="SmtpUser" type="text" class="input" id="SmtpUser" value="<%=rs("SmtpUser")%>"></td>
  </tr>
  <tr> 
    <td height="25">Smtp����</td>
    <td height="21"><input name="SmtpPwd" type="text" class="input" id="SmtpPwd" value="<%=rs("SmtpPwd")%>"></td>
  </tr>
</table>
<br>
<script language="JavaScript">
<!--
  function PicTypeShow(flag){
   if (flag==1){
       PicType1.style.display="";
	   PicType2.style.display="none";}
   if (flag==2){
       PicType1.style.display="none";
	   PicType2.style.display="";}
  }
  function WatermarkShow(flag){
   if (flag==1){
       Watermark1.style.display="";
	   Watermark2.style.display="";
	   }
   if (flag==2){
       Watermark1.style.display="none";
	   Watermark2.style.display="none";
	   }
  }
  function NewsClassShow(flag){
   if (flag==1){
       Show_NewsClass_num.style.display="";
	   }
   if (flag==0){
       Show_NewsClass_num.style.display="none";
	   }
  }
  function ProClassShow(flag){
   if (flag==1){
       Show_ProClass_num.style.display="";
	   }
   if (flag==0){
       Show_ProClass_num.style.display="none";
	   }
  }
  function qiyeClassShow(flag){
   if (flag==1){
       Show_qiyeClass_num.style.display="";
	   }
   if (flag==0){
       Show_qiyeClass_num.style.display="none";
	   }
	   }
   function downClassShow(flag){
   if (flag==1){
       Show_downClass_num.style.display="";
	   }
   if (flag==0){
       Show_downClass_num.style.display="none";
	   }
  }
//function ServerShow(){
//if (document.form1.Template[7].checked == true) {
//	show_server.style.display = "";
//  }else{
//	show_server.style.display = "none";
//}
//} 
-->
</script>
  <table width="98%" border="0" align="center" cellpadding="3" cellspacing="0" id="sbe_table">
    <tr > 
      <td height="25">�Ƿ��Զ�����Сͼ</td>
      <td height="21"><input type="radio" name="PicAuto" value="1" <%Call ReturnSel(true,rs("PicAuto"),2)%>>
        �� 
        <input type="radio" name="PicAuto" value="0" <%Call ReturnSel(false,rs("PicAuto"),2)%>>
        ��</td>
    </tr>
    <tr > 
      <td height="25">Сͼ����</td>
      <td height="21"><input type="radio" name="PicAutoType" value="1" <%Call ReturnSel(1,rs("PicAutoType"),2)%> onClick="PicTypeShow(1)">
        �ٷֱ� 
          <input type="radio" name="PicAutoType" value="2" <%Call ReturnSel(2,rs("PicAutoType"),2)%> onClick="PicTypeShow(2)">
        ���� </td>
    </tr>
    <tr id="PicType1" <%if rs("PicAutoType")=2 then%>style="display:none"<%end if%>>
      <td height="25">�ٷֱ�</td>
      <td height="21"><input name="PicPercent" type="text" class="input" id="PicPercent" value="<%=rs("PicPercent")%>" size="2" maxlength="2">
        %</td>
    </tr>
    <tr id="PicType2" <%if rs("PicAutoType")=1 then%>style="display:none"<%end if%>> 
      <td height="25">����</td>
      <td height="21">�� <input name="PicHeight" type="text" class="input" id="PicHeight" value="<%=rs("PicHeight")%>" size="5">
        px &nbsp;&nbsp;&nbsp;&nbsp;�� <input name="PicWidth" type="text" class="input" id="PicWidth" value="<%=rs("PicWidth")%>" size="5">
        px</td>
    </tr>
    <tr > 
      <td width="14%" height="25">�Ƿ�����ͼƬˮӡ</td>
      <td width="86%" height="21"><input type="radio" name="Watermark" value="1" <%Call ReturnSel(true,rs("Watermark"),2)%> onClick="WatermarkShow(1)">
        �� 
        <input type="radio" name="Watermark" value="0" <%Call ReturnSel(false,rs("Watermark"),2)%> onClick="WatermarkShow(2)">
        ��</td>
    </tr>
    <tr id="Watermark1"  <%if rs("Watermark")=false then%>style="display:none"<%end if%>> 
      <td height="25">���ִ�С</td>
      <td height="21"><input name="WatermarkSize" type="text" class="input" id="WatermarkSize" value="<%=rs("WatermarkSize")%>" size="5"></td>
    </tr>
    <tr id="Watermark2" <%if rs("Watermark")=false then%>style="display:none"<%end if%>> 
      <td height="25">ˮӡ����</td>
      <td height="21"><input name="WatermarkWord" type="text" class="input" id="WatermarkWord" value="<%=rs("Watermarkword")%>"></td>
    </tr>
  </table>
  <br>
  <br>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="0" id="sbe_table">
    <tr > 
      <td width="14%" height="25">ϵͳ����:</td>
      <td width="86%" height="21"><input name="banben" type="radio" id="banben" value="1" <%Call ReturnSel(1,rs("banben"),2)%>>
        ���İ�&nbsp;&nbsp;
        <input name="banben" type="radio" id="banben" value="2" <%Call ReturnSel(2,rs("banben"),2)%>>
      ��Ӣ�İ�</td>
    </tr>
    <tr > 
      <td width="14%" height="25">������������:</td>
      <td width="86%" height="21"><input name="weblink_leibie" type="radio" id="banben" value="0" <%Call ReturnSel(0,rs("weblink_leibie"),2)%>>
        ����&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name="weblink_leibie" type="radio" id="banben" value="1" <%Call ReturnSel(1,rs("weblink_leibie"),2)%>>
      ͼƬ&nbsp;&nbsp;
      &nbsp;
      &nbsp;
      <input name="weblink_leibie" type="radio" id="radio" value="2" <%Call ReturnSel(2,rs("weblink_leibie"),2)%>>
����+ͼƬ</td>
    </tr>
  </table>
  <br>
  <table width="98%" border="0" align="center" cellpadding="3" cellspacing="0" id="sbe_table">
    <tr > 
      <td width="14%" height="25">�Ƿ���ʾ��Ѷ����</td>
      <td width="86%" height="21"><input type="radio" name="ShowNewsClass" value="1" <%Call ReturnSel(true,rs("ShowNewsClass"),2)%> onClick="NewsClassShow(1)">
        ��ʾ
<input type="radio" name="ShowNewsClass" value="0" <%Call ReturnSel(false,rs("ShowNewsClass"),2)%> onClick="NewsClassShow(0)">
        ���� &nbsp;&nbsp;&nbsp;<span id="Show_NewsClass_num" <%if rs("ShowNewsClass")=false then response.Write("style='display:none'") end if%>>������
        <input name="NewsClass_num" type="text" class="input" style="ime-mode:Disabled;" onKeyPress="return event.keyCode>=48&&event.keyCode<=57||event.keyCode==45" value="<%=rs("NewsClass_num")%>" size="4" maxlength="4" onpaste="return !clipboardData.getData('text').match(/\D/)" ondragenter="return false">
      ��&nbsp;����(<font color="#FF0000">ֻ�������0������</font>)</span></td>
    </tr>
    <tr> 
      <td height="25">�Ƿ���ʾ��Ʒ����</td>
      <td height="21"> 
        <input type="radio" name="ShowProClass" value="1" <%Call ReturnSel(true,rs("ShowProClass"),2)%> onClick="ProClassShow(1)">
        ��ʾ 
        <input type="radio" name="ShowProClass" value="0" <%Call ReturnSel(false,rs("ShowProClass"),2)%> onClick="ProClassShow(0)">
        ����  <span id="Show_ProClass_num" <%if rs("ShowProClass")=false then response.Write("style='display:none'") end if%>>&nbsp;&nbsp;&nbsp;������
        <input name="ProClass_num" type="text" class="input" style="ime-mode:Disabled;" onKeyPress="return event.keyCode>=48&&event.keyCode<=57||event.keyCode==45" value="<%=rs("ProClass_num")%>" size="4" maxlength="4" onpaste="return !clipboardData.getData('text').match(/\D/)" ondragenter="return false">
��&nbsp;����(<font color="#FF0000">ֻ�������0������</font>)</span></td>
    </tr>
<tr> 
      <td height="25">�Ƿ���ʾ��Ϣ����</td>
      <td height="21"> 
        <input type="radio" name="ShowqiyeClass" value="1" <%Call ReturnSel(true,rs("ShowqiyeClass"),2)%> onClick="qiyeClassShow(1)">
        ��ʾ 
        <input type="radio" name="ShowqiyeClass" value="0" <%Call ReturnSel(false,rs("ShowqiyeClass"),2)%> onClick="qiyeClassShow(0)">
        ����  <span id="Show_qiyeClass_num" <%if rs("ShowqiyeClass")=false then response.Write("style='display:none'") end if%>>&nbsp;&nbsp;&nbsp;������
        <input name="qiyeClass_num" type="text" class="input" style="ime-mode:Disabled;" onKeyPress="return event.keyCode>=48&&event.keyCode<=57||event.keyCode==45" value="<%=rs("qiyeClass_num")%>" size="4" maxlength="4" onpaste="return !clipboardData.getData('text').match(/\D/)" ondragenter="return false">
��&nbsp;����(<font color="#FF0000">ֻ�������0������</font>)</span></td>
    </tr>
<tr> 
      <td height="25">�Ƿ���ʾ���ط���</td>
      <td height="21"> 
        <input type="radio" name="ShowdownClass" value="1" <%Call ReturnSel(true,rs("ShowdownClass"),2)%> onClick="downClassShow(1)">
        ��ʾ 
        <input type="radio" name="ShowdownClass" value="0" <%Call ReturnSel(false,rs("ShowdownClass"),2)%> onClick="downClassShow(0)">
        ����  <span id="Show_downClass_num" <%if rs("ShowdownClass")=false then response.Write("style='display:none'") end if%>>&nbsp;&nbsp;&nbsp;������
        <input name="downClass_num" type="text" class="input" style="ime-mode:Disabled;" onKeyPress="return event.keyCode>=48&&event.keyCode<=57||event.keyCode==45" value="<%=rs("downClass_num")%>" size="4" maxlength="4" onpaste="return !clipboardData.getData('text').match(/\D/)" ondragenter="return false">
��&nbsp;����(<font color="#FF0000">ֻ�������0������</font>)</span></td>
    </tr>
  </table>
  <br>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="0" id="sbe_table">
    <tr > 
      <td width="14%" height="25">��Ʒ�Ƿ��ύ����</td>
      <td width="36%" height="21"><input type="radio" name="Pro_order" value="1" <%Call ReturnSel(true,rs("Pro_order"),2)%>>
        �� 
        <input type="radio" name="Pro_order" value="0" <%Call ReturnSel(false,rs("Pro_order"),2)%>>
        ��</td>
      <td width="17%">���԰��Ƿ��лظ�</td>
      <td width="33%">&nbsp;
        <input type="radio" name="hy_message" value="1" <%Call ReturnSel(true,rs("hy_message"),2)%>>
        �� 
        <input type="radio" name="hy_message" value="0" <%Call ReturnSel(false,rs("hy_message"),2)%>>
      ��</td>
    </tr>
<tr > 
      <td width="14%" height="25">������Ƹ�Ƿ�ӦƸ</td>
      <td width="36%" height="21"><input name="sf_yingpin" type="radio" id="sf_yingpin" value="1" <%Call ReturnSel(true,rs("sf_yingpin"),2)%>>
        ��
        <input name="sf_yingpin" type="radio" id="sf_yingpin" value="0" <%Call ReturnSel(false,rs("sf_yingpin"),2)%>>
        ��</td>
      <td width="17%" height="25">��̨��¼�Ƿ�����֤��</td>
      <td width="33%" height="21">&nbsp;
<input type="radio" name="yanzhengma" value="1" <%Call ReturnSel(true,rs("yanzhengma"),2)%>>
��
  <input type="radio" name="yanzhengma" value="0" <%Call ReturnSel(false,rs("yanzhengma"),2)%>>
��</td>
    </tr>
  </table>
  <br>

  <table width="98%" border="0" align="center" cellpadding="3" cellspacing="0" id="sbe_table">
    <tr > 
      <td width="14%" height="25">��վ����ģ��</td>
      <td width="86%" height="21">
	    <input name="Template" type="checkbox" id="Template" value="0" checked disabled>
        ��̨����&nbsp; 
        <input name="Template" type="checkbox" id="Template" value="1" <%Call JudgeTemplate(rs("template"),", 1")%>>        
        ��ҵ��Ϣ &nbsp; 
        <input name="Template" type="checkbox" id="Template" value="2" <%Call JudgeTemplate(rs("template"),", 2")%>>
        ��Ʒչʾ&nbsp; 
        <input name="Template" type="checkbox" id="Template" value="3" <%Call JudgeTemplate(rs("template"),", 3")%>>
        ��Ѷ����&nbsp; 
        <input name="Template" type="checkbox" id="Template" value="4" <%Call JudgeTemplate(rs("template"),", 4")%>>
        ���ع��� &nbsp; 
        <input name="Template" type="checkbox" id="Template" value="5" <%Call JudgeTemplate(rs("template"),", 5")%>>
        Ȩ�޹���&nbsp;<br>
        <input name="Template" type="checkbox" id="Template" value="6" <%Call JudgeTemplate(rs("template"),", 6")%>>
        ������Ƹ&nbsp;
		<input name="Template" type="checkbox" id="Template" value="7" <%Call JudgeTemplate(rs("template"),", 7")%>>
        ��������&nbsp; &nbsp;
        <input name="Template" type="checkbox" id="Template" value="8" <%Call JudgeTemplate(rs("template"),", 8")%>>
        ��������&nbsp; 
        <input name="Template" type="checkbox" id="Template" value="9" <%Call JudgeTemplate(rs("template"),", 9")%>>
        �������� 
       <!-- <input name="Template" type="checkbox" id="Template" value="10" <%'Call JudgeTemplate(rs("template"),", 10")%>>
        FAQϵͳ --></td>
    </tr>
  </table>
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
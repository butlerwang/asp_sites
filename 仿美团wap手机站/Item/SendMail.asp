<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<%

Dim KS
Set KS=New PublicCls
Dim ChannelID,ID,RS,ArticleUrl,WebName,WebUrl
Dim ReturnInfo,Subject,MyName,MyMail,FrName,FrMail,MailBody
ChannelID=KS.ChkClng(KS.S("m"))
ID=KS.ChkClng(KS.S("ID"))
ArticleUrl=Request.ServerVariables("HTTP_REFERER")
if ID=0 or ChannelID=0 then
	Response.Write"<script>alert(""错误的参数！"");location.href=""javascript:history.back()"";</script>"
    Response.End
end if
Set RS=Server.CreateObject("Adodb.Recordset")
RS.Open "Select top 1 * From " & KS.C_S(ChannelID,2) & " Where ID=" & ID,conn,1,1
IF RS.EOF And RS.BoF Then
  RS.CLOSE
 SET RS=NOthing
 Call CloseConn()
 Set KS=Nothing
	Response.Write"<script>alert(""错误的参数！"");location.href=""javascript:history.back()"";</script>"
    Response.End
End if
WebName=KS.Setting(0)
WebUrl=KS.Setting(2)
MailServerAddress=KS.Setting(12)
IF KS.S("Action")="Send" Then
FrName=KS.S("FrName")
MyName=KS.S("MyName")
FrMail=KS.S("FrMail")
IF FrMail="" Then
		Response.Write"<script>alert(""好友邮箱地址不能为空！"");location.href=""javascript:history.back()"";</script>"
        Response.End
End IF
IF KS.IsValidEmail(FrMail)=false then
		Response.Write"<script>alert(""好友邮箱地址格式有误！" & FrMail & """);location.href=""javascript:history.back()"";</script>"
        Response.End
End if
MyMail=KS.S("MyMail")
IF MyMail="" Then
		Response.Write"<script>alert(""您的邮箱地址不能为空！"");location.href=""javascript:history.back()"";</script>"
        Response.End
End IF
if KS.IsValidEmail(MyMail)=false then
		Response.Write"<script>alert(""您的邮箱地址格式有误！"");location.href=""javascript:history.back()"";</script>"
        Response.End
End if
Content=KS.S("Content")


Subject="您好" & KS.S("FrName") & ",您的朋友"&KS.S("MyName")&"从" & KS.S("SiteName") & "给您发来的一篇信息资料"
	MailBody=MailBody &"<style>A:visited {	TEXT-DECORATION: none	}"
	MailBody=MailBody &"A:active  {	TEXT-DECORATION: none	}"
	MailBody=MailBody &"A:hover   {	TEXT-DECORATION: underline overline	}"
	MailBody=MailBody &"A:link 	  {	text-decoration: none;}"
	MailBody=MailBody &"A:visited {	text-decoration: none;}"
	MailBody=MailBody &"A:active  {	TEXT-DECORATION: none;}"
	MailBody=MailBody &"A:hover   {	TEXT-DECORATION: underline overline}"
	MailBody=MailBody &"BODY   {	FONT-FAMILY: 宋体; FONT-SIZE: 9pt;}"
	MailBody=MailBody &"TD	   {	FONT-FAMILY: 宋体; FONT-SIZE: 9pt	}</style>"

	MailBody=MailBody &"<table border='0' width='90%' align='center'><tr>"
	MailBody=MailBody &"<td valign='middle' align='top'>"
    If ChannelID=5 Then
	    MailBody=MailBody &Content & "<br>以下是商品介绍<br>" & RS("prointro") 
	Else
       MailBody=MailBody &Content & "<br>以下是信息正文<br>" & RS("ArticleContent") 
	End If
	MailBody=MailBody &"</td></tr></table>"

'开始发送
ReturnInfo=KS.SendMail(MailServerAddress,KS.Setting(13), KS.Setting(14),Subject,FrMail,KS.S("MyName"),MailBody,KS.Setting(11))
  IF ReturnInfo="OK" Then
    Response.Write ("<script>alert('信件成功发送!');window.close();</script>")
	 Response.End
  Else
    Response.Write ("<script>alert('信件发送失败!失败原因:\n" & ReturnInfo & "');window.close();</script>")
	Response.End
  End if
End if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>发送电子邮件</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<link href="/images/style.css" rel="stylesheet">
<body>
<table width="770" height="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
<form action="?action=Send" name="myform" method="post">
  <tr>
    <td><table width="60%" border="0" align="center" cellpadding="0" cellspacing="1">
      <tr bgcolor="#FFFFFF">
        <td width="107" height="30" align="center"> 好友姓名：</td>
        <td width="345" height="30"><input name="FrName" type="text" id="FrName" size="15" maxlength="20" /></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td height="29" align="center">好友邮箱：</td>
        <td height="30"><input name="FrMail" type="text" id="FrMail" maxlength="50" /></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td height="29" align="center">您的姓名：</td>
        <td height="30"><input name="MyName" type="text" id="MyName" size="15" maxlength="20" /></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td height="26" align="center">您的邮箱：</td>
        <td height="30"><input name="MyMail" type="text" id="MyMail" /></td>
      </tr>
      <tr align="center" bgcolor="#FFFFFF">
        <td height="23" colspan="2" > 邮件内容</td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td height="56" colspan="2"><br />
          您好!<br />
          我在<a href="<%=WebUrl%>" target="_blank">[<%=WebName%>]</a>上看到一篇标题为<font color="#FF0000"><%=RS("Title")%></font>的信息，希望能对您有所帮助。<br />
          网址为：<a href="<%=ArticleUrl%>" target="_blank"><%=ArticleUrl%></a><br />
          <input name="Content" type="hidden" id="Content" value="&lt;br&gt;您好!&lt;br&gt;我在&lt;a href=<%=WebUrl%> target=_blank&gt;[<%=WebName%>]&lt;/a&gt;上看到一篇标题为&lt;font color=#FF0000&gt;<%=RS("Title")%>&lt;/font&gt;的信息，希望能对您有所帮助。&lt;br&gt;网址为：&lt;a href=<%=ArticleUrl%> target=_blank&gt;<%=ArticleUrl%>&lt;/a&gt;&lt;br&gt;" />
        </td>
      </tr>
      <tr align="center" bgcolor="#FFFFFF">
        <td height="28" colspan="2"><input type="hidden" value="<%=WebName%>" name="SiteName">
            <input type="hidden" name="ID" value="<%=ID%>">
            <input type="hidden" name="m" value="<%=channelid%>">
            <input type="submit" name="Submit" class="fmbtn" value="发　送　邮　件" /></td>
      </tr>
    </table></td>
  </tr>
  </form>
</table>
</body>
</html> 

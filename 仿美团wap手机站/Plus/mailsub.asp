<!--#include file="../conn.asp"-->
<!--#Include file="../ks_cls/kesion.commoncls.asp"-->
<%
Dim KS:Set KS=New PublicCls
 Dim RS,Email,classid,activecode,ClassInfo,mailid,action
 Dim CheckUrl,MailBodyStr,ReturnInfo
 action=KS.S("Action")
 Email=KS.S("Email")
 activecode=KS.DelSQL(KS.S("ActiveCode"))
 mailid=KS.ChkClng(KS.S("id"))
if action="del" then
   if mailid=0 then ks.die "error!"
   set rs=server.CreateObject("adodb.recordset")
   rs.open "select top 1 * from ks_usermail where id=" & mailid & " and activecode='" & activecode &"'",conn,1,1
   if rs.eof and rs.bof then
     rs.close:set rs=nothing
      KS.Die "<script>alert('对不起，删除认证不通过,系统可能不存在此邮箱！');window.close();</script>"
   end if
   email=rs("email")
   rs.close : set rs=nothing
   conn.execute("update ks_usermail set activetf=0 where id=" & mailid)
   'conn.execute("delete from ks_usermail where id=" & mailid)
   KS.Die "<script>alert('恭喜，邮件" & email & "在本站的订阅服务已取消！');window.close();</script>"
elseIf Action="cancel" Then
  Set RS=Server.CreateObject("adodb.recordset")
  RS.Open "SELECT TOP 1 * From KS_UserMail Where Email='" & Email & "'",conn,1,1
  If RS.Eof And RS.Bof Then
    RS.Close :Set RS=Nothing
	KS.AlertHintScript "您输入的邮件不存在，如果不是您的邮件地址，请不要非法操作！"
  End If
  mailid=rs("id")
  activecode=rs("activecode")
  cassid=rs("classid")
  rs.close : set rs=nothing
  
 IF KS.IsNul(ClassID) Then
		ClassInfo= "全部"
 Else
		   ClassID=Replace(ClassID," ","")
		   ClassIDArr=Split(ClassID,",")
		   For I=0 To Ubound(ClassIDArr)
		     If I<>Ubound(ClassIDArr) Then
		      ClassInfo=ClassInfo & KS.C_C(ClassIDArr(i),1) & "，"
			 Else
		      ClassInfo=ClassInfo & KS.C_C(ClassIDArr(i),1) 
			 End If
		   Next
  End If
  
  CheckUrl = Request.ServerVariables("HTTP_REFERER")
  CheckUrl=KS.GetDomain &"plus/mailsub.asp?action=del&id=" &mailid &"&activecode=" & activecode
  MailBodyStr="<strong>在“" & KS.Setting(0) & "”网站的邮件订阅服务取消息确认！</strong><br/>"
  MailBodyStr=MailBodyStr & "原创建的订阅信息如下：<br/><br/>"
  MailBodyStr=MailBodyStr & "订阅邮箱：" & email & "<br/>"
  MailBodyStr=MailBodyStr & "订阅的栏目：<span style='color:blue'>" & ClassInfo & "</span><br/><br/>"
  MailBodyStr=MailBodyStr & "如果要取消订阅，请点击以下链接删除您的订阅服务：<br/>"
  MailBodyStr=MailBodyStr & "<a href='" & CheckUrl & "' target='_blank'>" & CheckUrl &"</a><br/><br/>"
  MailBodyStr=MailBodyStr & "<div style='text-align:right'><strong>说明：</strong>此邮件系统自动发送，不需要回复!</div>"
  ReturnInfo=KS.SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14), "邮件订阅服务取消确认信", Email,KS.Setting(0), MailBodyStr,KS.Setting(11))
  KS.Die "<script>alert('您的取消订阅服务请求已提交，请打开您的邮件完成最后的操作！');location.href='../';</script>"
ElseIf action="active" Then
 If mailid=0 Then KS.Die "error!"
 Set RS=Server.CreateObject("adodb.recordset")
 RS.Open "select top 1 * From KS_UserMail Where ActiveCode='" & ActiveCode & "' and id=" & mailid,conn,1,1
 If RS.Eof And RS.Bof Then
   RS.Close : Set RS=Nothing
   KS.Die "<script>alert('激活操作失败！');window.close();</script>"
 End If
 If RS("ActiveTF")=1 Then
   RS.Close : Set RS=Nothing
   KS.Die "<script>alert('该订阅服务已激活过了，不需要重复激活操作！');window.close();</script>"
 End If
 RS.Close :Set RS=Nothing
 Conn.Execute("Update KS_UserMail Set ActiveTF=1 Where ID=" & mailid)
   KS.Die "<script>alert('恭喜，您的邮件订阅服务已激活，您将会不定期的收到我们的订阅邮件服务，感谢您的支持！');location.href='../';</script>"
 
ElseIf KS.S("Action")="dosave" Then
 Email=KS.S("Email")
 ClassID=KS.S("ClassID")
 If Not KS.IsValidEmail(Email) Then
    KS.AlertHintScript "对不起，您输入的邮件不合法!"
 End If
 activecode=KS.MakeRandom(10)
 Set RS=Server.CreateObject("adodb.recordset")
 RS.Open "Select top 1 * From KS_UserMail Where Email='" & Email &"'",conn,1,3
 If RS.Eof Then
   RS.AddNEW
 End If
   RS("ActiveCode")=activecode
   RS("Email")=Email
   RS("ClassID")=ClassID
   If Not KS.IsNul(KS.C("UserName")) Then
   RS("UserName")=KS.C("UserName")
   RS("IsUser")=1
   Else
   RS("IsUser")=0
   End If
   RS("AddDate")=Now
   RS("ActiveTF")=0
   RS.Update
	RS.Close
	Set RS=Nothing
   mailid=KS.ChkClng(Conn.Execute("Select top 1 id From KS_UserMail Where Email='" & Email &"'")(0))
 
 
 IF KS.IsNul(ClassID) Then
		ClassInfo= "全部"
 Else
		   ClassID=Replace(ClassID," ","")
		   Dim ClassIDArr:ClassIDArr=Split(ClassID,",")
		   For I=0 To Ubound(ClassIDArr)
		     If I<>Ubound(ClassIDArr) Then
		      ClassInfo=ClassInfo & KS.C_C(ClassIDArr(i),1) & "，"
			 Else
		      ClassInfo=ClassInfo & KS.C_C(ClassIDArr(i),1) 
			 End If
		   Next
  End If
 
 
 
  CheckUrl = Request.ServerVariables("HTTP_REFERER")
  CheckUrl=KS.GetDomain &"plus/mailsub.asp?action=active&id=" &mailid &"&activecode=" & activecode
  MailBodyStr="<strong>请确认您在“" & KS.Setting(0) & "”网站的邮件订阅服务！</strong><br/>"
  MailBodyStr=MailBodyStr & "您创建的订阅信息如下：<br/><br/>"
  MailBodyStr=MailBodyStr & "订阅邮箱：" & email & "<br/>"
  MailBodyStr=MailBodyStr & "订阅的栏目：<span style='color:blue'>" & ClassInfo & "</span><br/><br/>"
  MailBodyStr=MailBodyStr & "请点击以下链接激活您的订阅请求：<br/>"
  MailBodyStr=MailBodyStr & "<a href='" & CheckUrl & "' target='_blank'>" & CheckUrl &"</a><br/><br/>"
  MailBodyStr=MailBodyStr & "<div style='text-align:right'><strong>说明：</strong>此邮件系统自动发送，不需要回复!</div>"
  
  ReturnInfo=KS.SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14), "邮件订阅服务激活邮件", Email,KS.Setting(0), MailBodyStr,KS.Setting(11))

 
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=KS.Setting(1)%> - 邮件订阅服务</title>
<style>
body{padding:0px;margin:0px;font-size:12px;text-align:center}
#warp{width:960px;margin: 0 auto; border:1px solid #ccc;}
#cphead2010{ background:#000; height: 32px; overflow: hidden; color: #c7e0ff; }
#cphead2010 a{ color: #c7e0ff; font-size: 12px; text-decoration: none; }
#cphead2010 a:hover{ text-decoration: underline; }
#cphead2010 *{ padding: 0; margin: 0; font-size: 12px; }
#cphead2010 img{ border: 0; }
#cphead2010 .cpnav{ height: 32px; }
#cphead2010 .cpnav dd{ float: left; padding: 8px 15px 0 15px; line-height: 20px; }
#cphead2010 .cpnav dd.load{ float: right; }
.box{}
.boxleft{width:610px;float:left;border-right:1px solid #cccccc;}
.boxleft .box1{font-size:14px;color:#fff;
word-spacing:8px;letter-spacing: 4px;font-weight:bold;padding-top:60px;height:70px;background:url(images/box1.gif) no-repeat;}
.boxleft .box2{text-align:left;padding-left:30px;background:url(images/box2.gif) repeat-y;}
.boxleft .box3{height:70px;background:url(images/box3.gif) no-repeat;}
.boxleft .email{width:208px;height:31px;line-height:31px;background:url(images/email.gif) no-repeat;border:0px;padding-left:5px;}
.boxright{text-align:left;width:300px;float:right;padding:20px}
</style>

</head>

<body>
<div id="warp">

<div id="cphead2010">
	<dl class="cpnav">
	<dd><a href="/" target="_blank">首页</a> - <a href="../ask/" target="_blank">问答</a> - <a href="../club/" target="_blank">论坛</a> - <a href="../user" target="_blank">会员</a> - <a href="../space" target="_blank">博客</a></dd>
	<dd class="load"><a href="../user/login">登录</a><span>|</span><a href="../user/reg/" target="_blank">注册</a></dd>
	</dl>
</div>

<form name="myform" action="mailsub.asp" method="post" />
<div class="box">
	<div class="boxleft">
	  <div class="box1">欢迎使用本站邮件订阅服务!</div>
	  <div class="box2">
	  
	  <%If KS.S("Action")="dosave" Then%>
	   <img src='../user/images/regok.jpg' align='left' style="margin:20px"/> <strong>恭喜，您的订阅已创建！</strong><br/><br/>
		订阅邮件：<span style='color:red'><%=KS.CheckXSS(Email)%></span><br/>
		订阅的栏目：<%=ClassInfo%>
		<br/><br/>
		请注意收取您的确认邮件！您需要单击确认邮件中的链接，确认您的请求，<br/>您本次的订阅操作才会生效！
	  <%Else%>
		  <input type="hidden" name="action" id="action" value="dosave" />
		  <img src="images/img13.gif" align="absmiddle"/>
		  <input type="text" name="email" id="email" class="email" value="<%=request("email")%>" maxlength="28"/>
		  <input type="submit" value="取消订阅" onclick="if(document.myform.email.value==''){alert('请输入您的邮件!');return false;}document.getElementById('action').value='cancel';"/>
		  <table border="0" width="100%" align="center">
		   <tr>
			<td colspan="10" height="40" align="left"><strong>请选择您感兴趣的栏目，如果不选将自动从以下栏目中挑选最新信息发送给您：</strong></td>
		   </tr>
		  <%
		  KS.LoadClassConfig()
		  Dim Node,I
		  I=0
		  For Each Node In Application(KS.SiteSN&"_class").DocumentElement.SelectNodes("class[@ks14=1 and @ks21=1]")
		   If I=0 Then
			KS.Echo "<tr>"
		   ElseIf I Mod 5=0 Then
			KS.Echo "</tr><tr>"
		   End If
		   KS.Echo "<td><label><input type='checkbox' name='classid' value='" & Node.SelectSingleNode("@ks0").text &"'/>" & Node.SelectSingleNode("@ks1").text & "</label></td>"
		   I=I+1
		  Next
		  %>
		   </tr>
		  </table>
		  <br/>
		  <input type="image" onclick="return(CheckForm())" src="images/img05.jpg"/>
		<%End If%>
	  </div>
	  <div class="box3"></div>
	</div>
	  </form>
	
	<div class="boxright">
	  <p><strong>为什么收不到新闻邮件？</strong></p>
		<p>可能的原因：<br />
		  1.没有激活订阅，订阅新闻以后需要您在72小时内到邮箱中激活才能接收到新闻。<br />
		  2.由于网络,邮件过滤等原因,个别邮箱可能收不到邮件订阅，请更换邮箱再次订阅。</p>
		<p><strong>如何取消邮件订阅？</strong></p>
		<p>有如下几种方式可选：<br />
		  1.在左方的表框内输入您所订阅的内容和邮件地址，按“取消订阅”即可。 <br />
		  2.在您接收到的邮件下方有“取消此订阅”链接，直接点击该链接可取消订阅。</p>
		<p><strong>如何重新订阅喜爱的栏目？</strong></p>
		<p>在左方重新输入您的订阅邮件，并选择自己喜爱的栏目重新提交订阅，然后进您的邮件点激活链接重新激活即可。</p>
		  
		  
		<p><strong>邮件订阅是否收费？</strong></p>
		<p>订阅服务是免费的，其内容由本站为您提供。 </p>
	</div>
</div>


</div>
<div style="clear:both;color:#333333;padding:16px;">
 <%=KS.Setting(18)%>
</div>
<script>

function CheckForm()
{
	var form = document.myform;
	var email = form.email.value;
	if (email == '') {
		alert('请填写订阅邮箱！');
		form.email.focus();
		return false;
	}
	if (checkMail(email) == false) {
		return false;
	}
	return true;
}

// 检查邮件地址是否包含非法字符
function checkMail(email) {
	var filter = /^([a-zA-Z0-9_\.\-])+\@(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9]{2,4})+$/;
	var info = '抱歉，邮件地址只能由英文字母a～z(不区分大小写)' +
				'、数字0～9、下划线_，减号-，点.组成，' +
				'不能有汉字及引号、大于号、小于号等特殊字符。' +
				'例：abc@hotmail.com。请检查您输入的邮件地址。';
	if (!filter.test(email)) {
		alert(info);
		return false;
	}
	return true;
}

</script>
</body>
</html>

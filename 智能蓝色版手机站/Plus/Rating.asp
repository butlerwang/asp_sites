<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../plus/md5.asp"-->
<% 

response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Response.CharSet="utf-8" 

Dim KS:Set KS=New PublicCls
Set KSUser = New UserCls
Call KSUser.UserLoginChecked()
Dim ChannelID,InfoID,RS,ProjectID
Dim totalPut, MaxPerPage,PageNum,SqlStr,N

ProjectID=KS.ChkClng(KS.S("ProjectID"))
ChannelID=KS.Chkclng(KS.S("ChannelID"))
InfoID=KS.ChkClng(KS.S("InfoID"))
DomainStr=KS.GetDomain
Select Case KS.S("Action")
 Case "showzcj" Call showzcj()
 Case "DoSave"  Call DoSave()
 Case Else 
   Call ShowComment()
 End Select
 Set KS=Nothing
 Set KSUser=Nothing
 


Function GetRegRnd()
		  Dim QuestionArr:QuestionArr=Split(KS.GetCurrQuestion(162),vbcrlf)
		  Dim RandNum,N: N=Ubound(QuestionArr)
          Randomize
          RandNum=Int(Rnd()*N)
          GetRegRnd=RandNum
End Function
Function GetRegQuestion(ByVal RndReg)
		  Dim QuestionArr:QuestionArr=Split(KS.GetCurrQuestion(162),vbcrlf)
		  GetRegQuestion=QuestionArr(rndReg)
End Function
Function GetRegAnswerRnd(ByVal RndReg)
		  GetRegAnswerRnd=md5(rndReg,16)
End Function
		
Sub showzcj()
  if KS.C("UserName")<>"" Then
    KS.Echo "document.commentform.cname.value='" & KS.C("UserName") &"';" &vbcrlf
  End If
  Dim RS:Set RS=Conn.Execute("select top 1 * From KS_MoodProject Where ID=" & ProjectID)
  If Not RS.Eof Then
    If RS("VerifyCodeTF")<>"1" Then
	 KS.echo "jQuery('#showverifycode').hide();" & vbcrlf
	End If
	If RS("ZCJTF")<>"1" Then
	 KS.echo "jQuery('#showzcj').hide();" & vbcrlf
	End If
  End If
  RS.Close
  Set RS=Nothing

          Dim RndReg:rndReg=GetRegRnd()
		 response.write "document.write('问题：" & GetRegQuestion(RndReg) & "');" &vbcrlf
		 response.write "document.write('<br/>您的答案：<input type=""text"" id=""QuestionAnswer"" name=""a" & GetRegAnswerRnd(RndReg) & """>');"&vbcrlf
End Sub

Sub DoSave()
 If ChannelID=0 Or InfoID=0 Then KS.Die "error!"
 %>
 <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv=Content-Type content="text/html; charset=utf-8">
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" /> 
<%
If EnabledSubDomain Then
 response.write "<script>document.domain=""" & RootDomain &""";</script>" &vbcrlf
end if
%>
<script type="text/javascript" src="../ks_inc/jquery.js"></script>
<script type="text/javascript" src="../ks_inc/lhgdialog.js"></script>

</head>
<body>
 <%
  dim from:from =KS.S("From")
  dim backstr
  Dim LoginTF:LoginTF=KSUser.UserLoginChecked
  if from="script" then
    backstr="history.back();"
  else
    backstr="parent.loading.style.display='none';"
  end if
  
 '商品点评，只有购买过的才能点评
 if ChannelID=5 Then
   if LoginTF=false Then
     KS.Die "<script>$.dialog.tips('对不起，请登录!',1,'error.gif',function(){parent.loading.style.display='none';});</script>"
   End If
   if conn.execute("select top 1 a.id from ks_order a inner join ks_orderitem b on a.orderid=b.orderid where b.proid=" & infoid &" and a.status=2 and a.username='" & KSUser.UserName & "'").eof then
  KS.Die "<script>$.dialog.tips('对不起，只有成功购买过本商品的用户才可以点评!',1,'error.gif',function(){" & backstr & "});</script>"
   end if
 End If
 
 Dim verific
 Dim AnounName:AnounName=KS.S("cname")
 If KS.IsNul(AnounName) Then KS.Die "<script>$.dialog.tips('昵称必须输入!',1,'error.gif',function(){" & backstr & "});</script>"
 Dim title:title=KS.S("title")
  If KS.IsNul(title) Then KS.Die "<script>$.dialog.tips('点评标题必须输入!',1,'error.gif',function(){" & backstr & "});</script>"
 Dim content:content=KS.S("content")
 If KS.IsNul(content) Then KS.Die "<script>$.dialog.tips('点评内容必须输入!',1,'error.gif',function(){" & backstr & "});</script>"
 Dim VerifyCode:VerifyCode=KS.S("VerifyCode")
 Set RS=Server.CreateObject("ADODB.RECORDSET")
 RS.Open "select top 1 * From KS_MoodProject Where ID=" & ProjectID,conn,1,1
 If RS.Eof And RS.Bof Then
   RS.Close:Set RS=nothing
   KS.Die "<script>$.dialog.tips('出错了，评分项目不存在!',1,'error.gif',function(){" & backstr & "});</script>"
 End If
 If RS("Status")=0 Then
		KS.Die "<script>$.dialog.tips('对不起，该点评项已锁定!',1,'error.gif',function(){" & backstr & "});</script>"
 End If
 If RS("VerifyCodeTF")="1" Then
  If lcase(Trim(Request("Verifycode")))<>lcase(Trim(Session("Verifycode"))) Then
		KS.Die "<script>$.dialog.tips('验证码有误，请重新输入!',1,'error.gif',function(){ $('#VerifyCode',parent.document).focus();" & backstr & "});</script>"
  End IF
 End If
 If RS("ZCJTF")="1" Then
   '检查注册回答问题
		  Dim CanReg,n
		        CanReg=false
				 For N=0 To Ubound(Split(KS.GetCurrQuestion(162),vbcrlf))
				   If Trim(Request("a" & MD5(n,16)))<>"" Then
					  If trim(Lcase(Request("a" & MD5(n,16))))<>trim(Lcase(Split(KS.GetCurrQuestion(163),vbcrlf)(n))) Then
					   KS.Die "<script>$.dialog.tips('对不起,防发帖机问题的回答不正确!',1,'error.gif',function(){$('#QuestionAnswer',parent.document).focus();" & backstr & "});</script>"
					   CanReg=false
					  Else
					   CanReg=True
					  End If
				   End If
				 Next
			 If CanReg=false Then KS.Die "<script>$.dialog.tips('对不起,防发帖机问题的回答不正确!',1,'error.gif',function(){$('#QuestionAnswer',parent.document).focus(); "& backstr & "});</script>"
 End If
 If RS("onlyuser")=1 And LoginTF=false Then KS.Die "<script>$.dialog.tips('对不起，点评只对会员开放!',1,'error.gif',function(){"& backstr & "});</script>"
 If RS("UserOnce")=1 And KS.ChkClng(Conn.Execute("select count(1) from KS_Comment Where UserName='" & KSUser.UserName & "' and channelID=" & channelid & " and infoid=" & InfoID &" and projectid="& KS.ChkClng(ProjectID))(0))<>0 Then KS.Die "<script>$.dialog.tips('对不起，只能点评一次!',1,'error.gif',function(){"& backstr & "});</script>"
 If RS("TimeLimit")="1" then
		 if now<RS("StartDate") then KS.Die "<script>$.dialog.tips('对不起，点评还没有开放!',1,'error.gif',function(){"& backstr & "});</script>"
		 if now>RS("ExpiredDate") then KS.Die "<script>$.dialog.tips('对不起，点评已结束!',1,'error.gif',function(){"& backstr & "});</script>"
 End If
 if RS("AllowGroupID")<>"" then
	if KS.FoundInArr(RS("AllowGroupID"),KSUser.groupid,",")=false then  KS.Die "<script>$.dialog.tips('对不起，您所在的用户组没有点评的权限!',1,'error.gif',function(){"& backstr & "});</script>"
 end if
 Dim IsVerify:IsVerify=RS("IsVerify")
 If IsVerify=1 Then
  verific=0
 Else
  verific=1
 End If
 RS.Close
 RS.Open "select top 1 * From KS_Comment Where 1=0",conn,1,3
 RS.AddNew
  RS("ProjectID")=ProjectID
  RS("ChannelID")=ChannelID
  RS("InfoID")=InfoID
  RS("UserIP")=KS.GetIP
  RS("Title")=Title
  RS("Content")=Content
  RS("AddDate")=Now
  If KS.C("UserName")="" Then
   RS("UserName")=AnounName
   RS("Anonymous")=1
  Else
   RS("UserName")=KSUser.UserName
   RS("Anonymous")=0
  End If
  For N=0 To 10
    RS("M" & N)=KS.ChkClng(Request("score" & N))
  Next
  RS("Verific")=Verific
  RS.Update
  RS.Close
  If Verific=1 Then

   dim backurl:backurl=""
    If KS.C_S(Channelid,7)=1 or KS.C_S(ChannelID,7)=2 Then
					 Dim KSRObj:Set KSRObj=New Refresh
					 RS.Open "select top 1 * From " & KS.C_S(ChannelID,2) & " Where ID=" & InfoID,Conn,1,1
					 Dim DocXML:Set DocXML=KS.RsToXml(RS,"row","root")
				     Set KSRObj.Node=DocXml.DocumentElement.SelectSingleNode("row")
					  KSRObj.ModelID=ChannelID
					  KSRObj.ItemID = KSRObj.Node.SelectSingleNode("@id").text 
					  backurl=KS.GetItemURL(ChannelID,KSRObj.Node.SelectSingleNode("@tid").text,KSRObj.Node.SelectSingleNode("@id").text,KSRObj.Node.SelectSingleNode("@fname").text)
					  Call KSRObj.RefreshContent()
					  Set KSRobj=Nothing
					RS.Close
	End If
  End If
  Set RS=Nothing
  If EnabledSubDomain Then
	Response.Cookies(KS.SiteSn).domain=RootDomain					
Else
    Response.Cookies(KS.SiteSn).path = "/"
End If

  Response.Cookies(KS.SiteSn)("mood_"&ChannelID&"_"&InfoID)="ok"
  if from="script" then
    if backurl="" then backurl=request.ServerVariables("HTTP_REFERER")
    KS.Die "<script>$.dialog.tips('恭喜，点评提交成功!',2,'success.gif',function(){location.href='" & backurl &"';});</script>"
  else
  KS.Die "<script>$.dialog.tips('恭喜，点评提交成功!',1,'success.gif',function(){top.location.reload();});</script>"
  end if
End Sub

Sub ShowComment()
  MaxPerPage=10
  If ChannelID=0 Or InfoID=0 Then KS.Die "error!"
  Dim TemplateID,KSR,Template,Title,PageTitle,PageTitleArr,IsRewrite
  Dim RS:Set RS=Conn.Execute("Select top 1 ProjectContent,TemplateID,IsRewrite From KS_MoodProject Where ID=" & ProjectID)
  If Not RS.Eof Then
    IsRewrite=RS("IsRewrite")
    TemplateID=RS("TemplateID") 
    Set KSR=New refresh
	Template=KSR.LoadTemplate(TemplateID)
	Dim NN,Sstr,CommentStr,ProjectContentArr:ProjectContentArr=Split(RS(0),"$$$")
	RS.Close
	If KS.C_S(ChannelID,6)="1" Then
	 Set RS=Conn.Execute("Select top 1 Title,PageTitle From " & KS.C_S(ChannelID,2) & " Where ID=" & InfoID)
	Else
	 Set RS=Conn.Execute("Select top 1 Title,'' as PageTitle From " & KS.C_S(ChannelID,2) & " Where ID=" & InfoID)
	End If
	If Not RS.Eof Then
	 Title=RS(0)
	 PageTitle=RS(1)
	  If Not KS.IsNul(PageTitle) Then
			 PageTitleArr=Split(PageTitle,"§")
			 If CurrentPage-1<=Ubound(PageTitleArr) Then
				Template=Replace(Template,"{$GetTitle}",PageTitleArr(CurrentPage-1))
			 Else
			    Template=Replace(Template,"{$GetTitle}",Title)
			 End If
	  ElseIF Currentpage>1 Then
	    Template=Replace(Template,"{$GetTitle}",Title& "(" & currentpage & ")")
	  Else
	    Template=Replace(Template,"{$GetTitle}",Title)
	  End If
	End If
	RS.Close
    Set RS=Conn.Execute("select * From KS_Comment Where ProjectID=" & ProjectID & " and ChannelID=" & ChannelID & " and InfoID=" & InfoID & " and verific=1 order by id desc")
	 
	 DiM Floor:Floor=Conn.Execute("Select count(1) From KS_Comment Where ProjectID=" & ProjectID & " and ChannelID=" & ChannelID & " and InfoID=" & InfoID & " and verific=1")(0)
	Template=Replace(Template,"{$GetInfoTitle}",Title)
	Template=Replace(Template,"{$GetCommentNum}",Floor)
	Dim AvgStr:AvgStr=""
	For NN=0 To Ubound(ProjectContentArr)
	 if Split(ProjectContentArr(NN),"|")(0)<>"" then
	   AvgStr=AvgStr & Split(ProjectContentArr(NN),"|")(0) & ":<span style='color:red'>" & round(conn.execute("select avg(m" & nn&") from ks_comment where channelid=" & channelid & " and infoid=" & infoid &" and projectid=" & projectid)(0),2) &"</span> 分&nbsp;&nbsp;"
	 End If
	Next
	Template=Replace(Template,"{$GetAgvStr}",AvgStr)
						 If Not RS.Eof Then
							 CommentStr=LFCls.GetXMLByNoCache("comments","/posttemplate/label","[@name='show']")
							  TotalPut = Floor
							  Floor=Floor-(CurrentPage-1)*MaxPerPage
							  if (TotalPut mod MaxPerPage)=0 then
								  PageNum = TotalPut \ MaxPerPage
								else
									PageNum = TotalPut \ MaxPerPage + 1
								end if
							  
							  If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
                               RS.Move (CurrentPage - 1) * MaxPerPage
							  Else
									CurrentPage = 1
							 End If
							 Dim I:I=0
							 Do While Not RS.Eof
							   Dim StarStr:StarStr=""
							   For NN=0 To Ubound(ProjectContentArr)
							    if Split(ProjectContentArr(NN),"|")(0)<>"" then
							     StarStr=StarStr & Split(ProjectContentArr(NN),"|")(0) & "：<img src='" & KS.GetDomain & "images/star/star-" & RS("M"&nn)& ".jpg'/>&nbsp;&nbsp;"
								end if
							   Next
							   Dim ContentStr:ContentStr=rs("content")
							   If Not KS.IsNul(rs("ReplyContent")) Then
							    ContentStr=ContentStr &"<div style='margin:10px;padding:4px;color:red;border:1px dashed #ccc;background:#FFFFEE;'>管理员回复：" &rs("ReplyContent") &"</div>" 
							   End If
							   Dim UserIP:UserIP=left(rs("userip"), InStrRev(rs("userip"), ".")) & "*"
							   Sstr=Sstr & Replace(Replace(Replace(Replace(Replace(Replace(Replace(commentstr,"{$Title}",rs("title")),"{$Content}",ContentStr),"{$UserName}",rs("UserName")),"{$PostTime}",rs("AddDate")),"{$UserIP}",UserIP),"{$ShowStar}",StarStr),"{$Floor}",Floor)
							   Floor=Floor-1
							   I=I+1
							   If I>=MaxPerPage Then Exit Do
							  RS.MoveNext
							 Loop
					 RS.Close
					 Set RS=Nothing
			 End If
	End If 
	Template=Replace(Template,"{$CommentList}",Sstr)
	If KS.ChkClng(IsRewrite)<>1 Then
	Template=Replace(Template,"{$ShowPage}",KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,false))
	Else
	Template=Replace(Template,"{$ShowPage}",ShowPage())
	End If
	
	
	Sstr="<table border=""0"" cellspacing=""0"" cellpadding=""0"">"
	KK=0
						   For NN=0 To Ubound(ProjectContentArr)
						    If Split(ProjectContentArr(kk),"|")(0)<>"" Then
								 Sstr=Sstr & "<tr>"
								 For MM=1 To 2
								 if Split(ProjectContentArr(kk),"|")(0)<>"" then
								 Sstr=Sstr & " <td width=""270"" valign=""top""><input type=""hidden"" name=""score" & KK &""" id=""score" & kk & """>"
								 Sstr=Sstr & "  <div class=""rate-text"">"  & Split(ProjectContentArr(kk),"|")(0) & "：</div>"
								 Sstr=Sstr & "		<div id=""Commentdemo" & KK & """ class=""add_comment_start""></div><div id=""score" & KK & "_desc"" class=""add_comment_start_desc""></div>"
								 Sstr=Sstr & "<script language=""javascript"">"
								 Sstr=Sstr  &"	$('#Commentdemo" & KK & "').rater(null, {maxvalue:5,curvalue:0}, function(el , value) {setRateValue(value, ""score" & KK & """);});"
								 Sstr=Sstr & "</script>"
								 Sstr=Sstr & "</td>"
								 end if
								 KK=KK+1
								 If KK>=Ubound(ProjectContentArr) Then Exit For
							   Next
							   Sstr=Sstr & "</tr>"
						   End If
						   If KK>=Ubound(ProjectContentArr) Then Exit For
						  Next
						  Sstr=Sstr & "</table>"
						   
						  CommentStr=LFCls.GetXMLByNoCache("comments","/posttemplate/label","[@name='post']")

						   CommentStr=Replace(CommentStr,"{$GetSiteUrl}",KS.GetDomain)
						   CommentStr=Replace(CommentStr,"{$ProjectID}",ProjectID)
						   CommentStr=Replace(CommentStr,"{$ChannelID}",ChannelID)
						   CommentStr=Replace(CommentStr,"{$ItemID}",InfoID)
						   CommentStr=Replace(CommentStr,"{$ScoreItem}",Sstr)
						   CommentStr=Replace(CommentStr,"{$Title}",Title)
						   If KS.C_S(ChannelID,14)<>"0" Then
						   CommentStr=Replace(CommentStr,"{$MaxLen}",KS.C_S(ChannelID,14))
						   CommentStr=Replace(CommentStr,"{$MaxLenNum}",KS.C_S(ChannelID,14))
						   Else
						   CommentStr=Replace(CommentStr,"{$MaxLen}","不限制")
						   CommentStr=Replace(Replace(CommentStr,"{$MaxLenNum}",0),"{$DisplayZS}","display:none;")
						   End If
	
	Template=Replace(Template,"{$GetWriteComment}",CommentStr)
	
	
	Template=KSR.KSLabelReplaceAll(Template) 
	Set KSR=Nothing
	KS.Die Template
End Sub

'伪静态分页
Public Function ShowPage()
		           Dim I, pageStr
				   pageStr= ("<div id=""fenye"" class=""fenye""><table border='0' align='right'><tr><td>")
					if (CurrentPage>1) then pageStr=PageStr & "<a href=""rating-" & channelid & "-" & infoid &"-" & projectid & "-" & CurrentPage-1 & ".html"" class=""prev"">上一页</a>"
				   if (CurrentPage<>PageNum) then pageStr=PageStr & "<a href=""rating-" & channelid & "-" & infoid & "-" & projectid & "-" & CurrentPage+1 & ".html"" class=""next"">下一页</a>"
				   pageStr=pageStr & "<a href=""rating-" & channelid &"-" & infoid & "-" & projectid & "-1.html"" class=""prev"">首 页</a>"
				 
					Dim startpage,n,j
					 if (CurrentPage>=7) then startpage=CurrentPage-5
					 if PageNum-CurrentPage<5 Then startpage=PageNum-10
					 If startpage<0 Then startpage=1
					 n=0
					 For J=startpage To PageNum
						If J= CurrentPage Then
						 PageStr=PageStr & " <a href=""#"" class=""curr""><font color=red>" & J &"</font></a>"
						Else
						 PageStr=PageStr & " <a class=""num"" href=""rating-" & channelid & "-" & infoid & "-" & projectid & "-" & J & ".html"">" & J &"</a>"
						End If
						n=n+1 : if n>=10 then exit for
					 Next
					
					 PageStr=PageStr & " <a class=""next"" href=""rating-" & channelid & "-" & infoid &"-" & projectid & "-" & PageNum & ".html"">末页</a>"
					 pageStr=PageStr & " <span>共" & totalPut & "条记录,分" & PageNum & "页</span></td></tr></table>"
				     PageStr = PageStr & "</td></tr></table></div>"
			         ShowPage = PageStr
End Function
%>
 

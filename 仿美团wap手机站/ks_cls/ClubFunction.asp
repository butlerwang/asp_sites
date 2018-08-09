<%
Sub Echo(sStr)
			If Immediate Then
				Response.Write    sStr
				Response.Flush()
			Else
				Templates    = Templates&sStr 
			End If 
End Sub 
		
Public Sub Scan(sTemplate)
			Dim iPosLast, iPosCur
			iPosLast    = 1
			While True 
				iPosCur    = InStr(iPosLast, sTemplate, "{@")
				If iPosCur>0 Then
					Echo    Mid(sTemplate, iPosLast, iPosCur-iPosLast)
					iPosLast    = Parse(sTemplate, iPosCur+2)
				Else 
					Echo    Mid(sTemplate, iPosLast)
					Exit Sub  
				End If 
		   Wend 
End Sub 

Sub GetClubPopLogin(ByRef FileContent)
 If Instr(FileContent,"{#GetClubPopLogin}")=0 Then Exit Sub
 Dim Str,QQEnable,AlipayEnable,SinaEnable
 If KS.IsNul(KS.C("UserName")) And KS.IsNul(KS.C("PassWord")) Then
   Str=str & "<form method=""post"" autocomplete=""off"" id=""loginform"" action=""" & KS.GetDomain & "user/checkuserlogin.asp"">"
   If KS.ChkClng(KS.Setting(34))=1 Then Str=str & "<div class=""fastlg1"">" Else Str=str & "<div class=""fastlg"">"
   Dim  XslDoc:Set XslDoc = KS.InitialObject("Msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
   If XslDoc.Load(request.ServerVariables("APPL_PHYSICAL_PATH") &"api/api.config") Then
      QQEnable    =cbool(XslDoc.documentElement.selectSingleNode("rs:data/z:row").getAttribute("api_qqenable"))
	  AlipayEnable=cbool(XslDoc.documentElement.selectSingleNode("rs:data/z:row").getAttribute("api_alipayenable"))
	  SinaEnable  =cbool(XslDoc.documentElement.selectSingleNode("rs:data/z:row").getAttribute("api_sinaenable"))
	  If QQEnable Or AlipayEnable Or SinaEnable Then
	    Str=Str & "<div class=""l""><strong>账号通</strong><br/>"
		If QQEnable Then Str=Str & "<a title=""使用qq账号登录"" href=""" & KS.GetDomain & "api/qq/redirect_to_login.asp""><img src=""" &KS.GetDomain & "images/default/icon_qq.png"" align=""absmiddle""/></a>&nbsp;&nbsp;"
		If SinaEnable Then Str=Str & "<a title=""使用新浪微博账号登录"" href=""" & KS.GetDomain & "api/sina/redirect_to_login.asp""><img src=""" &KS.GetDomain & "images/default/icon_sina.png"" align=""absmiddle""/></a>&nbsp;&nbsp;"
		If AlipayEnable Then Str=Str & "<a title=""使用支付宝登录"" target=""_blank"" href=""" & KS.GetDomain & "api/alipay/alipay_auth_authorize.asp""><img src=""" &KS.GetDomain & "images/default/icon_alipay.png"" align=""absmiddle""/></a>"
		'Str=Str & "<div style=""margin:3px"">内容互通，快速登录</div>"
		Str=Str & "</div>"
	 End If
   End If
   Set XslDoc=Nothing
   str=str& "<div class=""r""><p><a href=""" & KS.GetDomain & "user/reg/"">注册</a></p><p><a href=""" & KS.GetDomain & "user/getpassword.asp"">找回密码</a>"
      
   str=str & "</p></div><div class=""c""><p>账号&nbsp;<input type=""text"" style=""color:#999"" onfocus=""if(this.value=='UID/用户名/Email'){this.value='';}"" onblur=""if(this.value==''){this.value='UID/用户名/Email';}"" value=""UID/用户名/Email"" name=""username"" size=""13"" id=""username"" autocomplete=""off"" class=""textbox"" tabindex=""1"" />&nbsp;<label><input type=""checkbox""  name=""ExpiresDate"" value=""1"" />记住</label></p><p><table cellspacing=""0"" cellpadding=""0"" border=""0""><tr><td nowrap>密码&nbsp;</td><td>"
  If KS.ChkClng(KS.Setting(34))=1 Then
   str=str &"<input type=""password"" name=""password"" size=""6"" id=""password"" class=""textbox"" autocomplete=""off"" tabindex=""2"" /></td><td style=""padding-left:1px""><script>writeVerifyCode(""" & KS.GetDomain & """,0,""textbox verificcode"")</script></td>"
  Else
   str=str &"<input type=""password"" name=""password"" size=""13"" id=""password"" class=""textbox"" autocomplete=""off"" tabindex=""2"" /></td>"
  End If
  Str=Str  &"<td nowrap>&nbsp;<input type=""submit"" onclick=""return(ChkLogin(" & KS.ChkClng(KS.Setting(34)) & "));"" value=""登录"" class=""btn"" /></td></tr></table></p></div></div></form>"
 Else

   Dim GetMailTips,MyMailTotal,RS
   If DataBaseType=1 Then
			    Dim Cmd : Set Cmd = Server.CreateObject("ADODB.Command")
				Set Cmd.ActiveConnection=conn
				Cmd.CommandText="KS_CountMessages"
				Cmd.CommandType=4	
				CMD.Prepared = true 
				Cmd.Parameters.Append cmd.CreateParameter("@username",200,1,220)
				Cmd("@username")=KS.C("UserName")
				Set Rs=Cmd.Execute
				MyMailTotal=RS(0)
				Set Cmd=Nothing
				Set RS=Nothing
   Else
	   MyMailTotal=GCls.Execute("Select Count(ID) From KS_Message Where Incept='" &KS.C("UserName") &"' And Flag=0 and IsSend=1 and delR=0")(0)
   End If
   IF MyMailTotal>0 Then 
	  GetMailTips="<span style=""color:red"">" & MyMailTotal & "</span><bgsound src=""" & KS.GetDomain & "User/images/mail.wav"" border=0>"  
   Else
	  GetMailTips=0
   End If
    Dim KSUser:Set KSUser=New UserCls
	KSUser.UserLoginChecked
	dim userfacesrc:userfacesrc=KSUser.GetUserInfo("userface")
	if KS.IsNul(userfacesrc) then userfacesrc="../Images/Face/boy.jpg"
	if left(userfacesrc,1)<>"/" and lcase(left(userfacesrc,4))<>"http" then userfacesrc="../" & userfacesrc
   Str="<div style='float:left;' class=""avatar48""><a target=""_blank"" title=""进入我的空间"" href=""" & KS.GetSpaceUrl(KSUser.GetUserInfo("UserID")) & """><img onerror=""this.src='"& KS.GetDomain & "images/face/boy.jpg';"" src='" & userfacesrc & "' width='40' height='40' align='left'/></a></div> &nbsp;您好！<a href=""" & KS.GetDomain & "user/""><span style='color:red'>" & KS.C("UserName") & "</span></a> 欢迎来到会员中心!<br/>&nbsp;积分：<span style='color:green'>" & KSUser.GetUserInfo("Score") & "</span> 分 帖子：<span style='color:green'>" & KSUser.GetUserInfo("postnum") & "</span> 帖 精华：<span style='color:green'>" & KSUser.GetUserInfo("besttopicnum") & "</span> 帖<br/>【<a href='" & KS.GetDomain & "user/'>会员中心</a>】【<a href='" & KS.GetDomain & "user/user_mytopic.asp?action=fav'>收藏夹</a>】【<a href='" & KS.GetDomain & "user/user_Message.asp?action=inbox'>短消息"&GetMailTips&"</a>】【<a href='" & KS.GetDomain & "User/UserLogout.asp'>退出</a>】"
    If KS.ChkClng(KS.U_S(KSUser.GroupID,8))>0 and KS.ChkClng(KS.U_S(KSUser.GroupID,9))>0 And datediff("n",KSUser.GetUserInfo("LastLoginTime"),now)>=KS.ChkClng(KS.U_S(KSUser.GroupID,8)) then '判断积分奖励时间
     str=str & "<script>popShowMessage('" & KS.U_S(KSUser.GroupID,8) & "分钟后重新登录，奖励积分 +" & KS.U_S(KSUser.GroupID,9) & "分！');</script>"
	 Call KS.ScoreInOrOut(KSUser.UserName,1,KS.ChkClng(KS.U_S(KSUser.GroupID,9)),"系统",KS.ChkClng(KS.U_S(KSUser.GroupID,8)) & "分钟后,重新登录奖励获得",0,0)
	  Conn.Execute("Update KS_User Set LastLoginTime=" & SQLNowString & " Where UserName='" & KSUser.UserName & "'")
	  If IsObject(Session(KS.SiteSN&"UserInfo")) Then Session(KS.SiteSN&"UserInfo").DocumentElement.SelectSingleNode("row").SelectSingleNode("@lastlogintime").Text=now
	ElseIf Not KS.IsNUL(Session("PopTips")) Then
     str=str & "<script>$(document).ready(function(){popShowMessage('" & Session("PopTips") & "！');});</script>"
	 Session("PopTips")=""
	End if
	Set KSUser=Nothing
    Str="<div class=""fastlgs"">" & str & "</div>"
 End If
   FileContent=Replace(FileContent,"{#GetClubPopLogin}",str)
End Sub

'取得所有置顶帖子
Sub LoadTopTopic()
  If Not IsObject(Application(KS.SiteSN &"TopXML")) Then
   MustReLoadTopTopic
  End If
End Sub
Sub MustReLoadTopTopic()
	  Dim ListTopicFields:ListTopicFields="ID,UserName,UserID,Subject,AddTime,Verific,LastReplayUser,LastReplayUserID,LastReplayTime,TotalReplay,BoardID,Hits,IsPic,IsTop,IsBest,PostType,AnnexExt,CategoryId,ShowScore" rem 主题列表用到的字段
	  Dim RS:Set RS=Conn.Execute("Select top 500 " & ListTopicFields & " From KS_GuestBook Where Verific<>0 And IsTop<>0 Order BY LastReplayTime Desc")
	  If Not RS.Eof Then
		Set Application(KS.SiteSN &"TopXML")=KS.RsToXml(RS,"row","")
	  End If
	 RS.Close:Set RS=Nothing
End Sub

Sub LoadMasterUserID(BoardID,ByVal Master)
  Dim Users,RS,str
  If Instr(Master,",")=0 Then 
   Users="'" & Master & "'"
  Else
   Users="'" & Replace(Master,",","','") &"'"
  End If
  Set RS=Conn.Execute("Select UserID,UserName From KS_User Where UserName in (" & Users & ")")
  If Not RS.Eof Then
    Do While Not RS.Eof
	  If Str="" Then
	    str=rs(0) & "|" & rs(1)
	  Else
	    str=str & "@" & rs(0) & "|" & rs(1)
	  End If
	RS.MoveNext
	Loop
  End If
  RS.Close : Set RS=Nothing
  Application(KS.SiteSN &"Master"&BoardID)=str
End Sub

'检查进入版面权限
Function CheckPermissions(KSUser,BSetting,ByRef GuestTitle)
   If KSUser.GroupID="1" Then CheckPermissions="true":Exit Function
   Dim GroupPurview:GroupPurview= True : If Not KS.IsNul(BSetting(1)) and (KS.FoundInArr(Replace(BSetting(1)," ",""),KSUser.GroupID,",")=false Or LoginTF=false) Then GroupPurview=false
   Dim UserPurview:UserPurview=True : If Not KS.IsNul(BSetting(10)) and (KS.FoundInArr(BSetting(10),KSUser.UserName,",")=false or LoginTF=false) Then UserPurview=false
   If KSUser.GetUserInfo("ClubSpecialPower")="1" Then UserPurview=true:GroupPurview=True
   Dim ScorePurview:ScorePurview=KS.ChkClng(BSetting(11))
   Dim MoneyPurview:MoneyPurview=KS.ChkClng(BSetting(12))
   Dim Edays:Edays=0:If LoginTF=True Then Edays=KSUser.GetEdays
   If  BSetting(0)="0" And KS.IsNul(KS.C("UserName")) Then
		CheckPermissions=GetClubErrTips("error1",true)
		GuestTitle="无权进入"
   ElseIf Bsetting(54)="2" And KS.ChkClng(Edays)>0 Then
	    CheckPermissions="true"
   ElseIf Bsetting(54)="1" And KS.ChkClng(Edays)<0 Then
		CheckPermissions=GetClubErrTips("error2",true)
		GuestTitle="无权进入"
   Else
	   If ((GroupPurview=false) or (UserPurview=false)) and boardid<>0 Then
			CheckPermissions=GetClubErrTips("error2",true)
			GuestTitle="无权进入"
	   ElseIf KS.ChkClng(KSUser.GetUserInfo("Score"))<ScorePurView And ScorePurView>0 Then
			CheckPermissions=Replace(Replace(GetClubErrTips("error3",true),"{$Tips}","积分<span>" &ScorePurView&"</span>分"),"{$CurrTips}","积分<span>" & KSUser.GetUserInfo("Score") & "</span>分")
			GuestTitle="无权进入"
	   ElseIf KS.ChkClng(KSUser.GetUserInfo("Money"))<MoneyPurview And MoneyPurview>0 Then
			CheckPermissions=Replace(Replace(GetClubErrTips("error3",true),"{$Tips}","资金￥<span>" &formatnumber(MoneyPurview,2,-1,-1)&"</span>元"),"{$CurrTips}","资金￥<span>" & formatnumber(KSUser.GetUserInfo("money"),2,-1,-1) & "</span>元")
			GuestTitle="无权进入"
	   Else
		  CheckPermissions="true"
	   End If
  End If
End Function

Function GetClubErrTips(ErrId,ShowBack)
    Dim Str:str="<div class=""guest_box""><div class=""errtips"">" &_
	           "<div  class=""tishixx"">" & LFCls.GetConfigFromXML("GuestBook","/guestbook/template",ErrId) & "</div>"&_
			   "<div class=""clear""></div>"
	If ShowBack Then
	     str=str &"<div class=""closebut""><a href=""javascript:history.back()"">返回上一页</a>   <a href=""javascript:window.close()"">关闭本页</a></div>"
	End If
         GetClubErrTips=str &"</div></div>"
End Function

 '发帖，返回帖子ID号
Function InsertPost(BoardID,PostType,UserName,UserID,Subject,Content,Pic,AnnexExt,Purview,ShowIP,ShowSign,ShowScore,CategoryId,Hits,IsTop,IsBest,IsSlide,O_LastPost,verific,ByRef TableName)
			'====================取帖子存放数据表======================
			Dim Nodes,Doc
			set Doc = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			Doc.async = false
			Doc.setProperty "ServerHTTPRequest", true 
			Doc.load(Server.MapPath(KS.Setting(3)&"Config/clubtable.xml"))
			Set Nodes=Doc.DocumentElement.SelectSingleNode("item[@isdefault='1']")
			TableName=nodes.selectsinglenode("tablename").text
			Set Doc=Nothing
			'===========================================================
			
			Dim SqlStr,RSObj,IsPic,TopicID,N_LastPost
			IsPic=0
			If Not KS.IsNul(Pic) Then
			  If lcase(Right(pic,"3"))="gif" Then IsPic=1 Else IsPic=2   
		    End If
         on error resume next
		 Conn.Begintrans
		    SqlStr = "Insert Into KS_GuestBook(BoardID,Subject,PostTable,UserName,UserID,LastReplayUser,LastReplayUserID,face,IsPic,GuestIP,Verific,AddTime,LastReplayTime,Purview,ShowIP,ShowSign,ShowScore,CategoryId,TotalReplay,Hits,IsTop,IsBest,IsSlide,DelTF,AnnexExt,posttype) " & _
			 "values(" & BoardID &",'" & Replace(Subject,"'","''") & "','" & TableName & "','" & UserName & "'," & UserID & ",'" & UserName & "'," & UserID & ",'" & pic & "'," & IsPic &",'" & KS.GetIP & "'," & Verific &"," & SQLNowString &"," & SQLNowString &"," & Purview & "," & ShowIP & "," & ShowSign &"," &ShowScore&"," & CategoryId &",0," & Hits & "," & IsTop & "," & IsBest &"," & IsSlide &",0,'" & Replace(AnnexExt,"'","''") & "'," & posttype &")"
			
			   Conn.Execute(SQLStr)
				'得到帖子ID号
				Set RSObj=Conn.Execute("Select Max(ID) From KS_GuestBook")
				If Not RSObj.Eof Then
				 TopicID=RSObj(0)
				Else
				 TopicID=0
				End If
				RSObj.Close
			
			N_LastPost=TopicID&"$"& now & "$" & Replace(left(subject,200),"$","") & "$" & UserName & "$" &UserID&"$$"
			
			'写入到回复表
			Call InsertReply(TableName,UserName,UserID,TopicID,Content,ShowIP,ShowSign,0,Verific,SQLNowString)

         If err<>0 then
			Conn.RollBackTrans
		 Else
			Conn.CommitTrans
		 End IF

			'关联上传文件
			Call KS.FileAssociation(9994,TopicID,Content,0)
			Call UpdateBoardPostNum(1,BoardID,Verific,O_LastPost,N_LastPost) '更新版面数据
			UpdateTodayPostNum '更新今日发帖数
			InsertPost=TopicID
			
		  '=================同步到第三方微博===============================================
		  if not ks.isnul(pic) then 
		    if left(lcase(pic),4)<>"http" then pic=ks.getdomain & pic
		  end if 
		  if  KS.S("qqweibo")="1" Then
			Call KSUser.add_qq_weibo(Subject&"," & KS.GetClubShowUrl(TopicID),pic)
		  End If
		  If KS.S("sinaweibo")="1" Then
			dim result:result=KSUser.add_sina_weibo(Subject&"," & server.URLEncode(KS.GetClubShowUrl(TopicID)),pic)
		  End If
		'================================================================================
End Function

'发表回复
Sub InsertReply(TableName,UserName,UserID,ByVal TopicID,Content,ShowIP,ShowSign,ParentID,Verific,PostDate)
	Dim SQLStr:SqlStr = "Insert Into " & TableName &"(UserName,UserID,UserIP,TopicID,Content,ShowIP,ShowSign,ReplayTime,ParentId,Verific,DelTF)" &_
			" values('" & UserName &"'," & UserID &",'" & KS.GetIP & "'," & TopicID &",'" & Replace(Content,"'","''") &"'," & ShowIP &"," &ShowSign&"," & PostDate & "," & ParentID & "," & Verific &",0)" 
	Conn.Execute(SQLStr)
	'给用户增加帖子数
	If UserID<>0 Then
		Conn.Execute("Update KS_User Set PostNum=PostNum+1 Where UserID=" & UserID)
	End If
	If IsArray(BSetting) Then
	If KS.ChkClng(Request("categoryid"))<>-1 AND BSetting(68)="1" and BSetting(23)="1" Then
	  Conn.Execute("Update KS_GuestBook Set CategoryID=" & KS.ChkClng(Request("categoryid")) & " Where ID=" & TopicID)
	End If
	End If
End Sub

'更新版面发帖数据
Sub UpdateBoardPostNum(IsTopic,BoardID,Verific,O_LastPost,N_LastPost)
   If BoardID<>0 Then
       If IsTopic=1 Then
		 If verific=0 Then   '帖子需要审核
		  Conn.Execute("Update KS_GuestBoard set postnum=postnum+1,topicnum=topicnum+1 where id=" & BoardID)
         Else
		  Conn.Execute("Update KS_GuestBoard set lastpost='" & N_LastPost & "',postnum=postnum+1,topicnum=topicnum+1 where id=" & BoardID)
		 End If
	   Else
	      Conn.Execute("Update KS_GuestBoard set lastpost='" & N_LastPost & "',postnum=postnum+1 where id=" & BoardID)
	   End If
	   If KS.IsNul(O_LastPost) Then
			  Conn.Execute("Update KS_GuestBoard set todaynum=1 where id=" & BoardID)
			  Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@todaynum").text=1
	   Else
			 Dim LastPostTime:LastPostTime=Split(O_LastPost,"$")(1):If Not IsDate(LastPostTime) Then LastPostTime=now
			 If datediff("d",LastPostTime,Now())=0 Then
				  Conn.Execute("Update KS_GuestBoard set todaynum=todaynum+1 where id=" & BoardID)
				  Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@todaynum").text=KS.ChkClng(Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@todaynum").text)+1
			 Else
				  Conn.Execute("Update KS_GuestBoard set todaynum=1 where id=" & BoardID)
				  Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@todaynum").text=1
			 End If
		End If
		Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@postnum").text=KS.ChkClng(Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@postnum").text)+1
		If IsTopic=1 Then
		  Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@topicnum").text=KS.ChkClng(Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@topicnum").text)+1
		End If
		Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@lastpost").text=N_LastPost
	End  If
End Sub

'更新今日发帖数
Sub UpdateTodayPostNum()
			Dim Doc:set Doc = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			Doc.async = false
			Doc.setProperty "ServerHTTPRequest", true 
			Doc.load(Server.MapPath(KS.Setting(3)&"Config/guestbook.xml"))
			Dim XMLDate:XMLDate=doc.documentElement.attributes.getNamedItem("date").text
			If DateDiff("d",xmldate,now)=0 Then
			   doc.documentElement.attributes.getNamedItem("todaynum").text=doc.documentElement.attributes.getNamedItem("todaynum").text+1
			   If KS.ChkClng(doc.documentElement.attributes.getNamedItem("todaynum").text)>KS.ChkClng(doc.documentElement.attributes.getNamedItem("maxdaynum").text) then
				 doc.documentElement.attributes.getNamedItem("maxdaynum").text=doc.documentElement.attributes.getNamedItem("todaynum").text
			   end if
			Else
			  doc.documentElement.attributes.getNamedItem("date").text=now
			  doc.documentElement.attributes.getNamedItem("yesterdaynum").text=doc.documentElement.attributes.getNamedItem("todaynum").text
			  doc.documentElement.attributes.getNamedItem("todaynum").text=0
			End If
			doc.documentElement.attributes.getNamedItem("topicnum").text=doc.documentElement.attributes.getNamedItem("topicnum").text+1
			doc.documentElement.attributes.getNamedItem("postnum").text=doc.documentElement.attributes.getNamedItem("postnum").text+1
			 doc.save(Server.MapPath(KS.Setting(3)&"Config/guestbook.xml"))
            Set Doc=Nothing
End Sub

'取得投票
Function GetVote(TopicID,Xml)
Dim rs,votestr,VNode,VoteType,VoteID,VoteN,TotalVote,VoteColorArr,CanVote,VoteNums,VoteUserList,TimeLimit,EndTime,IPnums
	VoteColorArr=Array("#E92725","#F27B21","#F2A61F","#5AAF4A","#42C4F5","#0099CC","#3365AE","#2A3591","#592D8E","#DB3191","#cccccc")
	votestr="<div id=""showvote""><table width=""550"" class=""votetable"" cellspacing=""0"" cellpadding=""0""><tr><td colspan=""2"">"
   set rs=conn.execute("select top 1 * from ks_vote where topicid=" & TopicID)
   if rs("VoteType")="Single" Then votestr=votestr & "<strong>单选投票</strong>" Else votestr=votestr &"<strong>多选投票</strong>"
   VoteType=RS("VoteType") :TimeLimit=Rs("TimeLimit") : EndTime=Rs("TimeEnd")
   VoteID=RS("ID") : VoteNums=RS("VoteNums") : VoteUserList=rs("VoteUserList")
   IPnums=RS("IpNums")
   RS.Close : Set RS=Nothing
   If IpNums=1 And KS.FoundInArr(VoteUserList,KS.C("UserName"),",")=true Then CanVote=false Else CanVote=True
   if TimeLimit="1" then votestr=votestr & ",结束时间:"& endtime
   votestr=votestr & ",共有" &VoteNums &"人参与投票, <a href=""javascript:void(0)"" onclick=""showVoteUser('"& KS.Setting(66) &"'," & VoteID& ")"">查看参与用户</a></td></tr>"
						  
						  If Not IsObject(XML) Then Set XML=LFCls.GetXMLFromFile("voteitem/vote_"&VoteID)
						  For Each VNode In Xml.DocumentElement.SelectNodes("voteitem")
							   TotalVote=TotalVote+KS.ChkClng(VNode.childNodes(1).text)
						  Next
						  VoteN=1
						  For Each VNode In Xml.DocumentElement.SelectNodes("voteitem")
						   votestr=votestr & "<tr><td height='30' colspan=""2"">"
						   If CanVote=True Then
							   If VoteType="Single" Then
							   votestr=votestr&"<label><input type='radio' name='VoteOption' value='"& VNode.getAttribute("id") &"' />"
							   Else
							   votestr=votestr&"<label><input type='checkbox' name='VoteOption' value='"& VNode.getAttribute("id") &"' />"
							   End If
						   End If
						   votestr=votestr&VoteN &"、" & VNode.childNodes(0).text & "</label>"
						   votestr=votestr &"</td></tr>"
						   
						   dim perVote,pstr,votebg
							if totalVote=0 Then TotalVote=0.00000001
							perVote=round(VNode.childNodes(1).text/totalVote,4)
							votebg=round(480*perVote)
							perVote=perVote*100
							if perVote<1 and perVote<>0 then
								pstr="&nbsp;0" & perVote & "%"
							else
								pstr="&nbsp;" & perVote & "%"
							end if
						   
						   votestr=votestr & "<tr><td><div class=""vbg""><div style=""width:" & votebg & "px;background:" & VoteColorArr(voten-1) &""">&nbsp;</div></div></td><td align=""left"">" &pstr&"<em style=""color:" & VoteColorArr(voten-1) &""">(" & VNode.childNodes(1).text &")</em></td></tr>"
						   VoteN=VoteN+1
						  Next
						  votestr=votestr &"<tr><td style=""height:40px"" colspan=""2"">"
						  If CanVote Then
						  votestr=votestr&"<input type=""button"" onclick=""doVote('" & KS.Setting(66) & "'," & VoteID & ",'" & VoteType &"')"" id=""votebtn"" value=""投票"" />"
						  Else
						  votestr=votestr&"<input type=""button"" disabled id=""votebtn"" value=""投票"" />"
						  End If
						  votestr=votestr& "</td></tr>"
						VoteStr=VoteStr & "</table></div>"
		GetVote=VoteStr
End Function

 '从xml中加载模型字段
Sub LoadModelField(ChannelID,ByRef FieldXML,ByRef FieldNode)
	set FieldXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	FieldXML.async = false
	FieldXML.setProperty "ServerHTTPRequest", true 
	FieldXML.load(Server.MapPath(KS.Setting(3)&"Config/fielditem/field_" & ChannelID&".xml"))
	if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
	 Set FieldNode=FieldXML.DocumentElement.SelectNodes("fielditem[showonclubform=1]")
	end if
End Sub
		 
'取得点评
Function GetComments(CommentXML,BoardID,replayid,MaxPerPage,IsMaster)
     Dim Str,j,N,P,PageNum,TotalPut
	 P=KS.ChkClng(KS.S("P")) : If P<=0 Then P=1
	 N=0
     If IsObject(CommentXML) Then
		Dim UserFace,CN,CMT,CommentNodes:Set CommentNodes=CommentXML.DocumentElement.SelectNodes("row[@pid=" & replayid & "]")
		TotalPut=CommentNodes.length
		If TotalPut>0 Then
			    if (TotalPut mod MaxPerPage)=0 then
				    PageNum = TotalPut \ MaxPerPage
				else
					PageNum = TotalPut \ MaxPerPage + 1
				end if
				If P>PageNum Then P=PageNum

				Str= "<h3>点评 <span>共 <span class='red'>" &TotalPut& "</span> 条</span></h3>"
				For J=0 To TotalPut
				    Set CN=CommentNodes.Item((p-1)*MaxPerPage+n)
					If CN Is Nothing Then Exit For
			  ' For Each CN In CommentNodes
					CMT=replace(cn.selectsinglenode("@comment").text,chr(10),"<br/>")
					Str= Str & "<div class=""pstl"
					If TotalPut>1 Then Str=Str &" line"
					Str=Str & """>"
					If CN.SelectSingleNode("@userid").text="0" And Instr(CMT,"：")<>0 Then
							Dim K,KK,GD,Star,CommentArr:CommentArr=Split(CMT,"：")
							For K=0 To Ubound(CommentArr)-1
								If K=0 Then
									  GD=CommentArr(k)
								ElseIf Instr(CommentArr(k),"</i> ")<>0 Then
									  GD=split(CommentArr(k),"</i> ")(1)
								End If
								Star=KS.CutFixContent(CMT,GD&"：<i>","</i>",0)
								Str= Str &  GD & "：<span class='red'>" & formatnumber(star,1,-1,-1) & "</span> "
								For KK=0 To 4
								  if cint(kk+1)<=cint(star) Then
									 Str= Str & "<span class='currstar' title='" & star &"'>★</span>"
								  Else
									 Str= Str & "<span class='star'>★</span>"
								  End If
								Next
									Str= Str & "&nbsp;&nbsp;&nbsp;&nbsp;"
							Next
							Str= Str & "</div>"
					Else
							 UserFace=CN.SelectSingleNode("@userface").text
							 If Not KS.IsNUL(UserFace) Then
								If Left(UserFace,1)<>"/" And Left(lcase(UserFace),4)<>"http" Then UserFace=KS.GetDomain & UserFace
							 Else
								UserFace=KS.GetDomain & "images/face/boy.jpg"
							 End If 
							 Str= Str & "<div class=""psta""><a href=""" & KS.GetSpaceUrl(cn.selectsinglenode("@userid").text) & """ target=""_blank""><img onerror='this.src=""" & KS.Setting(3) & "images/face/boy.jpg""' src=""" & UserFace & """ /></a></div>"
							 Str= Str & "<div class=""psti"">"
							 Str= Str & "<a href=""" & KS.GetSpaceUrl(cn.selectsinglenode("@userid").text) & """ target=""_blank"">" & cn.selectsinglenode("@username").text &"</a>&nbsp;" & CMT & "&nbsp;"
										
							 dim ps:ps=cn.selectsinglenode("@prestige").text
							 If KS.ChkClng(ps)<>0 Then 
								   if ps>0 Then
										Str= Str & "威望<span class=""ww"">+" & ps &"</span>&nbsp;"
								   else
										Str= Str & "威望<span class=""ww"">" & ps &"</span>&nbsp;"
								   end if
							 end if
							 Str= Str & "<span class=""xg1"">发表于 " & KS.GetTimeFormat1(cn.selectsinglenode("@adddate").text,true) & "&nbsp;</span>"
							 If IsMaster Then str=str &" <a href='javascript:void(0)' onclick='delCmt(""" & KS.Setting(66) & """," & CN.SelectSingleNode("@id").text & "," & ReplayID&"," & BoardID&"," & p & ")'>删除</a>"
							 Str= Str & "</div>"
							Str= Str & "</div>"
					End If
					N=N+1
					If N>=MaxPerPage Then Exit For
			   Next
			   
		  Str=Str &"<div class=""cmtpage"">"
		  If PageNum>1 Then
			  If P=1 Then 
			   Str=Str &"<a href='javascript:void(0)' onclick='ShowCmtPage(""" & KS.Setting(66) & """,2," & replayid&"," & BoardID&")'>下一页 >> </a>"   
			  Else
				  If P>1 Then
				  Str=Str &"<a href=javascript:void(0) onclick=ShowCmtPage('" & KS.Setting(66) & "',"&p-1& "," & replayid&"," & BoardID&")><< 上一页</a>"   
				  End If
				For K=1 To PageNum
				  If K=P Then
				  Str=Str &"<a href='javascript:void(0)' class='curr'>" & k & "</a>"   
				  Else
				  Str=Str &"<a href=javascript:void(0) onclick=ShowCmtPage('" & KS.Setting(66) & "',"&k& "," & replayid&"," & BoardID&")>" & k & "</a>"   
				  End If
				Next
				If P>1 And P<>PageNum Then
				  Str=Str &"<a href=javascript:void(0) onclick=ShowCmtPage('" & KS.Setting(66) & "',"&p+1& "," & replayid&"," & BoardID&")>下一页 >></a>"   
				End If
			  End If
		  End If
		 Str=Str &"</div>"
			   
		  End If
	 End If
	 GetComments=str
End Function
%>
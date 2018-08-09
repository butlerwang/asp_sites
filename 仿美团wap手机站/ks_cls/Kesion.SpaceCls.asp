<!--#include file="Kesion.SpaceCalCls.asp"-->
<%

Class BlogCls
      Public KS,UserName,UserID,Domain,Node,Title,PreviewTemplateID
	  Private Sub Class_Initialize()
	   Set KS=New PublicCls
      End Sub
	 Private Sub Class_Terminate()
	  Set KS=Nothing
	 End Sub
	 %>
	 <!--#include file="ubbfunction.asp"-->
	 <%
	 '读出日志模板 FieldName 模板字段
	 Function GetTemplatePath(TemplateID,FieldName)
	  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
	  RS.Open "Select top 1 " & FieldName & " From KS_BlogTemplate Where ID=" & KS.ChkCLng(TemplateID),conn,1,1
	  If RS.Eof And RS.Bof Then
	    RS.Close
		if conn.execute("Select top 1 username From KS_enterprise Where UserName='" & UserName & "'").eof Then
		RS.Open "Select top 1 " & FieldName & " From KS_BlogTemplate Where flag=4 and IsDefault='true'",conn,1,1
		else
		RS.Open "Select top 1 " & FieldName & " From KS_BlogTemplate Where flag=2 and IsDefault='true'",conn,1,1
		end if
	  End If
	    Dim KSR:Set KSR = New Refresh 
		GetTemplatePath=KSR.LoadTemplate(RS(0))
		Set KSR=Nothing
        RS.Close:Set RS=Nothing
	 End Function

	 
	 '取得用户参数
	 Function GetUserBlogParam(UserName,FieldName)
	     Dim Num:Num=0
	     Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select Top 1 " & FieldName & " From KS_Blog Where UserName='" & UserName & "'",conn,1,1
		 if Not RS.Eof Then
		  Num=KS.ChkClng(RS(0))
		 End if
		 RS.Close:Set RS=Nothing
		 If Num=0 Then Num=10
		 GetUserBlogParam=Num
	 End Function
	 
	 '空间头部
	 Sub LoadSpaceHead() 
	     Exit Sub
	     With KS
		  .echo "<html>"&vbcrlf &"<title>" & Node.SelectSingleNode("@blogname").text & "-" & Title & "</title>" &vbcrlf
		  .echo "<meta http-equiv=""Content-Language"" content=""zh-CN"" />" &vbcrlf
          .echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" />" & vbcrlf
          .echo "<meta name=""generator"" content=""KesionCMS"" />" & vbcrlf
		  .echo "<meta name=""author"" content=""" & UserName & ","" />" & vbcrlf
		  .echo "<meta name=""keyword"" content=""" & Node.SelectSingleNode("@blogname").text & """ />"&VBCRLF
		  .echo "<meta name=""description"" content=""" & Node.SelectSingleNode("@descript").text & """ />"  & vbcrlf
		  .echo "<link href=""" & KS.GetDomain & "space/css/css.css"" type=""text/css"" rel=""stylesheet"">" & vbcrlf
		  .echo "<script src=""" & KS.GetDomain & "ks_inc/kesion.box.js"" language=""javascript""></script>"  & vbcrlf
		  .echo "<script src=""" & KS.GetDomain & "ks_inc/jquery.js"" language=""javascript""></script>"  & vbcrlf
		  .echo "<script src=""" & KS.GetDomain & "space/js/ks.space.js"" language=""javascript""></script>"  & vbcrlf
		  .echo "<script src=""" & KS.GetDomain & "space/js/ks.space.page.js"" language=""javascript""></script>"  & vbcrlf
		 End With
	 End Sub
	 
	 '日志链接
	 Function GetLogUrl(RS)
	  GetLogUrl=GetCurrLogUrl(RS("UserID"),RS("ID"))
	 End Function
	 Function GetCurrLogUrl(UserID,ID)
	  If KS.SSetting(21)="1" Then
	  GetCurrLogUrl=KS.GetDomain &"space/list-" & userid & "-" & id&KS.SSetting(22)
	  Else
	  GetCurrLogUrl="../space/?" & userid & "/log/" & id
	  End If
	 End Function
	 
	 '替换用户博客所有标签
	 Function ReplaceBlogLabel(Template)
	  UserName=Node.SelectSingleNode("@username").text
	  Template=Replace(Template,"{$ShowAnnounce}",KS.CheckXSS(KS.LoseHtml(Node.SelectSingleNode("@announce").text)))
	  Template=Replace(Template,"{$ShowBlogName}",KS.CheckXSS(KS.LoseHtml(Node.SelectSingleNode("@blogname").text)))
	  Template=Replace(Template,"{$ShowBlogDescript}",KS.CheckXSS(KS.LoseHtml(Node.SelectSingleNode("@descript").text)))
	  Template=Replace(Template,"{$GetSiteTitle}",Title)
	  Template=Replace(Template,"{$ShowLogo}",ReplaceLogo(Node.SelectSingleNode("@logo").text))
	 
	 
	  Dim b1,b2,b3,Banner:Banner=Node.SelectSingleNode("@banner").text
	  If Banner="" Or IsNull(Banner) Then Banner="|"
	  Banner=Split(Banner,"|") 
	  b1=Banner(0) : If B1="" Then b1="../images/ad1.jpg"
	  If Ubound(Banner)>=1 Then b2=Banner(1) 
	  If B2="" Then B2="../images/ad1.jpg"
	  If Ubound(Banner)>=2 Then B3=Banner(2) 
	  If B3="" Then B3="../images/ad1.jpg"
	  Template=Replace(Template,"{$ShowBannerSrc}",B1)
	  Template=Replace(Template,"{$ShowBannerSrc1}",B1)
	  Template=Replace(Template,"{$ShowBannerSrc2}",B2)
	  Template=Replace(Template,"{$ShowBannerSrc3}",B3)
	  
	  '自定义图片
	  If Instr(Template,"{$ShowPicture")<>0 Then
	     Dim PXml,PN,kk,rstr
	     Dim RSP:Set RSP=Server.CreateObject("ADODB.RECORDSET")
		 RSP.Open "Select Top 200 * From KS_BlogSkin Where TemplateID=" & KS.ChkClng(Node.SelectSingleNode("@templateid").text) &" and (IsDefault=1 Or UserName='" & Node.SelectSingleNode("@username").text & "') Order By OrderID,ID",conn,1,1
		 If Not RSP.Eof Then
		   Set PXml=KS.RsToXml(RSP,"row","")
		 End if
		 RSP.Close :Set RSP=Nothing
		 If IsOBject(Pxml) Then
		    Dim UNode,DNode,ND:Set DNode=PXml.DocumentElement.SelectNodes("row[@isdefault=1]")
			kk=0
			For Each ND In DNode
			   KK=KK+1
			   Set UNode=Pxml.DocumentElement.SelectSingleNode("row[@isdefault=0][@orderid=" & kk & "]")
			   If Not Unode  Is Nothing Then 			'如果用户有自定义，则更新
			     ND.SelectSingleNode("@photourl").text=UNode.SelectSingleNode("@photourl").text
			     ND.SelectSingleNode("@linkurl").text=UNode.SelectSingleNode("@linkurl").text
			   End If
			   
			   '替换标签
			   If ND.SelectSingleNode("@isbg").text="1" Then
			      rstr=ND.SelectSingleNode("@photourl").text
			   Else
			      dim w,h
				  w=ND.SelectSingleNode("@width").text  : If KS.ChkClng(w)<>0 Then w=" width=""" & w &""""
				  h=ND.SelectSingleNode("@height").text : If KS.ChkClng(h)<>0 Then h=" height=""" & h& """"
			      if Not KS.IsNul(ND.SelectSingleNode("@linkurl").text) Then
				   rstr="<a title='" & ND.SelectSingleNode("@descript").text &"' href='" & ND.SelectSingleNode("@linkurl").text &"' target='_blank'><img src='" & ND.SelectSingleNode("@photourl").text & "' border='0'" & w & h& "/></a>"
				  Else
				   rstr="<img title='" & ND.SelectSingleNode("@descript").text &"' src='" & ND.SelectSingleNode("@photourl").text & "' border='0'" & w & h& "/>"
				  End If
			   End If
			   Template=Replace(Template,"{$ShowPicture" & Kk &"}", rstr)
			Next
			
			
			
		 End If
		 
		 
	  End If
	  
	  
	  
	  Template=Replace(Template,"{$ShowNavigation}",ReplaceMenu)
	  Template=Replace(Template,"{$ShowUserLogin}","<iframe width=""170"" height=""122"" id=""login"" name=""login"" src=""../user/userlogin.asp"" frameBorder=""0"" scrolling=""no"" allowtransparency=""true""></iframe>")
	    If Instr(Template,"{$ShowNewLog}")<>0 Then
		 Template=Replace(Template,"{$ShowNewLog}",GetNewLog)
		 End If
		 If Instr(Template,"{$ShowNewAlbum}")<>0 Then
		 Template=Replace(Template,"{$ShowNewAlbum}",GetNewAlbum)
		 End If
		 If Instr(Template,"{$ShowNewInfo}")<>0 Then
		 Template=Replace(Template,"{$ShowNewInfo}",GetNewXX)
		 End If
		 If Instr(Template,"{$ShowClubTopic}")<>0 Then
		 Template=Replace(Template,"{$ShowClubTopic}",GetClubTopic)
		 End If
		 If Instr(Template,"{$ShowNewFresh}")<>0 Then
		 Template=Replace(Template,"{$ShowNewFresh}",GetNewFresh)
		 End If
		 '=================企业空间替换==========================
		 If Instr(Template,"{$ShowNews}")<>0 Then
		 Template=Replace(Template,"{$ShowNews}",GetEnterPriseNews)
		 End If
		 If Instr(Template,"{$ShowSupply}")<>0 Then
		 Template=Replace(Template,"{$ShowSupply}",GetSupply)
		 End If
		 If Instr(Template,"{$ShowProduct}")<>0 Then
		 Template=Replace(Template,"{$ShowProduct}",GetProduct)
		 End If
		 If Instr(Template,"{$ShowProductList}")<>0 Then
		 Template=Replace(Template,"{$ShowProductList}",GetProductList)
		 End If
		 If Instr(Template,"{$ShowIntro}")<>0 Then
		 Template=Replace(Template,"{$ShowIntro}",GetEnterpriseintro)
		 End If
		 If Instr(Template,"{$ShowShortIntro}")<>0 Then
		 Template=Replace(Template,"{$ShowShortIntro}",GetEnterpriseShortintro)
		 End If
		 
		 Template=Replace(Template,"{$ShowContact}",GetEnterpriseContact)
		 Template=Replace(Template,"{$ShowNews}",GetEnterpriseNews)
		 Template=ReplaceUserInfoContent(Template,0)

		 '========================================================
	 
	 
	   If Instr(Template,"{$ShowUserInfo}")<>0 Then
	   Template=Replace(Template,"{$ShowUserInfo}",GetUserInfo)
	   End If
	   If Instr(Template,"{$ShowCalendar}")<>0 Then
	   Template=Replace(Template,"{$ShowCalendar}",Getcalendar)
	   End If
	   If Instr(Template,"{$ShowUserClass}")<>0 Then
	   Template=Replace(Template,"{$ShowUserClass}",GetUserClass)
	   End If
	   If Instr(Template,"{$ShowComment}")<>0 Then
	   Template=Replace(Template,"{$ShowComment}",GetComment)
	   End If
	   If Instr(Template,"{$ShowMusicBox}")<>0 Then
	   Template=Replace(Template,"{$ShowMusicBox}",GetMusicBox)
	   End If
	   If Instr(Template,"{$GetMediaPlayer}")<>0 Then
	   Template=Replace(Template,"{$GetMediaPlayer}",GetMediaPlayer)
	   End If
	   If Instr(Template,"{$ShowMessage}")<>0 Then
	   Template=Replace(Template,"{$ShowMessage}",Replace(GetMessage,"{","｛#"))
	   End If
	   If Instr(Template,"{$ShowBlogInfo}")<>0 Then
	   Template=Replace(Template,"{$ShowBlogInfo}",GetBlogInfo)
	   End If
	   If Instr(Template,"{$ShowBlogTotal}")<>0 Then
	   Template=Replace(Template,"{$ShowBlogTotal}",GetBlogTotal)
	   End If
	   If Instr(Template,"{$ShowSearch}")<>0 Then
	   Template=Replace(Template,"{$ShowSearch}",GetSearch)
	   End If
	   If Instr(Template,"{$ShowVisitor}")<>0 Then
	   Template=Replace(Template,"{$ShowVisitor}",GetVisitor)
	   End If
	   Template=Replace(Template,"{$ShowXML}",GetXML)
	   Template=Replace(Template,"{$ShowUserName}",UserName)
	   Template=Replace(Template,"{$ShowUserID}",UserID)
	   Template=Replace(Template,"{$ShowSlidePhoto}",GetSlidePhoto(2))
	   
	   
	   Dim KSR:Set KSR = New Refresh 
	   Template=KSR.KSLabelReplaceAll(Template)
	   Set KSR=Nothing	
		
	   ReplaceBlogLabel=Template
	 End Function	 
	 
	 Function ReplaceLogo(Logo)
	  If KS.IsNul(Logo) Then Logo="../images/logo.jpg"
	  ReplaceLogo="<Img src=""" & Logo & """ align=""absmiddle"" width=""130"">"
	 End Function
	 
	 Function ReplaceMenu() 
	   Dim HomeUrl,BlogUrl,MessageUrl,ProductUrl,IntroUrl,NewsUrl,JobUrl,RyzsUrl,ClubUrl
	   Dim AlbumUrl,GroupUrl,FriendUrl,XXUrl,InfoUrl
	   If KS.SSetting(21)="1" Then
	    If Not KS.IsNul(Domain) Then
			If Instr(Domain,".")<>0 Then
			 HomeUrl   = "http://" & domain
			Else
			 HomeUrl   = "http://" & domain &"." & KS.SSetting(16)
			End If
		Else
		  HomeUrl ="" & userid
		End If
		BlogUrl   = "blog-" & userid
		ClubUrl   = "club-" & userid
		MessageUrl= "message-"&userid
		ProductUrl= "product-"&userid
		IntroUrl  = "intro-" &userid
		NewsUrl   = "news-" & userid
		JobUrl    = "job-" & userid
		RyzsUrl   = "ryzs-" &userid
		AlbumUrl  = "album-" & userid
		GroupUrl  = "group-" & userid
		FriendUrl = "friend-" & userid
		XXUrl     = "xx-" & userid
		InfoUrl   = "info-" & userid
	   Else
	    HomeUrl   = "../space/?" & userid
		BlogUrl   = "../space/?" & userid & "/blog"
		ClubUrl   = "../space/?" & userid & "/club"
		MessageUrl= "../space/?" & userid & "/message"
		ProductUrl= "../space/?" & userid & "/product"
		IntroUrl  = "../space/?" & userid & "/intro"
		NewsUrl   = "../space/?" & userid &"/news"
		JobUrl    = "../space/?" & userid & "/job"
		RyzsUrl   = "../space/?" & userid & "/ryzs"
		AlbumUrl  = "../space/?" & userid & "/album"
		GroupUrl  = "../space/?" & userid & "/group"
		FriendUrl = "../space/?" & userid & "/friend"
		XXUrl     = "../space/?" & userid & "/xx"
		InfoUrl   = "../space/?" & userid & "/info"
	   End If
	   If PreviewTemplateID<>0 Then    '判断是不是预览模板
		BlogUrl   = "../space/?" & userid & "/blog" & "&" & PreviewTemplateID
		BlogUrl   = "../space/?" & userid & "/club" & "&" & PreviewTemplateID
		MessageUrl= "../space/?" & userid & "/message" & "&" & PreviewTemplateID
		ProductUrl= "../space/?" & userid & "/product" & "&" & PreviewTemplateID
		IntroUrl  = "../space/?" & userid & "/intro" & "&" & PreviewTemplateID
		NewsUrl   = "../space/?" & userid &"/news" & "&" & PreviewTemplateID
		JobUrl    = "../space/?" & userid & "/job" & "&" & PreviewTemplateID
		RyzsUrl   = "../space/?" & userid & "/ryzs" & "&" & PreviewTemplateID
		AlbumUrl  = "../space/?" & userid & "/album" & "&" & PreviewTemplateID
		GroupUrl  = "../space/?" & userid & "/group" & "&" & PreviewTemplateID
		FriendUrl = "../space/?" & userid & "/friend" & "&" & PreviewTemplateID
		XXUrl     = "../space/?" & userid & "/xx" & "&" & PreviewTemplateID
		InfoUrl   = "../space/?" & userid & "/info" & "&" & PreviewTemplateID
	   End If
	   
	  If KS.Setting(56)="1" Then
		Dim ClubStr:ClubStr="<li><a href=""" & ClubUrl & """>论坛</a></li>"
	  End If

	   
	  if conn.execute("Select top 1 username From KS_enterprise Where UserName='" & UserName & "'").eof Then
	    ReplaceMenu="<div id=""Menu"">"_
	                 & "<ul>"_
					 &" <li><a href=""" & HomeUrl & """>首页</a></li>"_
					 &" <li><a href=""" & BlogUrl & """>博文</a></li>"_
					 &ClubStr&" <li><a href=""" & AlbumUrl & """>相册</a></li>"_
					 &" <li><a href=""" & GroupUrl & """>圈子</a></li>" _
					 &" <li><a href=""" & FriendUrl & """>好友</a></li>"_
					 &" <li><a href=""" & XXUrl & """>文集</a></li>"_
					 &" <li><a href=""" & InfoUrl & """>小档案</a></li>"_
					 &" <li><a href=""" & MessageUrl & """>留言板</a>"_
					 &"</ul>"_
					 &"</div>"
	  Else
	   	  ReplaceMenu="<div id=""Menu"">"_
	                 & "<ul>"_
					 &" <li><a href=""" & HomeUrl & """>首页</a></li>"_
					 &" <li><a href=""" & introUrl & """ title=""公司简介"">简介</a></li>"_
					 &" <li><a href=""" & NewsUrl & """ title=""公司动态"">动态</a></li>"_
					 &" <li><a href=""" & ProductUrl & """  title=""产品展示"">产品</a></li>"_
					 &" <li><a href=""" & JobUrl & """  title=""公司招聘"">招聘</a></li>"_
					 &" <li><a href=""" & AlbumUrl & """  title=""公司相册"">相册</a></li>"_
					 &" <li><a href=""" & ryzsurl & """  title=""公司证书"">证书</a></li>"_
					 &" <li><a href=""" & GroupUrl & """ title=""公司圈子"">圈子</a></li>" _
					 &" <li><a href=""" & BlogUrl & """  title=""博文日志"">博文</a></li>"_
					 &clubstr &" <li><a href=""" & XXUrl & """>文集</a></li>"_
					 &" <li><a href=""" & InfoUrl & """>联系我们</a></li>"_
					 &" <li><a href=""" & MessageUrl & """>留言板</a>"_
					 &"</ul>"_
					 &"</div>"
	  End If
	 End Function
	 
	 
	 
	 '取得联信息
	 Function UserInfo(ByVal Template)
	    Dim Str
	    Str=Location("首页 >> 联系档案")
		Str=ReplaceUserInfoContent(Template,1)
		UserInfo=Str
	 End Function

		 Function ReplaceUserInfoContent(ByVal Content,Flag)
	 	Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select top 1 * From KS_User Where UserName='" & UserName & "'",conn,1,1
		If RS.Eof And RS.Bof Then
		 RS.Close:Set RS=Nothing
		 ReplaceUserInfoContent=Content
		 Exit Function
		End If
       
	   If Flag=1 Then
			If RS("UserType")=1 Then 
			 Content=LFCls.GetConfigFromXML("space","/labeltemplate/label","companyinfo")
			 ReplaceUserInfoContent=ReplaceEnterpriseInfo(Content,RS("UserName"))
			 Exit Function
			Else
			 Content=LFCls.GetConfigFromXML("space","/labeltemplate/label","userinfo")
			End If
	   ElseIf RS("UserType")=1 Then
	     ReplaceUserInfoContent=ReplaceEnterpriseInfo(Content,RS("UserName"))
		 Exit Function
	   End If
	  
	  
        Dim Privacy:Privacy=RS("Privacy")
        Content=Replace(Replace(Content,"{$GetUserName}",RS("UserName")),"{$GetUserID}",RS("UserID"))
	  Dim UserFaceSrc:UserFaceSrc=RS("UserFace")
	  if left(userfacesrc,1)<>"/" and lcase(left(userfacesrc,4))<>"http" then userfacesrc="../" & userfacesrc
		Content=Replace(Content,"{$GetUserFace}","<img src=" & UserFaceSrc & " border=""1"" />")
		Content =ReplaceUserDefine(101,Content,RS)
          

		'联系方式
    	If Privacy=2 Then
		 Content=Replace(Content,"{$GetEmail}","保密")
		Else
		 Dim Email:Email=RS("Email")
		 If KS.IsNul(Email) Then Email="暂无"
		 Content=Replace(Content,"{$GetEmail}",KS.CheckXSS(Email))
		End If
    	If Privacy=2 Then
		 Content=Replace(Content,"{$GetQQ}","保密")
		Else
		 Dim QQ:QQ=RS("QQ")
		 If KS.IsNul(QQ) Then QQ="暂无"
		 Content=Replace(Content,"{$GetQQ}",KS.CheckXSS(QQ))
		End If
    	If Privacy=2 Then
		 Content=Replace(Content,"{$GetUC}","保密")
		Else
		 Dim UC:UC=RS("UC")
		 If KS.IsNul(UC) Then UC="暂无"
		 Content=Replace(Content,"{$GetUC}",KS.CheckXSS(UC))
		End If
		Content=Replace(Content,"{$GetRegDate}",RS("RegDate"))
		If Privacy=2 Then
		 Content=Replace(Content,"{$GetMSN}","保密")
		Else
		 Dim MSN:MSN=RS("MSN")
		 If KS.IsNul(MSN) Then MSN="暂无"
		 Content=Replace(Content,"{$GetMSN}",KS.CheckXSS(MSN))
		End If
    	If Privacy=2 Then
		 Content=Replace(Content,"{$GetHomePage}","保密")
		Else
		 Dim HomePage:HomePage=RS("MSN")
		 If Not IsNull(HomePage) Then
		 Content=Replace(Content,"{$GetHomePage}","<a href=""" & KS.CheckXSS(RS("HomePage")) & """ target=""_blank"">" & KS.CheckXSS(RS("HomePage")) & "</a>")
		 Else
		   Content=Replace(Content,"{$GetHomePage}","")
		 End iF
		End If


		If Privacy=2 or Privacy=1 Then
		 Content=Replace(Content,"{$GetRealName}","保密")
		Else
		 Dim RealName:RealName=RS("RealName")
		 If IsNull(RealName) Or RealName="" Then RealName="暂无"
		 Content=Replace(Content,"{$GetRealName}",KS.CheckXSS(RealName))
		End If
		If Privacy=2 or Privacy=1 Then
		 Content=Replace(Content,"{$GetSex}","保密")
		Else
		 Dim Sex:Sex=RS("Sex")
		 If IsNull(Sex) Or Sex="" Then Sex="暂无"
		 Content=Replace(Content,"{$GetSex}",KS.CheckXSS(Sex))
		End If
		If Privacy=2 or Privacy=1 Then
		 Content=Replace(Content,"{$GetBirthday}","保密")
		Else
		  Dim BirthDay:BirthDay=RS("BirthDay")
		 If IsNull(BirthDay) Or BirthDay="" Then BirthDay="暂无"
		 Content=Replace(Content,"{$GetBirthday}",KS.CheckXSS(BirthDay))
		End If
		If Privacy=2 or Privacy=1 Then
		 Content=Replace(Content,"{$GetIDCard}","保密")
		Else
		 Dim IDCard:IDCard=RS("IDCard")
		 If IsNull(IDCard) Or IDCard="" Then IDCard="暂无"
		 Content=Replace(Content,"{$GetIDCard}",KS.CheckXSS(IDCard))
		End If
		If Privacy=2 or Privacy=1 Then
		 Content=Replace(Content,"{$GetOfficeTel}","保密")
		Else
		 Dim OfficeTel:OfficeTel=RS("OfficeTel")
		 If IsNull(OfficeTel) Or OfficeTel="" Then OfficeTel="暂无"
		 Content=Replace(Content,"{$GetOfficeTel}",KS.CheckXSS(OfficeTel))
		End If
		If Privacy=2 or Privacy=1 Then
		 Content=Replace(Content,"{$GetHomeTel}","保密")
		Else
		 Dim HomeTel:HomeTel=RS("HomeTel")
		 If IsNull(HomeTel) Or HomeTel="" Then HomeTel="暂无"
		 Content=Replace(Content,"{$GetHomeTel}",KS.CheckXSS(HomeTel))
		End If

		If Privacy=2 or Privacy=1 Then
		 Content=Replace(Content,"{$GetMobile}","保密")
		Else
		 Dim Mobile:Mobile=RS("Mobile")
		 If IsNull(Mobile) Or Mobile="" Then Mobile="暂无"
		 Content=Replace(Content,"{$GetMobile}",KS.CheckXSS(Mobile))
		End If
		If Privacy=2 or Privacy=1 Then
		 Content=Replace(Content,"{$GetFax}","保密")
		Else
		 Dim Fax:Fax=RS("Fax")
		 If IsNull(Fax) Or Fax="" Then Fax="暂无"
		 Content=Replace(Content,"{$GetFax}",KS.CheckXSS(Fax))
		End If
		If Privacy=2 or Privacy=1 Then
		 Content=Replace(Content,"{$GetUserArea}","保密")
		Else
		 Dim Province:Province=RS("Province")
		 If IsNull(Province) Or Province="" Then Province=""
		 Dim City:City=RS("City")
		 If IsNull(City) Or Fax="" Then City="未知"
		 Content=Replace(Content,"{$GetUserArea}",KS.CheckXSS(Province & City))
		End If

		If Privacy=2 or Privacy=1 Then
		 Content=Replace(Content,"{$GetAddress}","保密")
		Else
		 Dim AddRess:AddRess=RS("AddRess")
		 If IsNull(AddRess) Or AddRess="" Then AddRess="暂无"
		 Content=Replace(Content,"{$GetAddress}",KS.CheckXSS(AddRess))
		End If
		If Privacy=2 or Privacy=1 Then
		 Content=Replace(Content,"{$GetZip}","保密")
		Else
		 Dim Zip:Zip=RS("Zip")
		 If IsNull(Zip) Or Zip="" Then Zip="暂无"
		 Content=Replace(Content,"{$GetZip}",KS.CheckXSS(ZIP))
		End If
		If Privacy=2 or Privacy=1 Then
		 Content=Replace(Content,"{$GetSign}","保密")
		Else
		 Dim Sign:Sign=RS("Sign")
		 If IsNull(Sign) Or Sign="" Then Sign="暂无"
		 Content=Replace(Content,"{$GetSign}",KS.CheckXSS(Ubbcode(KS.FilterIllegalChar(Sign),0)))
		End If
		If Instr(Content,"{$ShowNewFresh}")<>0 Then
		 Content=Replace(Content,"{$ShowNewFresh}",GetNewFresh)
		End If
        ReplaceUserInfoContent=Content
		rs.Close :Set RS=Nothing
  End Function
	 
	 
	 
	  Function ReplaceEnterpriseInfo(ByVal Content,username)
	   On Error Resume Next
	   Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
	   RS.Open "Select top 1 a.mapmarker,a.classid,a.smallclassid,a.CompanyName as myCompanyName,BusinessLicense,profession,Companyscale,Contactman,a.ZipCode as myZipCode,a.telphone as mytelphone,a.province as myprovince,a.city as mycity,a.address as myaddress,a.fax as myfax,a.Mobile as mymobile,a.qq as myqq,a.email as myemail,weburl,bankaccount,accountnumber,b.* From KS_EnterPrise a inner join ks_user b on a.username=b.username Where a.UserName='" & UserName & "'",conn,1,1
	   IF RS.Eof Then
	    RS.Close:Set RS=Nothing
		ReplaceEnterpriseInfo=""
	   End If
	   Content=Replace(Content,"{$GetCompanyName}",KS.CheckXSS(RS("myCompanyName")))
	   if isnull(RS("BusinessLicense")) then
	   Content=Replace(Content,"{$GetBusinessLicense}","---")
	   else
	   Content=Replace(Content,"{$GetBusinessLicense}",KS.CheckXSS(RS("BusinessLicense")))
	   end if
	   if instr(content,"{$GetProfession}")<>0 then
	   Content=Replace(Content,"{$GetProfession}",conn.execute("select top 1 classname from ks_enterpriseclass where id=" &RS("classid"))(0)&"-" &conn.execute("select top 1 classname from ks_enterpriseclass where id=" &RS("smallclassid"))(0))
	   end if
	   if isnull(RS("Companyscale")) then
	   Content=Replace(Content,"{$GetCompanyScale}","---")
	   else
	   Content=Replace(Content,"{$GetCompanyScale}",KS.CheckXSS(RS("Companyscale")))
	   end if
	   if isnull(rs("myprovince")) then
	   Content=Replace(Content,"{$GetProvince}","---")
	   else
	   Content=Replace(Content,"{$GetProvince}",KS.CheckXSS(RS("myprovince")))
	   end if
	   if isnull(rs("mycity")) then
	   Content=Replace(Content,"{$GetCity}","---")
	   else
	   Content=Replace(Content,"{$GetCity}",KS.CheckXSS(RS("mycity")))
	   end if
	   if isnull(RS("Contactman")) then
	   Content=Replace(Content,"{$GetContactMan}","---")
	   else
	   Content=Replace(Content,"{$GetContactMan}",KS.CheckXSS(RS("Contactman")))
	   end if
	   if isnull(RS("myaddress")) then
	   Content=Replace(Content,"{$GetAddress}","---")
	   else
	   Content=Replace(Content,"{$GetAddress}",KS.CheckXSS(RS("myaddress")))
	   end if
	   if isnull(RS("myZipCode")) Then
	   Content=Replace(Content,"{$GetZipCode}","---")
	   Else
	   Content=Replace(Content,"{$GetZipCode}",KS.CheckXSS(RS("myzipcode")))
	   End If
       If Isnull(RS("mytelphone")) Then
	   Content=Replace(Content,"{$GetTelphone}","---")
	   Else
	   Content=Replace(Content,"{$GetTelphone}",KS.CheckXSS(RS("mytelphone")))
	   End If
	   
	   If IsNull(rs("myfax")) then
	   Content=Replace(Content,"{$GetFax}","---")
	   else
	   Content=Replace(Content,"{$GetFax}",KS.CheckXSS(RS("myfax")))
	   end if
	   if isnull(rs("weburl")) then
	   Content=Replace(content,"{$GetWebUrl}","---")
	   else
	   Content=Replace(Content,"{$GetWebUrl}",KS.CheckXSS(RS("weburl")))
	   end if
	   if isnull(rs("bankaccount")) then
	   Content=Replace(Content,"{$GetBankAccount}","---")
	   else
	   Content=Replace(Content,"{$GetBankAccount}",KS.CheckXSS(RS("bankaccount")))
	   end if
	   if isnull(RS("accountnumber")) then
	   Content=Replace(Content,"{$GetAccountNumber}","---")
	   else
	   Content=Replace(Content,"{$GetAccountNumber}",KS.CheckXSS(RS("accountnumber")))
	   end if
	   if isnull(RS("myMobile")) then
	   Content=Replace(Content,"{$GetMobile}","---")
	   else
	   Content=Replace(Content,"{$GetMobile}",KS.CheckXSS(RS("mymobile")))
	   end if
	   if isnull(RS("myQQ")) then
	   Content=Replace(Content,"{$GetQQ}","---")
	   else
	   Content=Replace(Content,"{$GetQQ}",KS.CheckXSS(RS("myQQ")))
	   end if
	   if isnull(RS("myEmail")) then
	   Content=Replace(Content,"{$GetEmail}","---")
	   else
	   Content=Replace(Content,"{$GetEmail}",KS.CheckXSS(RS("myEmail")))
	   end if
        
	   Content=Replace(Content,"{$MapKey}",KS.Setting(175))
	   Dim MapMarker,MarkerArr,ii,MarkerStr
	   MapMarker=rs("mapmarker")
	   if Not KS.IsNul(MapMarker) Then
		 MarkerArr=Split(MapMarker,"|")
		 Content=Replace(Content,"{$MapCenterPoint}",MarkerArr(0))
		 For ii=0 to Ubound(MarkerArr)
			MarkerStr=MarkerStr & "point = new BMap.Point(" & MarkerArr(ii) & "); " & vbcrlf
			MarkerStr=MarkerStr & "addMarker(point, " & ii & ");" &vbcrlf
		 Next
	     Content=Replace(Content,"{$ShowMarkerList}",MarkerStr)
	   Else
	     Content=Replace(Content,"{$MapCenterPoint}",KS.Setting(176))
	  End If
         
	   Content =ReplaceUserDefine(101,Content,RS)
	   ReplaceEnterpriseInfo=Content
	End Function
	
	 '替换自定义字段
	Function ReplaceUserDefine(ChannelID,F_C,ByVal RS)
		   If Not IsObject(Application(KS.SiteSN&"_userfiledlist"&channelid)) Then
		     Set  Application(KS.SiteSN&"_userfiledlist"&channelid)=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			 Application(KS.SiteSN&"_userfiledlist"&channelid).appendChild(Application(KS.SiteSN&"_userfiledlist"&channelid).createElement("xml"))
				Dim D_F_Arr,K,Node,FieldName
				Dim KS_RS_Obj:Set KS_RS_Obj=Conn.Execute("Select FieldName From KS_Field Where ChannelID=" & ChannelID &" And ShowOnForm=1 Order By OrderID Asc")
				If Not KS_RS_Obj.Eof Then D_F_Arr=KS_RS_Obj.GetRows(-1)
			    KS_RS_Obj.Close:Set KS_RS_Obj=Nothing
				If IsArray(D_F_Arr) Then
					  For K=0 To Ubound(D_F_Arr,2)
						Set Node=Application(KS.SiteSN&"_userfiledlist"&channelid).documentElement.appendChild(Application(KS.SiteSN&"_userfiledlist"&channelid).createNode(1,"userfiledlist"&channelid,""))
						Node.attributes.setNamedItem(Application(KS.SiteSN&"_userfiledlist"&channelid).createNode(2,"fieldname","")).text=D_F_Arr(0,K)
					 Next
				 End If
		 End If
on error resume next
		 For Each Node in Application(KS.SiteSN&"_userfiledlist"&channelid).documentElement.SelectNodes("userfiledlist"&channelid)
			 FieldName=Node.selectSingleNode("@fieldname").text
			 If Left(Lcase(FieldName),3)="ks_" Then
			    
				If Not IsNull(RS(FieldName)) Then
				  F_C=Replace(F_C,"{$" & FieldName & "}",KS.CheckXSS(RS(FieldName)))
				Else
				  F_C=Replace(F_C,"{$" & FieldName & "}","")
				End If
			End If
		 Next

		ReplaceUserDefine=F_C
	End Function
	
	
	'最新一条广播
	Function GetNewFresh()
		Dim RS:Set RS=Conn.Execute("select top 1 b.id,a.userid,a.username,a.transtime,a.msg,b.adddate,b.copyfrom,b.note,b.cmtnum,b.username as busername,b.userid as buserid,b.transnum,a.type,a.id as rid from ks_userlogr a left join ks_userlog b on a.msgid=b.id where a.status=1 and a.userid=" & userid & " order by a.id desc")
		If RS.Eof And RS.Bof Then
		   RS.Close :Set RS=Nothing
		   GetNewFresh="<strong>广播：</strong><a href='" & KS.GetSpaceUrl(userid) & "'>" & UserName & "</a>没有广播消息，<a href='" & KS.GetDomain& "user/weibo.asp'>广播大厅</a>!"
		   Exit Function
		Else
		   Dim KSR:Set KSR=New refresh
		   GetNewFresh="<strong><a href='" & KS.GetSpaceUrl(userid) & "'>" & UserName & "</a>说：</strong>" & KSR.ReplaceEmot(rs("note")) & " <a href='" & KS.Setting(3)& "user/weibo.asp?userid=" & userid & "'>转播(<span style='color:#ff6600'>" & RS("transnum") & "</span>)</a>" 
		   Set KSR=Nothing
		End If
        RS.Close :Set RS=Nothing
	End Function
	
	
	'论坛话题
	Function GetClubTopic()
		 Dim Xml,Node,Str,RS:Set RS=Conn.Execute("Select top 10 id,Subject,AddTime from KS_GuestBook Where deltf=0 and verific=1 and username='" & username & "' order by id desc")
		 If RS.Eof And RS.Bof Then
		  RS.Close :Set RS=Nothing
		  GetClubTopic="没有发表任何讨论话题,<a href='" & KS.GetClubListUrl(0) & "' target='_blank'>参与</a>！"
		  Exit Function
		 Else
		   Set XML=KS.RsToXml(RS,"row","")
		   RS.Close :Set RS=Nothing
		   If IsObject(XML) Then
		    For Each Node In XML.DocumentElement.SelectNodes("row")
			 str=str & "<img src='../images/default/arrow_r.gif' align='absmiddle'/> <a href='" & KS.GetClubShowUrl(Node.SelectSingleNode("@id").text) & "' target='_blank'>" & Replace(Replace(Node.SelectSingleNode("@subject").text,"{","｛"),"}","｝") &"</a> " & KS.GetTimeFormat(Node.SelectSingleNode("@addtime").text) & "<br/>"
			Next
		   End If
		 End If
		 GetClubTopic=Str
	End Function
	
	 Function GetNewAlbum()
		 Dim Xml,RS:Set RS=Conn.Execute("Select top 4 * from KS_Photozp Where username='" & username & "' order by AddDate Desc,id desc")
		 If RS.Eof And RS.Bof Then
		   RS.Close:Set RS=Nothing
		   GetNewAlbum="没有上传照片！"
		 else
		   Set Xml=KS.RsToXml(RS,"row","")
		   RS.Close:Set RS=Nothing
		   GetNewAlbum=GetAlbum(Xml)
		   Xml=Empty
         end if
	 End Function
		
	 Function GetAlbum(Xml)
	 	 Dim Node,AddDate,N,Url
		   N=0
		   for each Node In Xml.DocumentElement.SelectNodes("row")
		    If N=0 Then  
		     GetAlbum=GetAlbum &"&nbsp;<span class=""titletips""><a class='username' href='../space/?" & userid & "'>" & UserName & "</a> 于" & Node.SelectSingleNode("@adddate").text & "上传的相片! </span> <a class='more' href='../space/?" & userid &"/album'>查看更多 >>></a><br/>"
			End If
			N=N+1
			If KS.SSetting(21)="1" Then
			Url="showalbum-" & userid & "-" & Node.SelectSingleNode("@xcid").text & "-" & n
			Else
			Url="../space/?" & userid &"/showalbum/" & Node.SelectSingleNode("@xcid").text & "/" & n
			End If
			GetAlbum=GetAlbum & "<a style=""border:1px solid #efefef;padding:5px"" class=""zp"" href="""& Url & """ target=""_blank""><img title=""" & Node.SelectSingleNode("@title").text & """ style=""margin-left:6px;margin-top:5px"" src=""" & Node.SelectSingleNode("@photourl").text & """ width=""106"" height=""81"" border=0></a>"
		   Next
		   Set Node=Nothing
	 End Function
	 
	 Function GetStatusStr(val)
           Select Case Val
		    Case 1:GetStatusStr="公开"
			Case 2:GetStatusStr="会员"
			Case 3:GetStatusStr="密码"
			Case 4:GetStatusStr="隐私"
		   End Select
			GetStatusStr="<font color=""red"">" & GetStatusStr & "</font>"
	 End Function
	 Function GetNewLog()
	     Str="  <table border=""0"" align=""center"" width=""100%"">" & vbcrlf
		 Dim RS,Str,i,url
		 Set RS=Server.CreateObject("ADODB.Recordset")
		 RS.Open "Select top 4 * From KS_BlogInfo Where UserName='" & UserName & "' and istalk=0 and status=0 order by id desc",conn,1,1
		 if rs.eof then
		   str=str & "没有写日志！"
		 else
		    i=0
		   do while not rs.eof
		    url=GetLogUrl(RS)
			str=str &"<tr><td class=""splittd"">"
		    if i=0 then
		    str=str & "<span class=""titletips""><a href=""" &  Url & """>" & rs("title") & "</a></span>[" & rs("adddate") & "]"
			str=str &"<div class=""intro"">" & KS.GotTopic(KS.LoseHtml(UbbCode(rs("content"),0)),200) &"... [<a href=""" & Url & """>阅读全文</a>]</div>"
			else
			str=str &"<img src=""" & KS.GetDomain & "images/default/arrow_r.gif"" align=""absmiddle""> <a href='" & url & "' target='_blank'>" & RS("title") & "</a>&nbsp;&nbsp;(" & RS("adddate") & ")<br/>"
			end if
			i=I+1
			str=str &"</td></tr>"
		   rs.movenext
		   loop
		 end if
		 str=str &"</table>"
		 RS.Close:Set RS=Nothing
		 GetNewLog=str
		End Function
		
     Function GetNewXX()
	    Dim SQLStr:SQLStr="Select top 5 ChannelID,InfoID,Title,Tid,Fname,AddDate,Intro from KS_ItemInfo Where Inputer='" & UserName & "' and verific=1 and deltf=0 Order By AddDate Desc,id desc"
		Dim Str,I,url
		Str="  <table border=""0"" align=""center"" width=""100%"">" & vbcrlf
		 Dim RS:Set RS=conn.execute(SQLStr)
		         If RS.EOF and RS.Bof  Then
					str=str & "<tr><td class=""splittd"" colspan=3><p>没有发布任何文档！</p></td></tr>"
				 Else 
				   I=0
					do while not rs.eof 
					      url=KS.GetItemUrl(rs("channelid"),RS("tid"),RS("infoid"),RS("fname"))
					      if i=0 then
		                   str=str &"  <tr><td colspan=3 class=""splittd""><span class=""titletips""><a class='username' href='../space/?" & userid & "'>" & username & "</a> 于" & rs("adddate") & "发表了“<a href='" & url & "' target='_blank'>" & KS.Gottopic(rs(2),28) & "</a>”</span>"
						    IF not KS.IsNul(RS(6)) Then
							 str=str & "<br/><span class=""intro"">" & rs(6) & "...</span>&nbsp;&nbsp;"
							End If
						    str=str & " <a href='../space/?" & userid & "/xx'>查看更多 >>></a></td></tr>" & vbcrlf
						  else
						  str=str & "<tr><td class=""splittd""><img src=""" & KS.GetDomain & "images/default/arrow_r.gif"" align=""absmiddle""> [" & KS.GetClassNP(RS("tid")) & "] <a href='" & url & "' target='_blank'>" & RS(2) & "</a>&nbsp;&nbsp;(" & RS(5) & ")</td></tr>"
						  end if
					      i=I+1
						  rs.movenext
						loop
				 End If	
		 str=str &"  </table>" & vbcrlf
		 rs.close:set rs=nothing
	     GetNewXX=str   
	   
	 End Function
	 
	 
	 Function GetEnterPriseNews()
	   Dim RS,XML,Url,Node:Set RS=Conn.Execute("Select top 10 ID,Title,AddDate From KS_EnterpriseNews where username='" & UserName & "' order by id desc")
	   If Not RS.eof Then Set Xml=KS.RsToXml(RS,"row","")
	   RS.Close:Set RS=Nothing
	   If IsObject(Xml) Then
	     GetEnterPriseNews="<table border='0' cellpadding='0' cellspacing='0'>" & vbcrlf
	   For Each Node In Xml.DocumentElement.SelectNodes("row")
	      If KS.SSetting(21)="1" Then Url= "show-news-"& userid & "-" & Node.SelectSingleNode("@id").text&KS.SSetting(22) Else Url="../space/?" & userid & "/shownews/" & Node.SelectSingleNode("@id").text
	   	   GetEnterPriseNews =GetEnterPriseNews & "<tr><td height='22'><img src='../images/default/arrow_r.gif' align='absmiddle'> <a href=""" & Url & """>" & Node.SelectSingleNode("@title").text & "(" & Node.SelectSingleNode("@adddate").text & ")</a></td></tr>"
	   Next
	     Xml=Empty : Set Node=Nothing
	     GetEnterPriseNews=GetEnterPriseNews & "</table>"
	  End If
	 End Function
	 Function GetSupply()
	   Dim RS:Set RS=Conn.Execute("Select top 10 ID,Title,AddDate,TypeID,Tid,Fname From KS_GQ where verific=1 and inputer='" & UserName & "' order by id desc")
	   If RS.Eof Then RS.Close:Set RS=Nothing:Exit Function
	   Dim I,SQL:Sql=RS.GetRows(-1)
	   RS.Close:Set RS=Nothing
	    GetSupply="<table border='0' cellpadding='0' cellspacing='0'>" & vbcrlf
	   For I=0 To Ubound(SQL,2)
	    GetSupply =GetSupply & "<tr><td height='22'><img src='../images/default/arrow_r.gif' align='absmiddle'>"& KS.GetGQTypeName(SQL(3,I)) & "<a href='" & KS.GetItemUrl(8,SQL(4,I),SQL(0,I),SQL(5,I)) & "' target='_blank'>" & SQL(1,I) &  "(" & SQL(2,I) & ")</a></td></tr>"
	   Next
	    GetSupply=GetSupply & "</table>"
	 End Function
	 Function GetProduct()
	   Dim RS:Set RS=Conn.Execute("Select top 8 ID,Title,PhotoUrl From KS_Product where verific=1 and inputer='" & UserName & "' order by id desc")
	   If RS.Eof Then RS.Close:Set RS=Nothing:Exit Function
	   Dim I,N,k,PhotoUrl,Url,SQL:Sql=RS.GetRows(-1)
	   RS.Close:Set RS=Nothing
	   If IsArray(SQL) Then
	    n=0
	    GetProduct="<div class=""productlist"">" & vbcrlf
	   For I=0 To Ubound(SQL,2)
	     GetProduct =GetProduct & "<ul>"
	     For K=1 To 4
		  PhotoUrl=sql(2,n) : If KS.SSetting(21)="1" Then Url="show-product-" &username & "-" & sql(0,n) & KS.SSetting(22) Else url="?" & UserName & "/showproduct/" & sql(0,n)
		 iF PhotoUrl="" Or IsNull(PhotoUrl) Then PhotoUrl="../Images/nopic.gif"
	      GetProduct =GetProduct & "<li><a href='" & Url & "' target='_blank'><img src='" & PhotoUrl & "' Width=""140"" height=""100"" border=""0""></a><br/><a href='" & Url & "' target='_blank'>"& KS.Gottopic(SQL(1,N),15) & "</a></li>"
		 n=n+1
		 if n> Ubound(SQL,2) Then Exit For
		 Next
		 GetProduct =GetProduct & "</ul>"
		 if n> Ubound(SQL,2) Then Exit For
	   Next
	    GetProduct =GetProduct & "</div>"
	  End If
	 End Function
	 
	 Function GetProductList()
	   Dim RS:Set RS=Conn.Execute("Select top 6 ID,Title,adddate From KS_Product where verific=1 and inputer='" & UserName & "' order by id desc")
	   If RS.Eof Then RS.Close:Set RS=Nothing:Exit Function
	   Dim I,Url,SQL:Sql=RS.GetRows(-1)
	   RS.Close:Set RS=Nothing
	   If IsArray(SQL) Then
	    GetProductList="<table border='0' cellpadding='0' cellspacing='0'>" & vbcrlf
	   For I=0 To Ubound(SQL,2)
	     If KS.SSetting(21)="1" Then Url="show-product-" &userid & "-" & sql(0,i) & KS.SSetting(22) Else url="?" & userid & "/showproduct/" & sql(0,i)
	     GetProductList=GetProductList & "<tr><td height='22'><img src='../images/default/arrow_r.gif' align='absmiddle'> <a href='" & Url & "' target='_blank'>"& SQL(1,i) & "(" & SQL(2,I) & ")</a></td></tr>"
		 GetProductList =GetProductList & "</tr>"
	   Next
	    GetProductList=GetProductList & "</table>"
	  End If
	 End Function
	 
	 Function GetEnterpriseintro()
	   On Error Resume Next
	   GetEnterpriseintro=KS.Htmlcode(Conn.execute("select Intro From KS_EnterPrise where UserName='" & UserName & "'")(0))
	 End Function
	 Function GetEnterpriseShortintro()
	   On Error Resume Next
	   Dim Url
	   If KS.SSetting(21)="1" Then Url="intro-" & username Else Url="../space/?" & username & "/intro"
	  GetEnterpriseShortintro=KS.Gottopic(KS.LoseHtml(KS.Htmlcode(Conn.execute("select Intro From KS_EnterPrise where UserName='" & UserName & "'")(0))),580) &"&nbsp;&nbsp;<a href=""" & Url & """>&nbsp;详细>>></a>"
	 End Function
	 
	 '幻灯显示图片
	 Function GetSlidePhoto(ChannelID)
	  Dim SQL,I,str,picarr
	  Dim RS:Set RS=Server.CreateObject("Adodb.Recordset")
	  RS.Open "Select top 6 ID,Title,Tid,InfoPurview,ReadPoint,Fname,PicUrls From " & KS.C_S(ChannelID,2)  & " Where Inputer='" & UserName & "' order by id desc",conn,1,1
	  If Not RS.Eof Then SQL=RS.GetRows()
	  RS.Close:Set RS=Nothing
	  If IsArray(SQL) Then
	    str="<script src='js/AutoChangePhoto.js'></script><div id=""divcenter_one"">" & vbcrlf
		str=str &"<div class=""divcenter_work_one"">" & vbcrlf
		str=str &"<DIV class=fpic>"

	   For I=0 To Ubound(SQL,2)
			picarr=split(split(SQL(6,I),"|||")(0),"|")
			If I=0 Then
			 str=str & "<A href=""../space/?" & UserName & "/showphoto/" & SQL(0,I) &""" target=""_blank"" id=""foclnk""><img src="""&picarr(1) &""" name=""focpic"" width=""605"" id=focpic style=""FILTER: RevealTrans ( duration = 1，transition=23 ); VISIBILITY: visible; POSITION: absolute"" /></a>" &vbcrlf
			str=str & "<DIV style=""MARGIN-TOP:385px;MARGIN-left:240px;FLOAT:left;WIDTH:120px;TEXT-ALIGN: center;position:absolute""><A href=""../space/?" & UserName & "/xx"" target=_blank><font color=white>更多作品>></font></A></DIV>" &vbcrlf
			
			str=str &"<DIV id=fttltxt style=""MARGIN-TOP:390px;MARGIN-left:250px;FLOAT:left;WIDTH:120px;TEXT-ALIGN: center;position:absolute""></DIV>" &vbcrlf
			str=str & "<DIV style=""MARGIN-LEFT:590px; WIDTH: 65px"">" &vbcrlf
			str=str & "<DIV class=thubpiccur id=tmb0 onmouseover=setfoc(0); onmouseout=playit();><A href=""../space/?" & UserName & "/showphoto/" & SQL(0,I) &""" target=_blank><IMG src=""" & picarr(2) & """ width=32 height=32 border=""0""></A></DIV>" &vbcrlf
            else
			 str=str & "<DIV class=thubpic id=tmb" & i & " onmouseover=setfoc("& I & "); onmouseout=playit();><A href=""../space/?" & UserName & "/showphoto/" & SQL(0,I) &""" target=_blank><img src=""" & picarr(2) & """ width=32 height=32 border=""0""></A></DIV>" &vbcrlf
			end if
	   Next
	   
	   	 str=str & "<SCRIPT language=javascript type=text/javascript>" &vbcrlf
		 str=str &"	var picarry = {};" &vbcrlf
		 str=str &" var lnkarry = {};" & vbcrlf
		 str=str &"	var ttlarry = {};"&vbcrlf
		
		For I=0 To Ubound(SQL,2)
		  picarr=split(split(SQL(6,I),"|||")(0),"|")	
		  str=str &"picarry[" & i & "] = '" & PicArr(1) & "';" & vbcrlf
		  str=str &"lnkarry[" & i & "] = '../space/?" & UserName & "/showphoto/" & SQL(0,I) &"'; "& vbcrlf
		  str=str &"ttlarry[" & i & "] = '';" & vbcrlf
		Next
		 str=str &"</SCRIPT>"
		 str=str &"</DIV>"
		 str=str &"</DIV>"
		 str=str &"</div></div>"
		 GetSlidePhoto=str
	 End If
	End Function
	
	 
	 Function GetEnterpriseContact()
	   Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
	   RS.Open "Select top 1 * From KS_EnterPrise Where UserName='" & UserName & "'",conn,1,1
	   IF RS.Eof Then
	    RS.Close:Set RS=Nothing
		GetEnterpriseContact=""
		Exit Function
	   End If
	   GetEnterpriseContact="联 系 人：" & RS("Contactman") & "<br/>"
	   GetEnterpriseContact=GetEnterpriseContact & "公司地址：" & RS("address") & "<br>"
	   GetEnterpriseContact=GetEnterpriseContact & "邮政编码：" & RS("zipcode") & "<br>"
	   GetEnterpriseContact=GetEnterpriseContact & "联系电话：" & RS("telphone") & "<br>"
	   GetEnterpriseContact=GetEnterpriseContact & "传真号码：" & RS("fax") & "<br>"
	   GetEnterpriseContact=GetEnterpriseContact & "公司网址：" & RS("weburl") & "<br>"
	   RS.Close:Set RS=Nothing
	 End Function
	 
	 '最新访客
	 Function GetVisitor()
	    Dim user_face,Visitors,str,XML,Node
		Dim RS:Set RS=Conn.Execute("Select top 10 b.userid,b.sex,a.Visitors,b.userface,a.adddate,b.isonline from KS_BlogVisitor a inner join KS_User b on a.Visitors=b.username where a.username='" & UserName & "' order by a.adddate desc ,id desc")
				If Not RS.Eof Then Set XML=KS.RsToXml(Rs,"row","")
				RS.Close:Set RS=Nothing
			    If IsObject(XML) Then
				  For Each Node In XML.DocumentElement.SelectNodes("row") 
				    user_face=Node.SelectSingleNode("@userface").text
					Visitors =Node.SelectSingleNode("@visitors").text
					If user_face="" or isnull(user_face) then 
					 if Node.SelectSingleNode("@sex").text="男" then  user_face="images/face/0.gif" else user_face="images/face/girl.gif"
					End If
			        If lcase(left(user_face,4))<>"http" and left(user_face,1)<>"/" then user_face=KS.GetDomain & user_face
			         str=str & "<li><a class='b' href='" & KS.GetSpaceUrl(Node.SelectSingleNode("@userid").text) & "' target='_blank'><img src='" & User_face & "' border='0'></a><br/><a href='" & KS.GetSpaceUrl(Node.SelectSingleNode("@userid").text) & " target='_blank'>" & Visitors & "</a><br />状态:"
					 If Node.SelectSingleNode("@isonline").Text="1" Then str=str & "<font color=red>在线</font>" Else str=str & "离线"
					 str=str & "</li>"
				  Next
				  XML=Empty : Set Node=Nothing
				End If
		 GetVisitor=str
	 End Function
	 
	 
	 '用户信息
	 Function GetUserInfo()
	  Dim str,RS:Set RS=Server.CreateObject("adodb.recordset")
	  rs.open "select top 1 userface,realname,qq from ks_user where username='" & username & "'",conn,1,1
	  if not rs.eof then
	    dim userfacesrc:userfacesrc=rs(0)
		dim realname:realname=KS.CheckXSS(rs(1)&"")
		if realname="" or isnull(realname) then realname=username
	    if left(userfacesrc,1)<>"/" and lcase(left(userfacesrc,4))<>"http" then userfacesrc="../" & userfacesrc
	     str="<div align=""center"" style=""padding-top:5px;"">"_
		 &"<img src=""" & userfacesrc & """ onerror=""this.src='../user/images/noavatar_middle.gif';"" style=""border:0px solid #cccccc;"" width=""170"" height=""190"" border=""0"">"_
		 &"<br /><br />"_
		 &"<div class=""userinfomenu""><li><a href=""../space/?" & UserID & "/message""><img border=""0"" src=""images/yes.gif"" align=""absmiddle""> 给我留言</a></li><li><a href=""javascript:void(0)"" onclick=""ksblog.addF(event,'" & UserName & "');""><img src=""images/adfriend.gif"" border=""0"" align=""absmiddle""> 加为好友</a></li><li> <a href=""javascript:void(0)"" onclick=""ksblog.sendMsg(event,'" & username & "')""><img src=""images/sendmsg.gif"" border=""0"" align=""absmiddle""> 发小纸条</a></li><li>"
		' if rs(2)<>"" and not isnull(rs(2)) then 
		' str=str &"<li><a target=blank href=tencent://message/?uin=" & rs(2) &"&Site=" & KS.Setting(2) & "&Menu=yes><img SRC=http://wpa.qq.com/pa?p=1:" & rs(2) & ":5 alt=""点击这里给我发消息"" border=""0""></a>"
		 'else
		 str=str &"<a href=""../space/?" & username & "/info""><img border=""0"" src=""images/card.gif"" align=""absmiddle""> 小档案</a>"
		' end if
		 str=str & "</li></div></div>"
	  end if
	  rs.close:set rs=nothing
	  GetUserInfo=str
	 End Function

	 'RSS订阅
	 Function GetXML()
	  GetXML="<a href=""rss.asp?UserName=" & UserName &""" target=""_blank""><img src=""../images/xml.gif"" border=""0""></a>"
	 End Function
	 '日历
	 Function Getcalendar()
	  Dim CalCls:Set CalCls=New CalendarCls
	  call CalCls.calendar(Getcalendar,username)
	  Set CalCls=Nothing
	 End Function
	 '搜索
	 Function GetSearch()
	  GetSearch="<table border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbcrlf
	  GetSearch=GetSearch &"<form action=""../space/?" & userid & "/blog"" method=""post"" name=""searchform"">" &vbcrlf
	  GetSearch=GetSearch & "<tr>" & vbcrlf
	  GetSearch=GetSearch & "<td align=""center"">关键字:<input type=""text"" size=""10"" name=""key"" style=""border-style: solid; border-width: 1px""><input type=""submit"" value="" 搜 索 ""></td>" & vbcrlf
	  GetSearch=GetSearch & "</tr>" & vbcrlf
	  GetSearch=GetSearch & "</form>"
	  GetSearch=GetSearch & "</table>" &vbcrlf
	 End Function
     '统计
	 Function GetBlogTotal()
	   GetBlogTotal="日志总数:"&conn.execute("select count(id) from ks_bloginfo where username='" & UserName &"' and status=0")(0) & " 篇"_
	   & "<br />回复总数:"&conn.execute("select count(id) from ks_blogcomment where username='" & UserName &"'")(0) & " 条"_
	   & "<br />留言总数:"&conn.execute("select count(id) from ks_blogmessage where status=1 and username='" & UserName &"'")(0) & " 条"_
	   & "<br />日志阅读:"&conn.execute("select Sum(hits) from ks_blogInfo where username='" & UserName &"' and status=0")(0) & " 人次"_
	   &"<br />总访问数:" & conn.execute("select top 1 hits from ks_blog where username='" & username & "'")(0) & " 人次"
	   
	 End Function
	 '专栏列表
	 Function GetUserClass()
	  Dim Str:Str="<div style='display:none'><form id='myclassform' action='../space/?" & username & "/blog' method='post'><input type='text' name='classid' id='classid'></form></div>"
	  Dim RS:Set RS=Conn.Execute("Select ClassID,ClassName From KS_UserClass Where UserName='" & UserName & "' and TypeID=2")
	  Do While Not RS.Eof 
	    Str=Str & "<a href=""javascript:void(0)"" onclick=""$('#classid').val(" & RS(0) & ");$('#myclassform').submit();"">" & RS(1) & "</a><br>" & vbcrlf
		RS.MoveNext
	  Loop
	  RS.Close:Set RS=Nothing
	  GetUserClass=Str
	 End Function
	 '音乐盒
	 Function GetMusicBox()
	  GetMusicBox="<object classid=""clsid:d27cdb6e-ae6d-11cf-96b8-444553540000""  width=""200"" height=""200"" id=""mp3player"" codebase=""http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=8,0,0,0"" ><param name=""movie"" value=""plus/mp3player.swf?config=plus/config_1.xml&file=plus/getmusiclist.asp?username=" & username & """ /><param name=""allowScriptAccess"" value=""always""><embed src=""plus/mp3player.swf?config=plus/config_1.xml&file=plus/getmusiclist.asp?username=" & username & """ allowScriptAccess=""always"" width=""200"" height=""200"" name=""mp3player""	type=""application/x-shockwave-flash"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" /></object>"
	 End Function
	 Function GetMediaPlayer()
	  on error resume next
	  GetMediaPlayer="<EMBED style=""WIDTH: 272px; HEIGHT: 29px"" src=""" & conn.execute("select top 1 url from ks_blogmusic where username='" & username & "'")(0) & """ width=299 height=10 type=audio/x-wav autostart=""true"" loop=""true""></DIV></EMBED>"
	 End Function
	 '最新日志
	 Function GetBlogInfo()
	  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
	  RS.Open "Select Top " & GetUserBlogParam(UserName,"ListLogNum") & " *  From KS_BlogInfo Where UserName='" & UserName & "' And Status=0 Order By ID Desc",conn,1,1
	  If Not RS.Eof Then
	    Do While Not RS.EOF
		 GetBlogInfo=GetBlogInfo & "<a title=""" & RS("UserName") & "发表于" & RS("AddDate")&""" href=""" &GetCurrLogUrl(RS("UserID"),RS("ID")) & """>" & RS("Title") & "</a><br>" & vbcrlf
		RS.MoveNext
		Loop
	  Else
	   GetBlogInfo="暂无日志!"
	  End If
	  RS.Close:Set RS=Nothing
	 End Function
	 '最新评论
	 Function GetComment()
	  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
	  RS.Open "Select Top " & GetUserBlogParam(UserName,"ListReplayNum") & " *  From KS_BlogComment Where UserName='" & UserName & "' Order By AddDate Desc",conn,1,1
	  If Not RS.Eof Then
	    Do While Not RS.EOF
		 GetComment=GetComment & "<img src=""../images/default/arrow_r.gif"" align=""absmiddle""> <a title=""" & RS("AnounName") & "发表于" & RS("AddDate")&""" href=""" &GetCurrLogUrl(UserID,RS("LogID")) & "#" & RS("ID") &""">" & KS.Gottopic(KS.LoseHtml(RS("Content")),25) & "</a><br />" & vbcrlf
		RS.MoveNext
		Loop
	  Else
	   GetComment="暂无评论!"
	  End If
	  RS.Close:Set RS=Nothing
	 End Function
	 '最新留言
	 Function GetMessage()
	  'GetMessage="<a href=""message.asp?UserName=" & UserName &"#write"">签写留言</a><br>"
	  Dim XML,Node,Url,MaxTop,RS,UserFace,userid,spaceurl
	  Set RS=Server.CreateObject("ADODB.RECORDSET")
	  MaxTop=KS.ChkClng(GetUserBlogParam(UserName,"ListGuestNum"))
	  If MaxTop=0 Then MaxTop=3
	  RS.Open "Select Top " & MaxTop & " m.*,u.userface,u.userid From KS_BlogMessage m left join KS_User u on m.AnounName=u.username Where m.UserName='" & UserName & "' and m.status=1 Order By m.id Desc",conn,1,1
	  If Not RS.Eof Then Set Xml=KS.RsToXml(rs,"row","")
	  RS.Close:Set RS=Nothing
	  If IsObject(Xml) Then
	    GetMessage=GetMessage & "<table width='100%' cellspacing='0' cellpadding='0'>"
	    For Each Node In Xml.DocumentElement.SelectNodes("row")
		  userface=Node.SelectSingleNode("@userface").text
		  If KS.IsNul(UserFace) Then UserFace="images/face/boy.jpg"
		  If lcase(left(userFace,4))<>"http" and left(userface,1)<>"/" then userface=KS.Setting(3) & userface
		  userid=KS.ChkClng(Node.SelectSingleNode("@userid").text)
		  If UserID<>0 Then spaceurl=KS.GetSpaceUrl(UserID) Else spaceurl="#"
		 If KS.SSetting(21)="1" Then Url="message-" & UserName & KS.SSetting(22)&"#"& Node.SelectSingleNode("@id").text Else Url="?" & username & "/message#" & Node.SelectSingleNode("@id").text
		 
		 GetMessage=GetMessage & "<tr><td class='splittd' style='padding:5px 4px 6px 0px;width:55px;text-align:center'><a href='" & SpaceUrl & "' target='_blank'><img style=""padding:2px;border:1px solid #ccc"" width=""50"" height=""50"" src=""" & UserFace & """ alt=""" & Node.SelectSingleNode("@username").text & """></a></td><td class='splittd'><a href='" & SpaceUrl & "' target='_blank' class='username'>" & KS.CheckXSS(Node.SelectSingleNode("@anounname").text) & "</a> 留言于" & Node.SelectSingleNode("@adddate").text 
		 IF KS.C("UserName")=UserName Then
			GetMessage=GetMessage &" <a href='../User/user_message.asp?Action=MessageDel&id=" & Node.SelectSingleNode("@id").text & "' onclick=""return(confirm('确定删除该留言吗?'))"">删除</a> | <a href='../user/user_message.asp?id=" & Node.SelectSingleNode("@id").text & "&Action=ReplayMessage' target='_blank'>回复</a>"
		 End If
		 
		 GetMessage=GetMessage & "<br/><a href=""" & url &""">" & KS.LoseHtml(Node.SelectSingleNode("@content").text) & "</a>"
		 If Not KS.IsNul(Node.SelectSingleNode("@replay").text) Then
			GetMessage=GetMessage&"<div style=""margin : 5px 20px; border : 1px solid #efefef; padding : 5px;background : #ffffee; line-height : normal;""><b>以下为space主人的回复:</b><br>" & KS.LoseHtml(Node.SelectSingleNode("@replay").text) & "<br><div align=right>时间:" & Node.SelectSingleNode("@replaydate").text &"</div></div>"
		 End If
		 GetMessage=GetMessage & "</td></tr>"
		Next
		GetMessage=GetMessage & "</table>"
		Xml=Empty : Set Node=Nothing
	  End If


         If KS.SSetting(25)="0" And KS.IsNul(KS.C("UserName")) Then
		  GetMessage=GetMessage & "<div style=""margin:20px""><strong>温馨提示：</strong>只有会员才可以留言,如果是会员请先<a href=""javascript:ShowLogin()"">登录</a>,不是会员请点此<a href=""../user/reg/"" target=""_blank"">注册</a>。</div>"
		 Else
		 GetMessage=GetMessage & "<div style=""clear:both""></div><a name=""write""></a><table width=""98%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"" class=""comment_write_table"">"
		 GetMessage = GetMessage & "<form name=""msgform"" action=""" & KS.GetDomain &"plus/ajaxs.asp?action=MessageSave"" method=""post"">"
		 GetMessage = GetMessage & "<input type=""hidden"" value=""" & UserName & """ name=""UserName"">"
		 GetMessage = GetMessage & "<input type=""hidden"" value="""" name=""scontent"">"
		 GetMessage = GetMessage & "<tr><td height=""30"" class=""comment_write_title""><strong>快速留言:</strong></td></tr>"
		GetMessage = GetMessage & "<tr>"
		GetMessage = GetMessage & "  <td height=""30"" colspan=""2"">昵称："
		GetMessage = GetMessage & "   <input name=""AnounName"" maxlength=""100"" type=""text"" value=""" & KS.C("UserName") & """ id=""AnounName"" style=""background:#FBFBFB;color:#999;border:1px solid #ccc;width:120""/>&nbsp;<font color=red>*</font> <span>验证码 </span><script>writeVerifyCode("""&KS.GetDomain&""",1)</script></td>"
		GetMessage = GetMessage & "</tr>"
		GetMessage = GetMessage & "  <tr>"
		
		GetMessage = GetMessage & "<td height=""25""><textarea  style=""color:#999;width:98%;border:1px solid #ccc;background:#FBFBFB;overflow:auto"" rows=""4"" id=""Content"" name=""Content"" onfocus=""if (this.value=='既然来了，就顺便留句话儿吧...') this.value='';"" onblur=""if (this.value=='') this.value='既然来了，就顺便留句话儿吧...';"">既然来了，就顺便留句话儿吧...</textarea></td>"
		
		GetMessage = GetMessage & "  </tr>"
		GetMessage = GetMessage & "  <tr>"
		GetMessage = GetMessage & "   <td height=""30""><input type=""button"" onclick=""return(CheckPostMsg());""  name=""SubmitComment"" value=""OK了，提交留言"" class=""btn""/>"
		GetMessage = GetMessage & "    </td>"
		GetMessage = GetMessage & "  </tr>"
		GetMessage = GetMessage & "  </form>"
		GetMessage = GetMessage & "</table>"
       End If

	 End Function
	 '天气
	 Function GetWeather(RS)
	    Dim TitleStr
	    Select Case RS("Weather")
		 Case "sun.gif":TitleStr="晴天"
		 Case "sun2.gif":TitleStr="和煦"
		 Case "yin.gif":TitleStr="阴天"
		 Case "qing.gif":TitleStr="清爽"
	     Case "yun.gif":TitleStr="多云"
		 case "wu.gif":TitleStr="有雾"
		 case "xiaoyu.gif":TitleStr="小雨"
	     case "yinyu.gif":TitleStr="中雨"
		 case "leiyu.gif":TitleStr="雷雨"
		 case "caihong.gif":TitleStr="彩虹"
		 case "hexu.gif":TitleStr="酷热"
		 case "feng.gif":TitleStr="寒冷"
		 case "xue.gif":TitleStr="小雪"
		 case "daxue.gif":TitleStr="大雪"
		 case "moon.gif":TitleStr="月圆"
		 case "moon2.gif":TitleStr="月缺"
		End Select
	 	GetWeather="<img src=""../User/images/weather/" & rs("Weather") & """ title=""" & TitleStr &""" align=""absmiddle"">"
	 End Function
	 
	 Function ReplaceLogLabel(UserName,ByVal TP,RS)
		   Dim EmotSrc:If RS("Face")<>"0" Then EmotSrc="<img src=""../User/images/face/" & RS("Face") & ".gif"" align=""absmiddle"" border=""0"">"
		   Dim MoreStr
		   MoreStr="<a href=""" & GetLogUrl(RS) & """>阅读全文("&RS("hits")&")</a> | <a href=""" & GetLogUrl(RS) & "#Comment"">回复（"& Conn.Execute("Select Count(ID) From KS_BlogComment Where LogID="  &RS("id"))(0) &"）</a>"
		   Dim ContentStr
		    If IsNull(RS("Password")) Or RS("PassWord")="" Then 
			 ContentStr=KS.GotTopic(KS.LoseHtml(RS("Content")),KS.ChkClng(GetUserBlogParam(UserName,"ContentLen")))
			Else
			 ContentStr="<form method='post' action='" & GetLogUrl(RS) & "' target='_blank'>请输入日志的查看密码：<input style='border-style: solid; border-width: 1' type='password' name='pass' size='15'>&nbsp;<input type='submit' value=' 查看 '></form>"
			End IF
			Dim JFStr:If RS("Best")="1" then JFStr="  <img src=""../images/jh.gif"" align=""absmiddle"">" else JFStr=""
		   TP=Replace(TP,"{$ShowLogTopic}",EmotSrc&"<a href=""" & GetLogUrl(RS) & """>" & RS("Title") & "</a>" & jfstr)
		   TP=Replace(TP,"{$ShowLogInfo}","[" & RS("AddDate") & "|by:" & RS("UserName") & "]")
		   TP=Replace(TP,"{$ShowLogText}",ContentStr)
		   TP=Replace(TP,"{$ShowLogMore}",MoreStr)
		   
		   TP=Replace(TP,"{$ShowTopic}",RS("Title"))
		   TP=Replace(TP,"{$ShowAuthor}",RS("UserName"))
		   TP=Replace(TP,"{$ShowAddDate}",RS("AddDate"))
		   TP=Replace(TP,"{$ShowEmot}",EmotSrc)
		   TP=Replace(TP,"{$ShowWeather}",GetWeather(RS))
		   ReplaceLogLabel=TP
		End Function

	 
	 Function Location(str)
	   Location= "<div class=""location"">"
	   Location=Location & str
	   Location=Location & " </div>"
	   'Location=Location & "<hr style=""clear:both"" size=1 color=#cccccc>"
	 End Function
    
	
	'=============================圈子相关标签替换=============================
	 '替换标签
	 Function ReplaceGroupLabel(RS,Template)
	  On Error Resume Next
	  Template=Replace(Template,"{$ShowAnnounce}",RS("Announce"))
	  Template=Replace(Template,"{$ShowNewUser}",GetUserList(RS("id"),"new"))
	  Template=Replace(Template,"{$ShowActiveUser}",GetUserList(RS("id"),"active"))
	  Template=Replace(Template,"{$ShowGroupInfo}",GetGroupInfo(rs))
	  Template=Replace(Template,"{$ShowNavigation}",GetGroupMenu(rs))
	  Template=Replace(Template,"{$ShowGroupName}",RS("TeamName"))
	  Template=Replace(Template,"{$ShowGroupURL}",KS.GetDomain & "space/group.asp?id=" & RS("id"))
	  Template=Replace(Template,"{$ShowUserLogin}","<iframe width=""170"" height=""122"" id=""login"" name=""login"" src=""../user/userlogin.asp"" frameBorder=""0"" scrolling=""no"" allowtransparency=""true""></iframe>")
	  ReplaceGroupLabel=Template
	 End Function
	 
	 '圈子导航
	 Function GetGroupMenu(rs)
	  GetGroupMenu="<li><a href=""group.asp?id=" & rs("id") &""">圈子首页</a></li>"_
				   &"<li><a href=""group.asp?id=" & rs("id") &"&isbest=1"">精华帖子</a></li>"_
				   &"<li><a href=""group.asp?id=" & rs("id") &"&action=users"">成员列表</a></li>"_
				   &"<li><a href=""group.asp?id=" & rs("id") &"&action=join"">加入本圈</a></li>"_
				   &"<li><a href=""group.asp?id=" & rs("id") &"&action=post"">发表话题</a></li>"_
				   &"<li><a href=""group.asp?id=" & rs("id") &"&action=info"">圈子信息</a></li>"_
	 End Function
     
	 '成员列表
	Function GetUserList(teamid,Flag)
	dim orderstr
	If Flag="active" then
	  orderstr=" order by LastLoginTime desc"
	else
	  orderstr=" order by a.id desc"
	end if
	dim rs:set rs=server.createobject("adodb.recordset")
	rs.open "select top 9 a.username,b.userid,b.userface from ks_teamusers a,ks_user b where a.username=b.username and status=3 and teamid="& teamid & orderstr,conn,1,1
	do while not rs.eof
			  Dim UserFaceSrc:UserFaceSrc=rs("UserFace")
			  if lcase(left(userfacesrc,4))<>"http" then userfacesrc=KS.Setting(2) & userfacesrc
	  GetUserList=GetUserList & "<UL class=bestuser>"
	  GetUserList=GetUserList & "<LI class=userimg><a href=""../space/?" & rs("userid") &"""  target=""_blank""><img src=""" & userfacesrc & """ width=""60"" height=""60""></a></li>"
	  GetUserList=GetUserList & "<LI class=username><A href=""../space/?" & rs("userid") & """ target=""_blank"">" & rs("username") & "</a></LI>"
	  GetUserList=GetUserList & "</UL>"
	rs.movenext
	loop
	End Function

    Function GetGroupInfo(rs)
	    GetGroupInfo="<img src=""" & rs("photourl") & """ border=""0"" width=""160"" height=""150"">"_
		             &"<br />圈子名称：" & rs("teamname")_
					 &"<br />创 建 者：" & rs("username")_
					 &"<br />创建时间：" & rs("adddate")_
					 &"<br />成员人数：" & conn.execute("select count(id) from ks_teamusers where status=3 and teamid=" & rs("id"))(0)_
					 &"<br />主题回复：" & conn.execute("select count(id) from ks_teamtopic where teamid=" & rs("id") & "and parentid=0")(0) & "/" &conn.execute("select count(id) from ks_teamtopic where teamid=" & rs("id") & "and parentid<>0")(0) _
	End Function
	'=============================圈子相关标签替换结束==========================

End Class
%> 

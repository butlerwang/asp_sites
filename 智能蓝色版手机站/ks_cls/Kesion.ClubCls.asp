<!--#include file="../Plus/Session.asp"-->
<%


Class ClubCls
        Private KS, KSR,ListStr,Node,BSetting,KSUser,GuestTitle,Master,MasterArr,FileContent,TopicID
		Private ListTemplate,pLoopTemplate,LoopTemplate,LoopTemplate1,LoopList,boardid,parentId,PostBtnStr,TopXML,TopicXml,TopicNode
		Private MaxPerPage, TotalPut , CurrentPage, TotalPage, i, j, Loopno,Immediate,Templates,ListType
	    Private SqlStr,Doc,ListUrl,startime,LoginTF,CachePage,CacheTime
		Private Sub Class_Initialize()
		 CachePage=true  '首页缓存,访问量或是数据量较大时,建议设置成true
		 CacheTime=0     '首页缓存时间设置,单位为分钟
		 If (Not Response.IsClientConnected)Then
			Response.Clear
			Response.End
		 End If
		  Immediate = true
		  Set KS=New PublicCls
		  
		  If KS.Setting(69)<>"" then  '如果绑定二级域名，访问不是二级域名时跳到二级域名下
		    if instr(lcase(KS.GetCurrentUrl),lcase(KS.Setting(69)))=0 then
			 Response.Redirect "http://" & KS.Setting(69)
			end if
		  End If
		  
		  Set KSR = New Refresh
		  Set KSUser=New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSUser=Nothing
		 Set KSR=Nothing
		End Sub
		%>
		<!--#include file="Kesion.IfCls.asp"-->
		<!--#include file="ClubFunction.asp"-->
		<%

		Public Sub Kesion()
		    startime=Timer()
			If KS.Setting(56)="0" Then Call KS.ShowTips("error","本站已关闭论坛功能!") 
			FCls.RefreshType = "guestindex" '设置刷新类型，以便取得当前位置导航等
			If Not KS.IsNul(Request.QueryString) Then Call LoadClubBoardList Else Call LoadClubIndex
			GetClubPopLogin FileContent
			FileContent=KSR.ReplaceGeneralLabelContent(FileContent)
			FileContent=Replace(Replace(FileContent,"｛#","{"),"#｝","}")  '标签替换回来
			FileContent=RexHtml_IF(FileContent)
			FileContent=Replace(FileContent,"{#ExecutTime}","页面执行" & FormatNumber((timer()-startime),5,-1,0,-1) & "秒 powered by CMS")
			If KS.Setting(59)="1" and ks.chkclng(boardid)=0 Then
			 Scan FileContent
			Else
			 KS.Echo FileContent
			End If
		End Sub
		
		'加载首页模板
		Sub LoadTemplate()
				Application(KS.SiteSn &"ClubIndexUpdateTime")=Now
				FileContent = KSR.LoadTemplate(KS.Setting(114))
				FileContent=KSR.ReplaceAllLabel(FileContent)
				FileContent=KSR.ReplaceLableFlag(FileContent)
				Application(KS.SiteSN&"ClubIndex")=FileContent
		End Sub
		'主页
		Sub LoadClubIndex()
		    If KS.Setting(114)="" Then KS.Die "请先到""基本信息设置->模板绑定""进行模板绑定操作!"
			FCls.RefreshFolderID = 0
			If CachePage=false Or KS.ChkCLng(CacheTime)=0 Or KS.IsNUL(Application(KS.SiteSN&"ClubIndex")) Or Not isDate(Application(KS.SiteSn &"ClubIndexUpdateTime")) Then
			   LoadTemplate
			ElseIf  isDate(Application(KS.SiteSn &"ClubIndexUpdateTime") And  DateDiff("n",Application(KS.SiteSn &"ClubIndexUpdateTime"),Now)>=KS.ChkCLng(CacheTime)) Then
			    LoadTemplate()
			Else
			    FileContent = Application(KS.SiteSN&"ClubIndex")
			End If
			 
			 If KS.Setting(59)="1" Then 
			  Call GetIndexList
			 Else
			  KS.LoadClubBoard : Call GetBoardList()
			 End If
			 ListTemplate = LoopList
			set Doc = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			Doc.async = false
			Doc.setProperty "ServerHTTPRequest", true 
			Doc.load(Server.MapPath(KS.Setting(3)&"Config/guestbook.xml"))
			dim loadnum:loadnum=0
			do while Doc.parseError.errorCode<>0   '出错重新加载
			 Doc.load(Server.MapPath(KS.Setting(3)&"Config/guestbook.xml"))
			 loadnum=loadnum+1
			 if loadnum>10 then exit do
			loop
			if loadnum>10 then ks.die "加载数据出错，请重新刷新页面试试或稍候访问!"
			
			Dim XMLDate:XMLDate=doc.documentElement.attributes.getNamedItem("date").text
			If DateDiff("d",xmldate,now)=0 Then
					  If KS.ChkClng(doc.documentElement.attributes.getNamedItem("todaynum").text)>KS.ChkClng(doc.documentElement.attributes.getNamedItem("maxdaynum").text) then
					   doc.documentElement.attributes.getNamedItem("maxdaynum").text=doc.documentElement.attributes.getNamedItem("todaynum").text
					   doc.save(Server.MapPath(KS.Setting(3)&"Config/guestbook.xml"))
					  end if
			Else
					  GCls.Execute("Update KS_GuestBoard Set TodayNum=0")
				      Application(KS.SiteSN&"_ClubBoard")=empty	
					  doc.documentElement.attributes.getNamedItem("date").text=now
					  doc.documentElement.attributes.getNamedItem("yesterdaynum").text=doc.documentElement.attributes.getNamedItem("todaynum").text
					  doc.documentElement.attributes.getNamedItem("todaynum").text=0
					  doc.save(Server.MapPath(KS.Setting(3)&"Config/guestbook.xml"))
			End If
	
			FileContent=Replace(FileContent,"{$TodayNum}",doc.documentElement.attributes.getNamedItem("todaynum").text)
			FileContent=Replace(FileContent,"{$YesterDayNum}",doc.documentElement.attributes.getNamedItem("yesterdaynum").text)
			FileContent=Replace(FileContent,"{$MaxDayNum}",doc.documentElement.attributes.getNamedItem("maxdaynum").text)
			FileContent=Replace(FileContent,"{$TopicNum}",doc.documentElement.attributes.getNamedItem("topicnum").text)
			FileContent=Replace(FileContent,"{$ReplayNum}",doc.documentElement.attributes.getNamedItem("postnum").text)
			FileContent=Replace(FileContent,"{$UserNum}",doc.documentElement.attributes.getNamedItem("totalusernum").text)
			FileContent=Replace(FileContent,"{$NewUser}",doc.documentElement.attributes.getNamedItem("newreguser").text)
			FileContent=Replace(FileContent,"{$MaxOnline}",doc.documentElement.attributes.getNamedItem("maxonline").text)
			FileContent=Replace(FileContent,"{$MaxOnlineDate}",doc.documentElement.attributes.getNamedItem("maxonlinedate").text)
			PostBtnStr="<a href=""javascript:Posted()""><img src=""" & KS.Setting(3) & KS.Setting(66) & "/images/button_post.png"" align=""absmiddle"" alt=""发帖""></a>"
			FileContent=Replace(FileContent,"{$PostButtonAction}",PostBtnStr)
			FileContent=Replace(FileContent,"{$GuestTitle}",KS.Setting(61))
			FileContent=Replace(FileContent,"{$GetGuestList}",ListTemplate)
		End Sub
	    '版面
		Sub LoadClubBoardList()
		   If Not KS.IsNul(KS.Setting(69)) and Request.QueryString<>"" Then
					  Dim QueryStr:QueryStr=Request.QueryString
					  Dim QArr:QArr=Split(Split(QueryStr,".")(0),"-")
					  If Ubound(Qarr)>=1 Then
					   BoardID=KS.ChkClng(Qarr(1))
					  Else
					   BoardID=KS.ChkClng(KS.S("BoardID"))
					  End If
					  If Ubound(QArr)>=2 Then  
					   CurrentPage = KS.ChkClng(Qarr(2))
					  Else
					   CurrentPage = KS.ChkClng(Request("page")) 
					  End If
			Else
					  BoardID=KS.ChkClng(KS.S("BoardID"))
					  CurrentPage = KS.ChkClng(Request("page")) 
			End If
			If KS.ChkClng(BoardID)=0 Then LoadClubIndex() : Exit Sub  '没有传递版本ID时，转到首页
			FCls.RefreshFolderID = BoardID '设置当前刷新目录ID 为"0" 以取得通用标签
		    KS.LoadClubBoard
			Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]")
			If Node Is Nothing Then LoadClubIndex() : Exit Sub  '没有找到版本参数时，转到首页
		   If KS.Setting(172)="" Then KS.Die "请先到""基本信息设置->模板绑定""进行模板绑定操作!"
		   FileContent = KSR.LoadTemplate(KS.Setting(172))
			
			BSetting=Node.SelectSingleNode("@settings").text
			ParentId=KS.ChkClng(Node.SelectSingleNode("@parentid").text)
			FileContent=Replace(FileContent,"{$BoardName}",Node.SelectSingleNode("@boardname").text)
			FileContent=Replace(FileContent,"{$GetBoardUrl}",KS.GetClubListUrl(boardid))
			Master=Node.SelectSingleNode("@master").text
			
             BSetting=BSetting&"$$$0$0$0$0$0$0$0$0$$$$$$$$$$$$$$$$"
			 BSetting=Split(BSetting,"$")
			 If CurrentPage<=0 Then CurrentPage=1
			 MaxPerPage=KS.ChkClng(BSetting(20)) : If MaxPerPage=0 Then MaxPerPage=KS.ChkClng(KS.Setting(51))

			 If Not KS.IsNul(KS.Setting(69)) Then
			  ListUrl="http://" & KS.Setting(69) & "/"
			 Else
			  ListUrl=KS.GetDomain & KS.Setting(66) &"/"
			 End If
			
			 LoginTF=KSUser.UserLoginChecked
			If parentid<>0 or KS.S("Istop")="1" or KS.S("IsBest")="1" then
			    If BSetting(0)="0" Then  '不允许游客浏览时才进一步判断权限 
				 Dim CheckResult:CheckResult=CheckPermissions(KSUser,BSetting,GuestTitle) '检查访问检查
				 If CheckResult="true" Then Call ShowBoardList Else ListTemplate=CheckResult
				Else
				 Call ShowBoardList
			    End If
			Else
				 KS.LoadClubBoard : Call GetBoardList()
				 ListTemplate=LoopList
			End IF
			
				FileContent=RexHtml_IF(FileContent) '列表页先过滤其它标签,减少标签解释
				FileContent=Replace(FileContent,"{$GuestTitle}",GuestTitle)
				FileContent=Replace(FileContent,"{$GetGuestList}",ListTemplate)
				FileContent=Replace(Replace(Replace(Replace(Replace(FileContent,"{$ShowManageCheckBox}",""),"{$Img}",""),"{$PageList}",""),"{$Jing}",""),"{$Status}","") '替换掉无用标签,加快解释
				FileContent=KSR.ReplaceAllLabel(FileContent)
				FileContent=KSR.ReplaceLableFlag(FileContent)
                FileContent=Replace(FileContent,"{$BoardID}",boardid)
		End Sub	
		
		Sub ShowBoardList()
		     session("clubnowboardpage")=1
		     if boardid<>0  Then
			     session("clubnowboardpage")=request("page")
				GuestTitle=KS.LoseHtml(Node.SelectSingleNode("@boardname").text)
			 else
				if KS.S("Istop")="1" then GuestTitle="置顶帖子" Else GuestTitle="精华帖子"
			 end if
				PostBtnStr="<span style=""position:relative;"" onmouseover=""$('#postlist').show()"" onmouseout=""$('#postlist').hide()""><a href=""" & KS.Setting(3) & KS.Setting(66) & "/post.asp?bid=" & boardid & """><img src=""" & KS.Setting(3) & KS.Setting(66) & "/images/button_post.png""></a><div id=""postlist"" class=""submenu noli"">"
				PostBtnStr=PostBtnStr&"<dl><img src=""" & KS.Setting(3) & KS.Setting(66) & "/images/new_post.gif"" align=""absmiddle""/> <a href=""" & KS.Setting(3) & KS.Setting(66) & "/post.asp?bid=" & boardid & """>发表帖子</a></dl>"
				If KS.ChkClng(bsetting(64))>0 Then
				PostBtnStr=PostBtnStr &"<dl><img src=""" & KS.Setting(3) & KS.Setting(66) & "/images/vote.gif"" align=""absmiddle""> <a href=""" & KS.Setting(3) & KS.Setting(66) & "/post.asp?bid=" & BoardID&"&posttype=1"">发起投票</a></dl>"
				End If
				PostBtnStr=PostBtnStr &"</div></span>"
				Call GetLoopList()
				'ks.echo "页面执行" & FormatNumber((timer()-startime),5,-1,0,-1) & "秒<br/>"
				GetClubPopLogin FileContent
				'ks.echo "页面执行" & FormatNumber((timer()-startime),5,-1,0,-1) & "秒<br/>"
				FileContent=Replace(FileContent,"{$PostButtonAction}",PostBtnStr)						   
				FileContent=Replace(FileContent,"{$GuestTitle}",GuestTitle)
				FileContent=RexHtml_IF(FileContent) '先过滤无用的标签,减少标签解释
				FileContent=KSR.KSLabelReplaceAll(FileContent)
				'ks.echo "页面执行" & FormatNumber((timer()-startime),5,-1,0,-1) & "秒"

				if instr(FileContent,"{#GetClubPopLogin}")<>0 Then GetClubPopLogin FileContent
				Scan FileContent
			 ks.die ""
		End Sub
		
		'首页列出帖子
		Sub GetIndexList()
		  LoopList="<table border=""0"" style=""margin:0px auto;width:98%"" align=""center"" class=""glist"" cellpadding=""0"" cellspacing=""0"">" & vbcrlf & "<thead class=""category"">"
		  LoopList=LoopList & "<tr><td style=""width:30px;text-align:center""></td><td class=""banmian"">主题</td><td style=""width:80px;text-align:center"">作者</td><td style=""width:60px;text-align:center"">回复</td><td style=""width:150px;text-align:center"">最后发表↓</td></tr></thead></table>"
		   BSetting="$$$0$0$0$0$0$0$0$0$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$"
		   BSetting=Split(BSetting,"$")
		   MaxPerPage=KS.ChkClng(KS.Setting(51)) : If MaxPerPage=0 Then MaxPerPage=20
		   LoopList=LFCls.GetConfigFromXML("GuestBook","/guestbook/template","indexlist")
		   Call GetLoopList()
		   LoopList=Replace(LoopList,"{$ListType}",ListType)
		End Sub
		
		Function Parse(sTemplate, iPosBegin)
			Dim iPosCur, sToken, sValue, sTemp
			iPosCur        = InStr(iPosBegin, sTemplate, "}")
			sTemp        = Mid(sTemplate,iPosBegin,iPosCur-iPosBegin)
			iPosBegin    = iPosCur+1
			iPosCur       = InStr(sTemp, ".")
			if iPosCur>1 Then
			sToken        = Left(sTemp, iPosCur-1)
			End If
			sValue        = Mid(sTemp, iPosCur+1) 
		
			Select Case lcase(sValue)
				Case "begin"
					sTemp            = "{@" & ( sToken & ".end}" )
					iPosCur            = InStr(iPosBegin, sTemplate, sTemp)
					ParseArea      sToken, Mid(sTemplate, iPosBegin, iPosCur-iPosBegin)
					iPosBegin        = iPosCur+Len(sTemp)
				case "boardid" echo boardid
				case "boardname" echo Node.SelectSingleNode("@boardname").text
				case "boardintro" echo Node.SelectSingleNode("@note").text

				case "master"
				    If KS.IsNul(Master) Then 
					  Echo "<a href='#'>暂无版主</a>"
					Else
					 If Not IsObject(Application(KS.SiteSN &"Master"&BoardID)) Then
					   Call LoadMasterUserID(BoardID,Master)
					 End If
					 Dim MyMaster:MyMaster=Application(KS.SiteSN &"Master"&BoardID)
					 If Not KS.IsNul(MyMaster) Then
						 MasterArr=Split(MyMaster,"@") 
						 For I=0 To Ubound(MasterArr)
						   If I=0 Then echo "<a href='" & KS.GetSpaceUrl(Split(MasterArr(i),"|")(0)) & "' target='_blank'>" & Split(MasterArr(i),"|")(1) & "</a>" Else echo "," & "<a href='" & KS.GetSpaceUrl(Split(MasterArr(i),"|")(0)) & "' target='_blank'>" & Split(MasterArr(i),"|")(1) & "</a>"
						 Next
					 End If
					End If
			   case "topicnum"  echo Node.SelectSingleNode("@topicnum").text
			   case "todaynum"  echo Node.SelectSingleNode("@todaynum").text
			   case "boardrules" echo Node.SelectSingleNode("@boardrules").text
			   case "executtime" echo "页面执行" & FormatNumber((timer()-startime),5,-1,0,-1) & "秒 powered by CMS"
			   case "showpage"
			    If Not KS.IsNul(Request("a")) or Not KS.IsNul(Request("c")) or Not KS.IsNul(Request("d"))  or Not KS.IsNul(Request("o")) or Not KS.IsNul(Request("isbest")) or Not KS.IsNul(Request("istop")) Then
				   echo KS.ShowPage(TotalPut,MaxPerPage,"",CurrentPage,false,false)
				Else
				   If KS.IsNul(KS.Setting(69)) Then
				    echo KS.GetClubPageList(0,MaxPerPage,CurrentPage,TotalPut,KS.ChkClng(BoardID),"/" & Gcls.ClubPreList)
				   else
				    echo KS.GetClubPageList(0,MaxPerPage,CurrentPage,TotalPut,KS.ChkClng(BoardID),Gcls.ClubPreList)
				   end if
				End If
				Case Else
					ParseNode sToken, sValue
		   End Select
		   Parse    = iPosBegin
		End Function
        Sub ParseArea(sTokenName, sTemplate)
			Select Case sTokenName
			  Case "toploop"
			    LoadTopTopic
				Dim TopParam
				If KS.ChkClng(BoardID)=0 Then TopParam="@istop=2" Else TopParam="@boardid=" & Boardid&" or @istop=2"
			    If CurrentPage=1 And IsObject(Application(KS.SiteSN &"TopXML")) Then
				  For Each TopicNode In Application(KS.SiteSN &"TopXML").DocumentElement.SelectNodes("row[" &TopParam& "]")
				     TopicID=TopicNode.SelectSingleNode("@id").text
					 scan sTemplate
				  Next
				  echo "<table border=""0"" style=""margin:0px auto;width:99%"" align=""center"" class=""topiclist"" cellpadding=""0"" cellspacing=""0""><tr><td style=""background:#E6F2FB;height:25px;padding-left:15px"">===普通主题===</td></tr></table>"
				End If
			  Case "loop"
			    If IsObject(TopicXML) Then
				  For Each TopicNode In TopicXML.DocumentElement.SelectNodes("row")
				     TopicID=TopicNode.SelectSingleNode("@id").text
					 scan sTemplate
				  Next
				End If
				
			End Select
		End Sub
		Sub ParseNode(sTokenType, sTokenName)
					Select Case lcase(sTokenType)
					    case "item"
						  select case lcase(sTokenName)
						    case "userid" echo TopicNode.SelectSingleNode("@userid").text
						    case "ico" 
							  dim IcoUrl,TitleTips
							  If KS.ChkClng(TopicNode.SelectSingleNode("@posttype").text)=1 Then
			                   IcoUrl="vote.gif" : TitleTips="投票主题"
							  ElseIf cint(TopicNode.SelectSingleNode("@istop").text)=1 Then
							   IcoUrl="top.gif" : TitleTips="本版面置顶"
							  ElseIf cint(TopicNode.SelectSingleNode("@istop").text)=2 Then
							   IcoUrl="ztop.gif": TitleTips="总置顶"
							  ElseIf cint(TopicNode.SelectSingleNode("@verific").text)=2 Then
							   IcoUrl="lock.gif": TitleTips="屏闭主题"
							  ElseIf KS.ChkClng(TopicNode.SelectSingleNode("@hits").text)>KS.ChkClng(BSetting(27)) and KS.ChkClng(TopicNode.SelectSingleNode("@totalreplay").text)>KS.ChkClng(BSetting(28)) Then
							   IcoUrl="hot.gif": TitleTips="热门主题"
							  Else
							   IcoUrl="common.gif": TitleTips="普通主题"
							  End If
							  echo "<a href='" & KS.GetClubShowUrl(TopicID) &"' target='_blank'><img border='0' src='" & KS.Setting(3) & KS.Setting(66) & "/images/" & IcoUrl & "' title='" & TitleTips & "'></a>"
							case "author" 
							  Dim PostUser:PostUser=TopicNode.SelectSingleNode("@username").text
							  If KS.IsNul(PostUser) Then
							   echo "<a href=""#"" class=""author"" target=""_blank"">游客</a>"
							  Else
							   echo "<a href=""" & KS.GetSpaceUrl(TopicNode.SelectSingleNode("@userid").text) & """ class=""author"" target=""_blank"">" & PostUser& "</a>"
							  End If
							case "pubtime" echo KS.GetTimeFormat(TopicNode.SelectSingleNode("@addtime").text)
							case "replaytimes" echo TopicNode.SelectSingleNode("@totalreplay").text
							case "hits" echo TopicNode.SelectSingleNode("@hits").text 
							case "lastreplayuser"
							  dim LastReplayUser:LastReplayUser=TopicNode.SelectSingleNode("@lastreplayuser").text
							  If KS.IsNul(LastReplayUser) Then
							   echo "<a href=""#"" target=""_blank"">游客</a>"
							  Else
							   echo "<a href=""" & KS.GetSpaceUrl(TopicNode.SelectSingleNode("@lastreplayuserid").text) & """ class=""author"" target=""_blank"">" & LastReplayUser& "</a>"
							  End If
							case "lastreplaytime" echo KS.GetTimeFormat1(TopicNode.SelectSingleNode("@lastreplaytime").text,true)
							case "subjectlist"
							   If KS.S("A")="m" Then echo "<input type='checkbox' name='m' onclick=""showmanage(this.checked,this.value,'" & KS.Setting(66) & "'," & BoardID & ")"" value='" & TopicID & "'/>"
							   If KS.ChkClng(BSetting(25))>0 and isobject(Application(KS.SiteSN&"_ClubBoardCategory")) Then
								Dim CategoryNode,CategoryId,categoryName,categoryIco
								CategoryId=TopicNode.SelectSingleNode("@categoryid").text
								Set CategoryNode=Application(KS.SiteSN&"_ClubBoardCategory").DocumentElement.SelectSingleNode("row[@categoryid=" & CategoryId&"]")
								If Not CategoryNode Is Nothing Then
								categoryname=CategoryNode.SelectSingleNode("@categoryname").text : If Instr(categoryname,"[")=0 and categoryname<>"" Then categoryname="<span class=""scategory"">[" & categoryname & "]</span>"
								categoryIco=CategoryNode.SelectSingleNode("@ico").text
									If KS.ChkClng(BSetting(25))=2 Then
									echo " <a href=""" & ListUrl & "?boardid=" & boardid& "&c=" &CategoryId&"""><Img src='" & categoryIco & "' alt='" &CategoryName & "' border='0' align='absmiddle'/></a>"
									Else
									echo "<a href=""" & ListUrl & "?boardid=" & boardid& "&c=" &CategoryId&""">" & CategoryName &"</a>"
									End If
								End If
							  End If
							  
							   If KS.Setting(59)="1" and request("boardid")="" Then   '首页直接显示帖子
								KS.LoadClubBoard
								Dim BNode:Set BNode=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & TopicNode.SelectSingleNode("@boardid").text &"]")
								If Not BNode Is Nothing Then
								 echo "[<a href='" & KS.GetClubListUrl(TopicNode.SelectSingleNode("@boardid").text) & "'>" & BNode.SelectSingleNode("@boardname").text &"</a>] "
								End If
							  End If	
							  
							  echo "<a "
							  if KS.ChkClng(BSetting(65))<>0 Then echo " class=""topiclink"""
							  echo " href=""" & KS.GetClubShowUrl(TopicID) & """ title=""" & KS.LoseHtml(replace(replace(TopicNode.SelectSingleNode("@subject").text,"｛#","{"),"#｝","}")) & """>" & replace(replace(TopicNode.SelectSingleNode("@subject").text,"｛#","{"),"#｝","}") & "</a>"
							  If cint(TopicNode.SelectSingleNode("@showscore").text)>0 then 
							   echo "<span class=""sj""> - [售价：<span style='color:red'>" &TopicNode.SelectSingleNode("@showscore").text & "</span> "
							   Dim CurrChargeType
							   If TopicNode.SelectSingleNode("@istop").text<>"0" Then
							    Dim CurrentNode:Set CurrentNode=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & TopicNode.SelectSingleNode("@boardid").text &"]")
			                    If Not CurrentNode Is Nothing Then 
								  CurrChargeType=KS.ChkClng(Split(CurrentNode.SelectSingleNode("@settings").text&"$$$0$0$0$0$0$0$0$0$$$$$$$$$$$$$$$$","$")(56))
								Else
								  CurrChargeType=KS.ChkClng(Bsetting(56))
							    End If
								Set CurrentNode=Nothing
							   Else
							     CurrChargeType=KS.ChkClng(Bsetting(56))
							   End If
								   Select Case CurrChargeType
									  case 0 echo KS.Setting(46) &KS.Setting(45)
									  case 1 echo "元人民币"
									  case 2 echo "个积分"
									End Select
							   echo "]</span>"
							  End If
							  
							  If KS.ChkClng(TopicNode.SelectSingleNode("@hits").text)>KS.ChkClng(BSetting(27)) and KS.ChkClng(TopicNode.SelectSingleNode("@totalreplay").text)>KS.ChkClng(BSetting(28)) Then echo " <img align='absmiddle' src='" & KS.Setting(3) & KS.Setting(66) & "/images/hot_1.gif' title='热门'/>"
							  if cint(TopicNode.SelectSingleNode("@verific").text)=0 Then
							   echo " <span style='color:red'>[未审]</span>"
							  ElseIf cint(TopicNode.SelectSingleNode("@verific").text)=2 Then
							   echo " <span style='color:green'>[屏闭]</span>"
							  End If
							  If cint(TopicNode.SelectSingleNode("@isbest").text)=1 Then echo "<Img src='" & KS.Setting(3) & KS.Setting(66) & "/images/jing.gif' border='0' alt='精华帖子' align='absmiddle'/> "
							  Dim AnnexExt,TotalReplay,MaxPage,pages,K
							  AnnexExt=TopicNode.SelectSingleNode("@annexext").text
							  If Not KS.IsNul(AnnexExt) Then
			                   echo " <Img src='" & KS.Setting(3) & "editor/ksplus/fileicon/" & AnnexExt &".gif' alt='" & AnnexExt & "附件' border='0' align='absmiddle'/>"
			                  Else
								  If KS.ChkClng(TopicNode.SelectSingleNode("@ispic").text)=1 Then
									echo " <Img src='" & KS.Setting(3) & KS.Setting(66) & "/images/image_s.gif' alt='Gif图片附件' border='0' align='absmiddle'/>"
								  ElseIf KS.ChkClng(TopicNode.SelectSingleNode("@ispic").text)=2 Then
									echo " <Img src='" & KS.Setting(3) & KS.Setting(66) & "/images/image_s.gif' alt='JPG图片附件' border='0' align='absmiddle'/>"
								  End If
							  End If
							  
							  '主题边分页
							  TotalReplay=KS.ChkClng(TopicNode.SelectSingleNode("@totalreplay").text)
							  If TotalReplay<>0 Then
							     MaxPage=KS.ChkClng(BSetting(21)) : If MaxPage=0 Then MaxPage=10
								 If TotalReplay Mod MaxPage = 0 Then
										Pages=TotalReplay\MaxPage
								 Else
										Pages=TotalReplay\MaxPage + 1
								 End If
							   If Pages>1 Then
									    echo "<span class=""topic-pages""><img src='" &KS.Setting(3) & KS.Setting(66) & "/images/multipage.gif' title='分页'/>"
										if pages>5 then
										   echo " <a href='" & KS.GetClubShowUrlPage(TopicID,1) & "'>1</a><a href='" & KS.GetClubShowUrlPage(TopicID,2) & "'>2</a>... <a href='" & KS.GetClubShowUrlPage(TopicID,pages-3) & "'>" & pages-3 &"</a> <a href='" & KS.GetClubShowUrlPage(TopicID,pages-2) & "'>" & pages-2 &"</a> <a href='" & KS.GetClubShowUrlPage(TopicID,pages-1) & "'>" & pages-1 &"</a> <a href='" & KS.GetClubShowUrlPage(TopicID,pages) & "'>" & pages &"</a>"
										Else
										   For k=1 to Pages
											 echo " <a href='" & KS.GetClubShowUrlPage(TopicID,k) & "'>"&k&"</a>"
										   Next
										End If
								   echo "</span>"
								End if
							  End If
							  If KS.ChkClng(BSetting(42))<>0 and isdate(TopicNode.SelectSingleNode("@lastreplaytime").text) Then
							   If DateDiff("h",TopicNode.SelectSingleNode("@lastreplaytime").text,now)<=KS.ChkClng(BSetting(42)) Then
							  echo " <img src='" &KS.Setting(3) & KS.Setting(66) & "/images/new.gif' />"
							   End If
							  End If
						  end select
					end select
		End Sub
		'列出版面
		Sub GetBoardList()
		  Dim LC,PNode,Node,Xml,Str,TStr,Bparam,LastPost,LastPost_A,S_Style,S_Num,S_N
          Set Xml=Application(KS.SiteSN&"_ClubBoard")
		  If parentid=0 and boardid<>0 Then BParam="id=" & boardid Else BParam="parentid=0"
		  If IsObject(xml) Then
		       PLoopTemplate=LFCls.GetConfigFromXML("GuestBook","/guestbook/template","boardclass")
		       LoopTemplate=LFCls.GetConfigFromXML("GuestBook","/guestbook/template","board")
			   LoopTemplate1=LFCls.GetConfigFromXML("GuestBook","/guestbook/template","board1")
			   For Each Pnode In Xml.DocumentElement.SelectNodes("row[@" & BParam & "]")
					 LC=PLoopTemplate
					 GuestTitle=PNode.SelectSingleNode("@boardname").text
					 LC=Replace(LC,"{$BoardUrl}",KS.GetClubListUrl(PNode.SelectSingleNode("@id").text))
					 LC=replace(LC,"{$BoardID}",PNode.SelectSingleNode("@id").text)
					 LC=replace(LC,"{$BoardName}",PNode.SelectSingleNode("@boardname").text)
					 LC=replace(LC,"{$Intro}",PNode.SelectSingleNode("@note").text)
					 If KS.IsNul(PNode.SelectSingleNode("@master").text) then
					 LC=replace(LC,"{$Master}","暂无版主")
					 else
					 LC=replace(LC,"{$Master}",PNode.SelectSingleNode("@master").text)
					 end if
					 LC=replace(LC,"{$TotalSubject}",PNode.SelectSingleNode("@topicnum").text)
					 LC=replace(LC,"{$TotalReply}",PNode.SelectSingleNode("@postnum").text)
					 LC=replace(LC,"{$TodayNum}",PNode.SelectSingleNode("@todaynum").text)
                     S_Style=KS.ChkClng(Split(PNode.SelectSingleNode("@settings").text&"$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$","$")(52))
					 If s_style<>0 Then s_style=1
					 If KS.IsNUL(request.cookies("clubdis_"&PNode.SelectSingleNode("@id").text)) Or request.cookies("clubdis_"&PNode.SelectSingleNode("@id").text)="0" Then
					 LC=replace(LC,"{$OpenTF}",1) : LC=replace(LC,"{$OpenStyle}",""):LC=replace(LC,"{$OpenICO}","close")
					 Else
					 LC=replace(LC,"{$OpenTF}",0) : LC=replace(LC,"{$OpenStyle}","style=""display:none"""):LC=replace(LC,"{$OpenICO}","open")
					 
					 End If
					 If Request("BoardID")<>"" Then S_Style=0 
                     S_Num=KS.ChkClng(Split(PNode.SelectSingleNode("@settings").text&"$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$","$")(52)) : If S_Num<=0 Then S_Num=3
					 
					 tstr="": S_N=0
					 If S_Style=1 Then Tstr="<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%""><tr>"
				   For Each Node In Xml.DocumentElement.SelectNodes("row[@parentid=" & Pnode.SelectSingleNode("@id").text & "]")
				     If S_Style=1 Then str=LoopTemplate1 Else str=LoopTemplate
					 str=Replace(str,"{$BoardUrl}",KS.GetClubListUrl(Node.SelectSingleNode("@id").text))
					 str=replace(str,"{$BoardID}",Node.SelectSingleNode("@id").text)
					 str=replace(str,"{$BoardName}",Node.SelectSingleNode("@boardname").text)
					 str=replace(str,"{$Intro}",Node.SelectSingleNode("@note").text)
					 str=replace(str,"{$PhotoUrl}",Split(Node.SelectSingleNode("@settings").text&"$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$","$")(51))
					 If KS.IsNul(Node.SelectSingleNode("@master").text) then
					 str=replace(str,"{$Master}","暂无版主")
					 else
					 str=replace(str,"{$Master}",Node.SelectSingleNode("@master").text)
					 end if
					 
					 LastPost=Node.SelectSingleNode("@lastpost").text
					 If KS.IsNul(LastPost) Then
					  str=Replace(Replace(Replace(replace(str,"{$NewTopic}","无"),"{$LastPostUrl}","#"),"{$LastPostUser}","无"),"{$LastPostTime}","-")
					 Else
					  LastPost_A=Split(LastPost,"$")
					  If LastPost_A(0)="0" or LastPost_A(2)="无" then
					  str=Replace(Replace(Replace(replace(str,"{$NewTopic}","无"),"{$LastPostUrl}","#"),"{$LastPostUser}","无"),"{$LastPostTime}","-")
					  else
					  str=replace(str,"{$LastPostUrl}",KS.GetClubShowUrl(LastPost_A(0)))
					  str=replace(str,"{$NewTopic}","<a href='" & KS.GetClubShowUrl(LastPost_A(0)) & "'>" & KS.gottopic(KS.LoseHtml(Replace(LastPost_A(2),"{","｛#")),38) & "</a>")
					  str=replace(str,"{$LastPostUser}","<a href='" & KS.GetSpaceUrl(KS.ChkClng(LastPost_A(4))) &"' target='_blank'>" &LastPost_A(3) & "</a>")
					  str=replace(str,"{$LastPostTime}",KS.GetTimeFormat1(LastPost_A(1),true))

					  end if
					 End If

					 str=replace(str,"{$TotalSubject}",Node.SelectSingleNode("@topicnum").text)
					 str=replace(str,"{$TotalReply}",Node.SelectSingleNode("@postnum").text)
					 str=replace(str,"{$TodayNum}",Node.SelectSingleNode("@todaynum").text)
					 If S_Style=1 Then  '首页版面横排
					  If S_N Mod S_Num=0 And S_N<>0 Then Tstr=Tstr & "</tr><tr class=""board_row"">"
					  TStr=TStr&"<td class=""board_g"" width=""" & round(100/S_Num) & "%"">" & str & "</td>"
					 Else
					  TStr=TStr&str
					 End If
					 S_N=S_N+1
				  Next
				    If S_Style=1 Then 
					 for i=S_N Mod S_Num to S_Num
					  'TStr=TStr&"<td class=""board_g"" width=""" & round(100/S_Num) & "%"">&nbsp;</td>"
					 Next
					 Tstr=Tstr & "</tr></table>"
					End If
					 If Not KS.IsNul(PNode.SelectSingleNode("@note").text) And KS.IsNul(Request.QueryString) Then tstr =tstr & PNode.SelectSingleNode("@note").text
					LC=Replace(LC,"<!--boardlist-->",tstr)
				  LoopList=LoopList & LC
			 Next
		  End If
		End Sub
		'列出帖子
		Sub GetLoopList()
		    Dim Param
			Dim OrderArr:OrderArr=Array("默认排序|0|0","帖子ID号↓|1|0","帖子ID号↑|1|1","浏 览 数↓|2|0","回复时间↓|0|0","回复时间↑|0|1","浏 览 数↓|2|0","浏 览 数↑|2|1","回 帖 数↓|3|0","回 帖 数↑|3|1")
			Dim DateArr:DateArr=Array("全部时间|0","一天|1","三天|3","一周内|7","一个月内|30","三个月内|90","半年内|180","一年内|365")
		    If KSUser.GetUserInfo("ClubSpecialPower")="1" or KSUser.GetUserInfo("ClubSpecialPower")="2" Or (KSUser.GetUserInfo("ClubSpecialPower")="3" and KS.FoundInArr(Master,KSUser.UserName,",")=true) Then 
			 Param=" Where deltf=0"	
			 if KS.S("A")="m" then
			   FileContent=Replace(FileContent,"{$ShowManageButton}","<a href=""" & KS.GetClubListUrl(boardid) & """><img src=""" & KS.Setting(3) & KS.Setting(66) & "/images/button_manage.png""></a>")
			 else
			   FileContent=Replace(FileContent,"{$ShowManageButton}","<a href=""" & ListUrl & "?page=" & currentpage &"&a=m&boardid=" & BoardID & """><img src=""" & KS.Setting(3) & KS.Setting(66) & "/images/button_manage.png""></a>")
			 end if
			Else  
			 Param=" Where deltf=0 and verific<>0"
			 if KS.S("a")="m" then
			   KS.Die "<script>alert('您没有管理的权限,请不要非法访问!');history.back(-1);</script>"
			 end if
			End If
			
			
			ListType="<li>主题：</li>"
			if request.querystring.count=1 then
			 ListType=ListType & "<li class=""current""><a href='" & ListUrl &"?boardid=" & boardid & "'>全部</a></li>"
			else
			 ListType=ListType & "<li><a href='" & ListUrl &"?boardid=" & boardid & "'>全部</a></li>"
			end if
			If KS.ChkClng(KS.S("Istop"))=1 Then 
			 Param=Param & " and istop<>0"
			 ListType=ListType & "<li class=""current""><a href='" & ListUrl &"?boardid=" & boardid & "&istop=1'>置顶</a></li>"
			Else
			 ListType=ListType & "<li><a href='" & ListUrl &"?boardid=" & boardid & "&istop=1'>置顶</a></li>"
			End If
			If KS.ChkClng(KS.S("IsBest"))=1 Then 
			 Param=Param & " and isbest=1"
			 ListType=ListType & "<li class=""current""><a href='" & ListUrl &"?boardid=" & boardid & "&isbest=1'>精华</a></li>"
			Else
			 ListType=ListType & "<li><a href='" & ListUrl &"?boardid=" & boardid & "&isbest=1'>精华</a></li>"
			End If
			ListType=ListType & "&nbsp;&nbsp;<li>| &nbsp;&nbsp; </li>"
            
			Dim D:D=KS.ChkClng(KS.S("D"))
			Dim O:O=KS.ChkClng(KS.S("O"))
			Dim C:C=KS.ChkClng(KS.S("C"))
			'按时间查看
			ListType=ListType & "<li style=""position:relative;_padding-top:6px"" onmouseover=""$('#datelist').show()"" onmouseout=""$('#datelist').hide()"">" & vbcrlf
			if d<=Ubound(DateArr) Then
			  ListType=ListType & "<a href=""#"">" & split(DateArr(d),"|")(0) & " <img src=""" & KS.Setting(3) & KS.Setting(66) &"/images/arw_d2.gif"" align=""absmiddle""/></a>" & vbcrlf
			  If D<>0 Then Param=Param & " and datediff(" & DataPart_D & ",AddTime," & SQLNowString &")<" & split(DateArr(d),"|")(1)
			Else
			ListType=ListType & "<a href=""#"">全部时间 <img src=""" & KS.Setting(3) & KS.Setting(66) &"/images/arw_d2.gif"" align=""absmiddle""/></a>" & vbcrlf
			End If
			ListType=ListType & "<div id=""datelist"" class=""submenu"" style=""left:0px;"">" & vbcrlf
			For I=0 To Ubound(DateArr)
			  ListType=ListType & "<dl><a href=""" & ListUrl & "?boardid=" & boardid & "&d=" & I & """>" & Split(DateArr(i),"|")(0) &"</a></dl>"
			Next
			ListType=ListType & "</div></li>" & vbcrlf
			ListType=ListType & "&nbsp;&nbsp;<li>| &nbsp;&nbsp; </li>"
			'排序方式
			ListType=ListType & "<li style=""position:relative;_padding-top:6px"" onmouseover=""$('#orderlist').show()"" onmouseout=""$('#orderlist').hide()"">" & vbcrlf
			if O<=Ubound(OrderArr) Then
			  ListType=ListType & "<a href=""#"">" & split(OrderArr(o),"|")(0) & " <img src=""" & KS.Setting(3) & KS.Setting(66) &"/images/arw_d2.gif"" align=""absmiddle""/></a>" & vbcrlf
			Else
			ListType=ListType & "<a href=""#"">默认排序 <img src=""" & KS.Setting(3) & KS.Setting(66) &"/images/arw_d2.gif"" align=""absmiddle""/></a>" & vbcrlf
			End If
			ListType=ListType & "<div id=""orderlist"" class=""submenu"" style=""left:0px;"">" & vbcrlf
			For I=0 To Ubound(OrderArr)
			  ListType=ListType & "<dl><a href=""" & ListUrl & "?boardid=" & boardid & "&o=" & I & """>" & Split(OrderArr(i),"|")(0) &"</a></dl>"
			Next
			ListType=ListType & "</div></li>" & vbcrlf
			
		    FileContent=Replace(FileContent,"{$ListType}",ListType)
			
			'版面分类
			If BSetting(23)="1" And BSetting(26)="1" Then
			  KS.LoadClubBoardCategory
			  Dim CategoryNode,CategoryXML,CategoryStr,categoryImg
			  Set CategoryXML=Application(KS.SiteSN&"_ClubBoardCategory").DocumentElement.SelectNodes("row[@boardid=" &BoardID &"]")
			  If CategoryXML.length>0 Then 
				  CategoryStr="<p class=""boardcategory cl"">" & vbcrlf
				  If C=0 Then
				   CategoryStr=CategoryStr & "<strong class=""otp brw"">全部</strong>" &vbcrlf
				  Else
				   Param=Param & " and categoryId=" & KS.ChkClng(KS.S("C"))
				   CategoryStr=CategoryStr & "<a href='" & KS.GetClubListUrl(boardid) & "' class='brw'>全部</a>" &vbcrlf
				  End If
				  For Each CategoryNode In CategoryXML
				   If CategoryNode.SelectSingleNode("@ico").text<>"" Then
				   categoryImg="<img class=""vm"" src=""" & CategoryNode.SelectSingleNode("@ico").text &""" /> "
				   Else
				   categoryImg=""
				   End If
				   If trim(C)=trim(CategoryNode.SelectSingleNode("@categoryid").text) Then
				  CategoryStr=CategoryStr & "<strong class=""otp brw"">" & categoryImg & CategoryNode.SelectSingleNode("@categoryname").text & "</strong>" &vbcrlf
				   Else
				     CategoryStr=CategoryStr & "<a href=""" & ListUrl & "?boardid=" & boardid & "&c=" &CategoryNode.SelectSingleNode("@categoryid").text &""" class=""brw"">" & categoryImg & CategoryNode.SelectSingleNode("@categoryname").text & "</a>"
				   End If

				  Next
				  CategoryStr=CategoryStr &"</p>"
			  End If
		      FileContent=Replace(FileContent,"{$BoardCategory}",CategoryStr)
			  
			End If

		  If BoardID<>0 Then Param=Param &" and boardid=" & boardid
          
		  Dim RS,ListTopicFields
		  ListTopicFields="ID,UserName,UserID,Subject,AddTime,Verific,LastReplayUser,LastReplayUserID,LastReplayTime,TotalReplay,BoardID,Hits,IsPic,IsTop,IsBest,PostType,AnnexExt,CategoryId,ShowScore" rem 主题列表用到的字段

		  Param=Param & " and istop=0"
		  If DataBaseType=1 Then
			    Dim Cmd : Set Cmd = Server.CreateObject("ADODB.Command")
				Set Cmd.ActiveConnection=conn
				Cmd.CommandText="KS_ClubsList"
				Cmd.CommandType=4	
				CMD.Prepared = true 	
				Cmd.Parameters.Append cmd.CreateParameter("@pagenow",3)
				Cmd.Parameters.Append cmd.CreateParameter("@pagesize",3)
				Cmd.Parameters.Append cmd.CreateParameter("@inConditions",200,1,220)
				Cmd.Parameters.Append cmd.CreateParameter("@ListFields",200,1,220)
				Cmd.Parameters.Append cmd.CreateParameter("@inOrder",3)
				Cmd.Parameters.Append cmd.CreateParameter("@inSort",3)
				Cmd("@pagenow")=CurrentPage
				Cmd("@pagesize")=MaxPerPage
				Cmd("@inConditions")=param
				Cmd("@ListFields")=ListTopicFields
				Cmd("@inOrder")=split(OrderArr(o),"|")(1)
				Cmd("@inSort")=split(OrderArr(o),"|")(2)
				Set Rs=Cmd.Execute
				Set Cmd=Nothing
				totalPut=GCls.Execute("Select Count(1) From KS_GuestBook " & Param)(0)
				If Not RS.Eof Then Set TopicXML=KS.RsToXml(RS,"row","")
		  Else
			 Dim OrderField,SortStr
			 Select Case split(OrderArr(o),"|")(1)
			  case 1 OrderField="Id"
			  case 2 OrderField="hits"
			  case 3 OrderField="TotalReplay"
			  case else OrderField="LastReplayTime"
             End Select
			 If split(OrderArr(o),"|")(2)=0 Then SortStr="Desc" Else SortStr="ASC"
	 
			 If CurrentPage=1 Then
			  SqlStr = "SELECT Top " & MaxPerPage & " " & ListTopicFields & " From KS_GuestBook " & Param &" ORDER BY IsTop Desc," & OrderField & " " & SortStr 
			 Else
			  SqlStr = "SELECT " & ListTopicFields & " From KS_GuestBook " & Param &" ORDER BY " & OrderField & " " & SortStr  
			 End If
			 Set RS=GCls.Execute(sqlstr)
			 IF RS.Eof And RS.Bof Then
				  totalput=0
				  LoopList = "<tr><td colspan=5>此版面没有" & KS.Setting(62) & "!</td></tr>"
				  exit sub
			  Else
								TotalPut=GCls.Execute("Select Count(1) From KS_GuestBook " & Param)(0)
								If CurrentPage < 1 Then CurrentPage = 1
								If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
									RS.Move (CurrentPage - 1) * MaxPerPage
								End If
								Set TopicXML=KS.ArrayToXml(RS.GetRows(MaxPerPage),rs,"row","")
				End IF
		 End If	
		   RS.Close:Set RS=Nothing
		End Sub
		
End Class


Class ClubDisplayCls
        Private KS, KSR,ListStr,ID,Node,managestr,TotalReplay,TreplayNum,PostTable
		Private ListTemplate,LoopTemplate,LoopList,FileContent,RST,master,PostType,CheckIsMaster
		Private MaxPerPage, TotalPut , CurrentPage, TotalPage, i, j, Loopno,ShowScore,IsBest,IsTop,DelTF,Verific,Subject,Hits
	    Private SqlStr,GuestTitle,AllowShow,CategoryID,CategoryNode,categoryname,startime,IsClose
		Public UserFields,PostUserName,PostUserID,BSetting,N,KSUser,LoginTF,TopicID,BoardID
		Public ReplayID,XML,TopicNode,UserXML,CommentXML,Un,Immediate,Templates
		Private LC,UserNames,PIDS,RS,ChannelID,InfoID
		Private re
		Private Sub Class_Initialize()
		 If (Not Response.IsClientConnected)Then
			Response.Clear
			Response.End
		 End If
		 Immediate = true
		 UserFields="UserID,UserName,UserFace,Sign,Sex,Score,Prestige,BestTopicNum,LoginTimes,RegDate,email,qq,postNum,LastLoginTime,ClubGradeID,IsOnline,LockOnClub,Medal,issfzrz,MsgNum,FansNum"
		  Set KS=New PublicCls
		 
		  If KS.Setting(69)<>"" then  '如果绑定二级域名，访问不是二级域名时跳到二级域名下
		    if instr(lcase(KS.GetCurrentUrl),lcase(KS.Setting(69)))=0 then
			 Response.Redirect "http://" & KS.Setting(69)
			end if
		  End If

		  Set KSUser=New UserCls
		  Set KSR = New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		%>
		<!--#include file="Kesion.IfCls.asp"-->
		<!--#include file="ClubFunction.asp"-->
		<%
		Function Parse(sTemplate, iPosBegin)
			Dim iPosCur, sToken, sValue, sTemp
			iPosCur        = InStr(iPosBegin, sTemplate, "}")
			sTemp        = Mid(sTemplate,iPosBegin,iPosCur-iPosBegin)
			iPosBegin    = iPosCur+1
			iPosCur       = InStr(sTemp, ".")
			if iPosCur>1 Then
			sToken        = Left(sTemp, iPosCur-1)
			End If
			sValue        = Mid(sTemp, iPosCur+1) 
		
			Select Case lcase(sValue)
				Case "begin"
					sTemp            = "{@" & ( sToken & ".end}" )
					iPosCur            = InStr(iPosBegin, sTemplate, sTemp)
					ParseArea      sToken, Mid(sTemplate, iPosBegin, iPosCur-iPosBegin)
					iPosBegin        = iPosCur+Len(sTemp)
				case "subject" echo Replace(Replace(subject,"｛#","{"),"#｝","}")
				case "subjectnohtml" echo KS.CheckXSS(KS.LoseHtml(Replace(Replace(Replace(subject,"'","\'"),"｛#","{"),"#｝","}")))
				case "description" 
				 If IsObject(Xml) Then
				 Set TopicNode=Xml.DocumentElement.SelectSingleNode("row[@parentid='0']/@content")
				  If Not TopicNode Is   Nothing Then  echo KS.Gottopic(KS.LoseHtml(Replace(Ubbcode(topicnode.text,0),chr(10),"")),150)
				 End If
				case "hits" echo hits
				case "totalreplay" 
				 If KS.ChkClng(totalreplay)>0 Then echo totalreplay-1 Else Echo 0
				case "guesttitle" echo guesttitle
				case "topicid" echo TopicID
				case "boardid" echo boardid
				case "boardurl" Echo KS.GetClubListUrl(BoardID)
				case "posttable" echo PostTable
		        
				
				case "executtime" echo "页面执行" & FormatNumber((timer()-startime),5,-1,0,-1) & "秒 powered by CMS"
				case "boardcategory"
						   If CategoryID<>0 Then
							   KS.LoadClubBoardCategory
							   Set CategoryNode=Application(KS.SiteSN&"_ClubBoardCategory").DocumentElement.SelectSingleNode("row[@categoryid=" &CategoryID &"]")
							   If Not CategoryNode Is Nothing Then
							   categoryname=CategoryNode.SelectSingleNode("@categoryname").text : If Instr(categoryname,"[")=0 Then categoryname="[" & categoryname & "]"
								   If CheckIsMaster Then
								   echo "<a href=""javascript:void(0)"" onclick=""mcategory('" & subject & "','" & KS.Setting(66) & "'," & boardid & "," & TopicID & "," & CategoryID & ")"" title='点击可以修改帖子归类' id='category'>" & categoryname & "</a>"
								   Else
									 echo "<span id='category'>" & categoryname & "</span>"
								   End If
							   End If
							   Set CategoryNode=Nothing
						   Else
						      If CheckIsMaster And BSetting(23)="1" Then
								   echo "<a href=""javascript:void(0)"" onclick=""mcategory('" & subject & "','" & KS.Setting(66) & "'," & boardid & "," & TopicID & "," & CategoryID & ")"" title='点击可以修改帖子归类' id='category'>[设归类]</a>"
							  End If
						   End If
				case "managemenu"
					If CheckIsMaster Then
					   echo "<div class=""backlist"">"
					  If ChannelID=0 And InfoID=0 Then   '没有绑定模型的可以推送
					   echo "<a href=""javascript:void(0)"" onclick=""topicpush(" & id & ",'" & KS.Setting(66) & "'," & BoardID&",'" & KS.LoseHtml(subject) & "')"">主题推送</a> | "
					  End If
					   echo "<a href=""javascript:void(0)"" onclick=""topicfav(" & id & ",'" & KS.Setting(66) & "'," & BoardID&")"">收藏主题</a> | "
					  if verific=1 then
						echo "<a href=""javascript:void(0)"" onclick=""lockorunlock(0,"&id &",'" & KS.Setting(66) & "'," & BoardID&")"">屏闭主题</a> | "
					  else
						echo "<a href=""javascript:void(0)"" onclick=""lockorunlock(1,"&id &",'" & KS.Setting(66) & "'," & BoardID&")"">解除屏闭</a> | "
					  end if
					  if IsClose=1 then
						echo "<a href=""javascript:void(0)"" onclick=""openorclose(0,"&id &",'" & KS.Setting(66) & "'," & BoardID&")"">打开主题</a> | "
					  else
						echo "<a href=""javascript:void(0)"" onclick=""openorclose(1,"&id &",'" & KS.Setting(66) & "'," & BoardID&")"">关闭主题</a> | "
					  end if
					  
						echo "<a href=""javascript:void(0)"" onclick=""delsubject("&id &",'" & KS.Setting(66) & "'," &boardid&")"">删除帖子</a> | <a href=""javascript:void(0)"" onclick=""movetopic('" & KS.Setting(66) & "'," & id & ",'" & KS.LoseHtml(subject) & "')"">移动帖子</a> | "
					  if istop<>0 then
						echo "<a href='javascript:void(0)' onclick=""canceltop(" & ID & ",'" & KS.Setting(66) & "',"&boardid &");"">取消置顶</a> | "
					  else
						echo "<a href='javascript:void(0)' onclick=""settop(" & ID & ",'" & KS.Setting(66) & "',"&boardid &",1);"">设为置顶</a> | <a href='javascript:void(0)' onclick=""settop(" & ID & ",'" & KS.Setting(66) & "',"&boardid &",2);"">设为总置顶</a> | "
					  end if
					  if isbest=1 then
						echo "<a href='javascript:void(0)' onclick=""cancelbest(" & ID & ",'" & KS.Setting(66) & "',"&boardid &");"">取消精华</a> | "
					  else
						echo "<a href='javascript:void(0)' onclick=""setbest(" & ID & ",'" & KS.Setting(66) & "',"&boardid &");"">设为精华</a> | "
					  end if
					echo "</div>" 
				  End If 
				  case "turnto"
					 If TreplayNum>2 Then
							   Echo "转到：<input type='text' style='background:transparent;width:30px;border:0px;border-bottom:1px solid #ccc;text-align:center;' name='tofloor' id='tofloor' value='' size='1'/>&nbsp;<input type='button' value='GO' class='btn' onclick=""TurnToFloor('" & KS.Setting(3) & KS.Setting(66) & "'," & TreplayNum & "," & MaxPerPage & "," & TopicID &");"" style='padding:0px;width:22px'/>"
					 End If
				Case "jing"
						  If CurrentPage=1 Then
							 If isbest=1 Then
								echo "<img style='float:right;right:195px;position:absolute' src='"  &KS.GetDomain & KS.Setting(66) & "/images/jh.gif' align='absmiddle' alt=""本贴被认定为精华"" title=""本贴被认定为精华"">"
							 End If
							 If IsTop<>0 Then
								echo "<img style='right:50px;float:right;position:absolute' src='"  &KS.GetDomain & KS.Setting(66) & "/images/zd.gif' align='absmiddle' alt=""本贴被置顶显示"" title=""本贴被置顶显示"">"
							 End If
							End If
				case "showpage"
						   If AllowShow=true Then
							If KS.IsNul(Request.QueryString("UserName")) Then
							  echo KS.GetClubPageList(BoardID,MaxPerPage,CurrentPage,TotalPut,TopicID,GCls.ClubPreContent)
							Else
							  echo KS.ShowPage(TotalPut,MaxPerPage,"",CurrentPage,false,false)
							End If
						   End If
				Case Else
					ParseNode sToken, sValue
			End Select 
			Parse    = iPosBegin
		End Function 
		
		Sub ParseArea(sTokenName, sTemplate)
					Select Case sTokenName
						Case "loop"
						      Application(KS.SiteSN&"LoopTemplate"&BoardID)=sTemplate
							  If IsObject(XML) Then
								 For Each TopicNode In Xml.DocumentElement.SelectNodes("row")
									 If IsObject(UserXML) Then set UN=UserXml.DocumentElement.SelectSingleNode("row[@username='" & lcase(TopicNode.SelectSingleNode("@username").text) & "']")  Else Set UN=Nothing
									  n=n+1
									  ReplayID=TopicNode.SelectSingleNode("@id").text
									  scan sTemplate
									 I=I+1
									 
								 Next
									Set Un=Nothing
							   
							  End If
						case "replay"
							If KSUser.GetUserInfo("LockOnClub")="1" Or IsClose=1 Then Exit Sub
							If KS.Setting(54)<>"3" And LoginTF=false Then Exit Sub
							If BSetting(62)<>"" And BSetting(62)<>"0" Then 
							  If KS.FoundInArr(BSetting(62),KSUser.GroupID,",")=false Then Exit Sub
							End If
							If BSetting(0)="0" Then  '要权限的版块求
							 If CheckPermissions(KSUser,BSetting,GuestTitle)<>"true" then Exit Sub
							End If
							sTemplate=Replace(Replace(sTemplate,"{#InstallDir#}",KS.Setting(3)),"{#ClubDir#}",KS.Setting(66))
							scan sTemplate
						   
					End Select 
		End Sub 
		Sub ParseNode(sTokenType, sTokenName)
					Select Case lcase(sTokenType)
					    case "item"
						  select case lcase(sTokenName)
						     case "n" echo n
						     case "floor" echo GetFloor(n)
						     case "pubtime" echo KS.GetTimeFormat1(TopicNode.SelectSingleNode("@replaytime").text,true)
							 case "pubip"
							    Select Case KS.ChkClng(KS.Setting(58))
								   case 1 
									If KSUser.GetUserInfo("ClubSpecialPower")="1" Then echo "Post IP：" & TopicNode.SelectSingleNode("@userip").text
								   case 2
									If KSUser.GetUserInfo("ClubSpecialPower")="1" Or KSUser.GetUserInfo("ClubSpecialPower")="2" Or CheckIsMaster=true Then echo "Post IP：" & TopicNode.SelectSingleNode("@userip").text
								   case 3
									 If TopicNode.SelectSingleNode("@showip").text="1" And KSUser.GetUserInfo("ClubSpecialPower")<>1 and CheckIsMaster=false and TopicNode.SelectSingleNode("@username").text<>KS.C("UserName") Then
									 Else
									  echo "Post IP：" & TopicNode.SelectSingleNode("@userip").text
									 End If
								  End Select
							 case "showauthoronly"
							     If Request.QueryString("UserName")="" Then
			                      echo " | <a href='" & KS.Setting(3) & KS.Setting(66) & "/display.asp?id=" & TopicID &"&username=" & TopicNode.SelectSingleNode("@username").text &"'>只看该作者</a>"
								  Else
								  echo " | <a href='" & KS.GetClubShowUrl(TopicID)&"'>显示全部帖子</a>"
								  End If
								  Echo " <a href='" &KS.GetDomain & "space/?" & PostUserID &"/club' target='_blank'>查看该作者主题</a>"
							 case "username" echo TopicNode.SelectSingleNode("@username").text
							 case "userid" echo TopicNode.SelectSingleNode("@userid").text
							 case "spaceurl" echo KS.GetSpaceUrl(TopicNode.SelectSingleNode("@userid").text)
							 case "onlineico"
							   If UN Is Nothing Then Exit Sub
							   If UN.SelectSingleNode("@isonline").text="1" Then
			                     echo "<img src='" & KS.GetDomain & "user/images/online.gif' title='当前在线' align='absmiddle'/>"
			                   Else
			                     echo "<img src='" & KS.GetDomain & "user/images/notonline.gif' title='当前不在线' align='absmiddle'/>"
			                   End If
							 case "showbshare" 
							 if N=1 Then
							  echo "<tr><td class=""topicleft"" style=""border-bottom:none"">&nbsp;</td><td style=""padding-left:10px;height:20px""><a class=""bshareDiv"" href=""#"">分享到</a><script language=""javascript"" type=""text/javascript"" src=""http://static.bshare.cn/b/button.js#uuid=8a5892db-a8f6-4b91-b5dd-93753bd581aa&style=2&textcolor=#000&bgcolor=none&bp=qqmb,sinaminiblog,sohubai,renren&ssc=false&sn=true&text=分享到""></script></td></tr>"
							  End If
							 case "usersignandbottomad"
							    
							    If UN Is Nothing Then 
								Sign=""
								ElseIf TopicNode.SelectSingleNode("@showsign").text="1" Then 
								 Sign=UN.SelectSingleNode("@sign").text
								Else
								 Sign=""
								End If
								Sign=KS.FilterIllegalChar(Sign)
							      Dim BottomAD:BottomAD=GetAdByRnd(37)
								  If BottomAD<>"" Then
								   If Sign<>"" Then Sign="<div class=""usersign"">" & KS.CheckXss(Sign) &"</div>"
								   Sign=Sign & "<div class=""bottomad"">" & BottomAD &"</div>"
								  End If
								  If Sign<>"" THEN echo "<tr><td class=""topicleft"" style=""border-bottom:none"">&nbsp;</td><td>" & Ubbcode(sign,n) &"</td></tr>"
							 case "bottomsimplemenu"
								  If Not KS.IsNul(KS.C("UserName")) Then
									  If (N=1 And BSetting(46)="1") Or (N>1 And BSetting(47)="1") Then
									  echo "<img src='" &KS.Setting(3) & KS.Setting(66) &"/images/Icon_2.gif' align='absmiddle'> <a onclick=""comments('" & KS.Setting(66) &"'," & topicid & "," & replayid & "," & boardid & "," & n & "," & PostUserID &")"" href='javascript:void(0);'>点评</a> | "
									  End If
									 If TopicNode.SelectSingleNode("@verific").text="1" Then echo "<img src='" &KS.Setting(3) & KS.Setting(66) &"/images/repquote.gif' align='absmiddle'> <a href='#reply' onclick=""reply("&n&",'" & TopicNode.SelectSingleNode("@username").text & "','" & TopicNode.SelectSingleNode("@replaytime").text & "')"">引用</a> "
								 End If
							 case "quoteandreply"
							  If Not KS.IsNul(KS.C("UserName")) Then
							      If (N=1 And BSetting(46)="1") Or (N>1 And BSetting(47)="1") Then
								  echo "<img src='" &KS.Setting(3) & KS.Setting(66) &"/images/Icon_2.gif' align='absmiddle'> <a onclick=""comments('" & KS.Setting(66) &"'," & topicid & "," & replayid & "," & boardid & "," & n & "," & PostUserID &")"" href='javascript:void(0);'>点评</a> | "
								  End If
								 If TopicNode.SelectSingleNode("@verific").text="1" Then echo "<img src='" &KS.Setting(3) & KS.Setting(66) &"/images/repquote.gif' align='absmiddle'> <a href='#reply' onclick=""reply("&n&",'" & TopicNode.SelectSingleNode("@username").text & "','" & TopicNode.SelectSingleNode("@replaytime").text & "')"">引用</a> | <img src='" &KS.Setting(3) & KS.Setting(66) &"/images/fastreply.gif' align='absmiddle'> <a href='#reply' >回复</a> | "
							 End If
							 echo "<img src='" & KS.Setting(3) & "images/good.gif'><a href=""javascript:void(0)"" onclick=""support(" & TopicID & ","& ReplayID &",'" & KS.Setting(66) &"')"">支持(<span style='color:red' id='supportnum" &ReplayID&"'>" & KS.ChkClng(TopicNode.SelectSingleNode("@support").text) & "</span>)</a> | <img src='" & KS.Setting(3) & "images/bad.gif'><a href=""javascript:void(0)"" onclick=""opposition(" & TopicID & ","& ReplayID &",'" & KS.Setting(66) &"')"">反对(<span style='color:#999999' id='oppositionnum" & ReplayID & "'>" & KS.ChkClng(TopicNode.SelectSingleNode("@opposition").text) & "</span>)</a>"
							 case "topicmanagemenu"
							   If CheckIsMaster Then
							     echo "<a href='" &KS.Setting(3) & KS.Setting(66) &"/ajax.asp?n=" & n & "&action=verify&topicid=" & TopicID & "&replyid=" &ReplayID &"&boardid=" &boardid&"' onclick=""return(confirm('确定审核该回复吗?'));"">审核</a> | "
								 
								 If TopicNode.SelectSingleNode("@verific").text="1" Then
							     echo "<a href='" &KS.Setting(3) & KS.Setting(66) &"/ajax.asp?action=replylock&topicid=" & TopicID & "&replyid=" & ReplayID & "&boardid=" &boardid&"' onclick=""return(confirm('确定屏蔽该信息吗?'));"">屏蔽</a> | "
							     Else
							     Echo "<a href='" &KS.Setting(3) & KS.Setting(66) &"/ajax.asp?action=replyunlock&topicid=" & TopicID & "&replyid=" & ReplayID & "&boardid=" &boardid&"' onclick=""return(confirm('确定取消屏蔽该信息吗?'));"">解屏</a> | "
							     End If
							     If N=1 Then
							      Echo "<a href='" & KS.Setting(3) & KS.Setting(66) & "/post.asp?action=edit&bid=" & boardid&"&id=" & ReplayID & "&topicid=" & TopicID & "&istopic=1'>编辑主题</a> | <a href=""javascript:void(0)"" onclick=""delsubject("&TopicID &",'" & KS.Setting(66) & "'," &boardid&")"">删除主题</a>"
							     Else
							      echo "<a href='" & KS.Setting(3) & KS.Setting(66) & "/post.asp?action=edit&bid=" & boardid&"&id=" & ReplayID & "&topicid=" & TopicID & "&istopic=0&page=" & CurrentPage & "'>编辑</a> | <a onclick=""delreply('" & KS.Setting(66) &"'," & topicid & "," & replayid & "," & boardid & "," & N & ")"" href='javascript:void(0);'>删除</a>"
							     End If
							  
							  ElseIf KS.ChkClng(BSetting(29))=1 And KSUser.UserName= PostUserName Then
								 If N=1 Then
								  echo "<img src=""" & KS.Setting(3) & KS.Setting(66) & "/images/edit.gif"" align=""absmiddle""/><a href='" & KS.Setting(3) & KS.Setting(66) & "/post.asp?action=edit&bid=" & boardid&"&id=" & ReplayID & "&topicid=" & TopicID & "&istopic=1'>编辑主题</a>"
								  Else
								  echo "<img src=""" & KS.Setting(3) & KS.Setting(66) & "/images/edit.gif"" align=""absmiddle""/><a href='" & KS.Setting(3) & KS.Setting(66) & "/post.asp?action=edit&bid=" & boardid&"&id=" & replayID & "&topicid=" & TopicID & "&istopic=0&page=" & CurrentPage & "'>编辑</a> "
								  End If
							  End If
							  If BSetting(66)<>"1" Then
							  echo " <a href=""#top""><img border=""0"" src=""" & KS.Setting(3) & KS.Setting(66) & "/images/p_up.gif"" alt=""回到顶部""/>顶端</a> <a href=""#reply""><img border=""0"" src=""" & KS.Setting(3) & KS.Setting(66) & "/images/p_down.gif"" alt=""回到底部""/>底部</a> "
							  End If
							 case "showusermanage"
							   If CheckIsMaster And  Not UN  Is Nothing Then
							         If BSetting(66)<>"1" Then
							           echo "<div style=""margin:8px;"">"
									 End If
									 
									  If UN.SelectSingleNode("@lockonclub").text="1" Then
										echo "<a onclick='return(confirm(""确定对该用户解除锁定操作吗？""))' href='" &KS.Setting(3) & KS.Setting(66) &"/ajax.asp?action=unlockuser&userid=" & UN.SelectSingleNode("@userid").text &"'>解除锁定</a>"
									  Else
										echo "<a onclick='return(confirm(""确定锁定该用户吗？""))' href='" &KS.Setting(3) & KS.Setting(66) &"/ajax.asp?action=lockuser&userid=" & UN.SelectSingleNode("@userid").text &"' >锁定该用户</a>"
									  End If
										echo "  <a href=""javascript:void(0)"" onclick=""delusertopic(" & topicid&"," & currentpage & "," & n  &",'"&postusername &"'," & boardid &",'" & KS.Setting(66) & "')"" >删除帖子</a>"
									 If BSetting(66)<>"1" Then
									  echo "</div>"
									 End If
							  End If
							 case "content"
							   'ks.die TopicNode.SelectSingleNode("@content").text
							   Dim Content,UserIsLock,Sign,RightAD,Mstr,MyContent
							   
							   if BSetting(66)<>"1" Then  RightAd=GetAdByRnd(36)
							   If Not KS.IsNul(RightAd) Then echo "<span class=""rightAd"">" & RightAd &"</span>"
							   If Not Un Is Nothing Then UserIsLock=KS.ChkClng(UN.SelectSingleNode("@lockonclub").text) Else UserIsLock=0
							    
								If UserIsLock=1 Then
									if CheckIsMaster=true or KSUser.GetUserInfo("ClubSpecialPower")="1" or KSUser.GetUserInfo("ClubSpecialPower")="2" then
									 MyContent=GetContent(Mstr): Content="<div class=""nopurview"">该用户所发的帖已全被锁定,由于您是版主/管理员所以可以看到此信息.</div>" & Mstr & MyContent
									else
									 Content="<div class=""nopurview"">对不起，该用户所发的帖已全被锁定!</div>"
									end if
								ElseIf TopicNode.SelectSingleNode("@verific").text="2" then
									if CheckIsMaster=true or KSUser.GetUserInfo("ClubSpecialPower")="1" or KSUser.GetUserInfo("ClubSpecialPower")="2" then
									 MyContent=GetContent(Mstr): Content="<div class=""nopurview"">该信息已屏蔽,由于您是版主/管理员所以可以看到此信息.</div>" & Mstr & MyContent
									else
									 Content="<div class=""nopurview"">对不起，该信息已屏蔽!</div>"
									end if
								ElseIf TopicNode.SelectSingleNode("@verific").text="0" then
									if CheckIsMaster=true or KSUser.GetUserInfo("ClubSpecialPower")="1" or KSUser.GetUserInfo("ClubSpecialPower")="2" then
									 MyContent=GetContent(Mstr): Content="<div class=""nopurview"">该信息未审核,由于您是版主/管理员所以可以看到此信息.</div>" &  Mstr & MyContent
									else
									 Content="<div class=""nopurview"">对不起，该信息未审核!</div>"
									end if
								ElseIf N=1 Then  '主题
									  Content=GetContent(Mstr)
									  If IsClose=1 Then Content=Content & "<div class=""closetips"">Tips:本主题已被版主或管理员关闭，可以正常浏览，但不能发表回复！</div>"
									  Session("TopicMustReply")=0
									 If ShowScore>0 or instr(Content,"[replyview]")<>0 Then
									    If ChannelID<>0 And InfoID<>0 Then Content=Content & TopicNode.SelectSingleNode("@content").text
										If Instr(Content,"$@$")=0 Then Content=Content & "$@$"
										Dim ChargeUnit,ChargeField,CArr:CArr=Split(Content,"$@$")
										Dim CUserlist:CUserlist=CArr(1)
										If ShowScore>0 Then '消费积分
										     Select Case KS.ChkClng(BSetting(56))
												  case 0 ChargeUnit=KS.Setting(46) &KS.Setting(45):ChargeField="point"
												  case 1 ChargeUnit="元人民币":ChargeField="money"
												  case 2 ChargeUnit="个积分":ChargeField="score"
											 End Select
										     Dim FreeContent:FreeContent=KS.CutFixContent(Content, "[free]", "[/free]", 1)
											 If FreeContent<>"" Then Carr(0)=Replace(CArr(0),FreeContent,"")
											 If LoginTF=false Then
												Content=FreeContent & "<div class=""nopurview""><img src='" & KS.GetDomain & "user/images/money.gif' border='0'>对不起，您还没有登录，请先登录！本帖售价<span style='color:red'> " & ShowScore & " </span>" & ChargeUnit & "，登录后并支付才可以查看！</div>"
											 ElseIf KSUser.GetUserInfo("ClubSpecialPower")="1" or KSUser.GetUserInfo("ClubSpecialPower")="2" Then
											    Content=FreeContent & "<div class=""nopurview""><img src='" & KS.GetDomain & "user/images/money.gif' border='0'>该主题需要消费 <span style='color:red'>" & ShowScore &"</span> " & ChargeUnit & "才能看,由于您是本论坛管理团队所以可以查看.</div>" & CArr(0)
											 Else
												  If KS.FoundInArr(CUserlist,KSUser.GetUserInfo("UserID"),",")=false And Cint(KSUser.GetUserInfo(ChargeField))<Cint(ShowScore) Then
													Content=FreeContent & "<div class=""nopurview""><img src='" & KS.GetDomain & "user/images/money.gif' border='0'>对不起，本帖售价<span style='color:red'>" & ShowScore & "</span>" & ChargeUnit &",您当前余额<font color=green>" & KSUser.GetUserInfo(ChargeField) &"</font>" & ChargeUnit & "，不足支付！</div>"
												  Else
													  If KS.FoundInArr(CUserlist,KSUser.GetUserInfo("UserID"),",") or CheckIsMaster=true or ksuser.username=TopicNode.SelectSingleNode("@username").text Then
														Content="<div class=""nopurview""><img src='" & KS.GetDomain & "user/images/money.gif' border='0'>该主题售价 <span style='color:red'>" & ShowScore &"</span> " & ChargeUnit & ",您已获得权限，主题内容如下：</div>" & CArr(0)
													  ElseIf KS.S("ShowByScore")="true" Then
														Session("PopTips")="消费" & KS.ChkClng(ShowScore) & ChargeUnit
														Dim PayPoint : PayPoint=(ShowScore*KS.ChkClng(BSetting(58)))/100
														Dim TcMsg:TcMsg="论坛主题“" & Subject & "”的售价分成"
														Select Case KS.ChkClng(BSetting(56))
														 case 0
														  If PayPoint>0 Then Call KS.PointInOrOut(9994,TopicID,TopicNode.SelectSingleNode("@username").text,1,PayPoint,"系统",TcMsg,0)   '支付分成
														  Call KS.PointInOrOut(9994,TopicID,KSUser.UserName,2,ShowScore,"系统","在论坛查看主题[" & Subject & "]消费!",0)
														 case 1
														  If PayPoint>0 Then Call KS.MoneyInOrOut(TopicNode.SelectSingleNode("@username").text,TopicNode.SelectSingleNode("@username").text,PayPoint,4,1,now,0,"系统",TcMsg,9994,TopicID,1)
														  Call KS.MoneyInOrOut(KSUser.UserName,KSUser.UserName,ShowScore,4,2,now,0,"系统","在论坛查看主题[" & Subject & "]消费!",9994,TopicID,1)
														 case 2
														 If KS.ChkClng(PayPoint)>0 Then Call KS.ScoreInOrOut(TopicNode.SelectSingleNode("@username").text,1,KS.ChkClng(PayPoint),"系统",TcMsg,0,0)
														 Session("ScoreHasUse")="+" '设置只累计消费积分
														 Call KS.ScoreInOrOut(KSUser.UserName,2,KS.ChkClng(ShowScore),"系统","在论坛查看主题[" & Subject & "]消费!",9994,TopicID)
														End Select
														Conn.Execute("Update " & posttable & " Set Content='" & Replace(Content,"'","''")& KSUser.GetUserInfo("UserID") & "," &"' Where ID=" &TopicNode.SelectSingleNode("@id").text)
														Content=CArr(0)
													  Else
														Content=FreeContent & "<div class=""nopurview"" style=""text-align:center""><img src='" & KS.GetDomain & "user/images/money.gif' border='0'>本帖售价<font color=red>" & ShowScore & "</font>" & ChargeUnit &",您当前余额<font color=green>" & KSUser.GetUserInfo(ChargeField) &"</font>" & ChargeUnit &",是否确定查看？<br/><input type=""button"" class=""btn"" onClick=""location.href='" & KS.Setting(3) & KS.Setting(66) & "/display.asp?id=" & TopicID &"&showbyscore=true'"" style=""padding:2px 10px"" value="" 确认支付 ""></div>"
													  End If
												 End If
										   End If
										Else   '需要回复
										  Dim replyContent,rept:rept=0 : Session("TopicMustReply")=1
										  If Cbool(LoginTF)=true Then
											if KS.FoundInArr(CUserlist,KSUser.GetUserInfo("UserID"),",") or CheckIsMaster=true or KSUser.GetUserInfo("ClubSpecialPower")="1" or KSUser.GetUserInfo("ClubSpecialPower")="2" or ksuser.username=TopicNode.SelectSingleNode("@username").text then rept=1
											if rept=1 then
												ReplyContent="<div class=""replytips""><font color=""gray"">以下内容只有<b>回复</b>后才可以浏览</font><hr color='#f1f1f1' size='1'>" & KS.CutFixContent(CArr(0), "[replyview]", "[/replyview]", 0) & "</div>"
											  else
												ReplyContent="<div class=""nopurview""><img src='" & KS.Setting(3) & KS.Setting(66) & "/images/locked.gif' align='absmiddle'/><font color=""red"">以下内容只有<b>回复</b>后才可以浏览</font></div>"
											   end if
										  else
										     ReplyContent="<div class=""nopurview""><img src='" & KS.Setting(3) & KS.Setting(66) & "/images/locked.gif' align='absmiddle'/><font color=""red"">以下内容只有<b>回复</b>后才可以浏览,请先登录！</font></div>"
										  End If
										   content=replace(CArr(0),KS.CutFixContent(CArr(0), "[replyview]", "[/replyview]", 1),ReplyContent)
										End If
									  End If
									  
									  
									 
									   
									Content=MStr & Content
									If Instr(Content,"$@$")<>0 then Content=Split(Content,"$@$")(0)
									If KS.ChkClng(PostType)=1 Then  Content=Content & GetVote(TopicID,"")  '投票
									  
									  
								  ElseIf TopicNode.SelectSingleNode("@verific").text="1" Then
									 Content=bbimg(TopicNode.SelectSingleNode("@content").text)
								  Else
								   if CheckIsMaster=true  then
									 Content="<div style=""margin-left:20px; margin-top: 15px; background-color: #ffffee; border: 1px solid #f9c943; width: 300px; height: 30px;line-height:30px;"">该信息未审核,由于您是版主所以可以看到此信息.</div>" & bbimg(KS.HtmlCode(TopicNode.SelectSingleNode("@content").text))
								   ElseIf Not KSUser.GetUserInfo("ClubSpecialPower")="1" or KSUser.GetUserInfo("ClubSpecialPower")="2" Then
									 Content="<div style=""margin-left:20px; margin-top: 15px; background-color: #ffffee; border: 1px solid #f9c943; width: 300px; height: 30px;line-height:30px;"">该信息未审核,由于您是管理员所以可以看到此信息.</div>" & bbimg(KS.HtmlCode(TopicNode.SelectSingleNode("@content").text))
									Else
									Content="<div style=""margin-left:20px; margin-top: 15px; background-color: #ffffee; border: 1px solid #f9c943; width: 300px; height: 50px;line-height:50px; "">本站启用审核机制,该信息未通过审核!</div>"
								   End If
								 end if
							   Content=replace(replace(Content,"｛#","{"),"#｝","}")  '过滤科汛标签
							   Content=Ubbcode(KSR.ReplaceEmot(Content),n)
							  Dim TopAD
							  If BSetting(66)<>"1" Then TopAD=GetAdByRnd(68)  '帖子顶部广告
							  If TopAD<>"" Then
							   Content="<div class=""topad"">" & TopAD &"</div><div class=""clubcontent"" id=""content" & n& """>" & Content & "</div>"
							  Else
							   Content="<div class=""clubcontent"" id=""content" & n& """>" & Content & "</div>"
							  End If
							  Content=KSR.ScanAnnex(Content)
							  echo KS.FilterIllegalChar(Content)
							  
							  echo "<span class=""threadcommnets"" id=""comment_" & replayid&""">"   '点评模块
							  echo GetComments(CommentXML,Boardid,replayid,KS.ChkClng(BSetting(44)),CheckIsMaster)
							  echo "</span>"
						     case "userinfo" 
							   If UN Is Nothing Then
							  	  echo "<div class=""userface""><img src='../Images/Face/boy.jpg' width='82' height='90'></div>"
								  echo "<div style='height:26px;padding-left:5px;margin-top:10px;text-align:left'>用 户 组：游客</div>"
								   PostUserName="游客" : PostUserID=0
							  Else
							  
							   Dim UserFaceSrc:UserFaceSrc=UN.SelectSingleNode("@userface").text
							   PostUserName=UN.SelectSingleNode("@username").text : PostUserID=UN.SelectSingleNode("@userid").text
							   if lcase(left(userfacesrc,4))<>"http" and left(userfacesrc,1)<>"/" then userfacesrc="../" & userfacesrc
							    
								'==================弹出提示开始========================
							   echo "<div class=""bui" & KS.ChkClng(BSetting(66)) & " bui"" id=""user" & n& """ style=""display:none"" onmouseover=""showPopUserInfo(" & n &")"" onmouseout=""hidPopUserInfo(" & n & ")""><div class=""l""><div id='f" & n &"'></div>"
							   echo "<div style='margin-top:5px;padding-left:2px'><img src='" & KS.GetDomain & "images/user/log/106.gif'><a href='javascript:void(0)' onclick=""addF(event,'" & PostUserName & "')"">加为好友</a> <img src='" & KS.GetDomain & "images/user/mail.gif'><a href='javascript:void(0)' onclick=""sendMsg(event,'" & PostUserName & "')"">发送消息</a></div></div>"
							   echo "<div class='r'>"
							   echo "<li class=""line""><a href='" & KS.GetSpaceUrl(PostUserID) & "' target='_blank'><img src='" & KS.GetDomain & KS.Setting(66) & "/images/home.gif' width='16' height='16' border='0' align='absmiddle' alt='TA的空间'></a>空间  |" 
							   IF KS.ChkClng(KS.SSetting(55))=1 Then 
							   echo " <a href='../user/weibo.asp?userid=" & PostUserID & "' target='_blank'><img src='" & KS.GetDomain & KS.Setting(66) & "/images/homepage.gif' border='0' align='absmiddle' alt='TA的微博'></a>微博  |" 
							   End If
								 If UN.SelectSingleNode("@email").text <> "" Then
								echo "  <a href='mailto:" & UN.SelectSingleNode("@email").text & "' target='_blank'><img src='" & KS.GetDomain & KS.Setting(66) & "/images/email.gif' width='18' height='18' border='0' align='absmiddle' alt='电子邮件:[ " & UN.SelectSingleNode("@email").text &" ]'></a>邮件" & vbcrlf
								 Else
							    echo "  <a href='#'><img src='" & KS.GetDomain & KS.Setting(66) & "/images/email-gray.gif' width='18' height='18' border='0' align='absmiddle' alt='电子邮件'></a>邮件" & vbcrlf
								End If
								echo "  |" 
								If UN.SelectSingleNode("@qq").text <> "" and UN.SelectSingleNode("@qq").text <> "0" Then
								echo " <a href='#'><img src='" & KS.GetDomain & KS.Setting(66) & "/images/qq.gif' width='16' height='16' border='0' align='absmiddle' alt='QQ号码:[ " & UN.SelectSingleNode("@qq").text & " ]'></a>QQ号码"
								Else
								echo "  <a href='#'><img src='" & KS.GetDomain & KS.Setting(66) & "/images/qq-gray.gif' width='16' height='16' border='0' align='absmiddle' alt='QQ号码'></a>QQ号码" & vbcrlf
								End If	
								
								echo "</li><li><span>用户:</span>" & PostUserName &"</li><li><span>性别:</span>" & UN.SelectSingleNode("@sex").text &"</li><li><span>积分:</span>" & UN.SelectSingleNode("@score").text & "分</li><li><span>威望:</span>" & UN.SelectSingleNode("@prestige").text &" </li>"
							    echo "<li><span>帖子:</span>" & UN.SelectSingleNode("@postnum").text & "</li><li><span>精华:</span>" & UN.SelectSingleNode("@besttopicnum").text &"</li>"
							    echo "<li><span>广播:</span>" & UN.SelectSingleNode("@msgnum").text & "</li><li><span>粉丝:</span>" & UN.SelectSingleNode("@fansnum").text &"</li>"
							    echo "<li class=""line""><span>登录次数:</span>" & UN.SelectSingleNode("@logintimes").text & " 次</li><li class=""line""><span>注册时间:</span>" & UN.SelectSingleNode("@regdate").text &"</li>"
							    echo "<li class=""line""><span>最后登录:</span>" & UN.SelectSingleNode("@lastlogintime").text & "</li></div></div>"               
								'==================弹出提示结束========================
								
								If BSetting(66)<>"1" Then
								   If UN.SelectSingleNode("@isonline").text="1" Then
									echo "<div class=""username"">" & PostUserName & " <span style='color:#ff6600'>当前在线</span></div>"
								   Else
									echo"<div class=""username"">" & PostUserName & " <span style='color:#888888'>当前离线</span></div>"
								   End If
                                End If
								If BSetting(66)<>"1" Or N=1 Then
									echo "<div onmouseover=""popUserInfo(this," & n & ");""><div class=""userface""><a href='" & KS.GetSpaceUrl(PostUserID) & "' target='_blank'><img style='width:expression(this.width>130?""130px"":this.width+""px"");' onload='if (this.width>130){this.width=130;}' onerror='this.src=""../images/face/boy.jpg""' src='" & UserFaceSrc &"' border='0'/></a></div></div>"
							   echo "<div class=""tns xg2""><table cellspacing=""0"" cellpadding=""0""><th><p><a href='../space/?" & PostUserID &"/club' target='_blank'>" & UN.SelectSingleNode("@postnum").text &"</a></p>主题</th><th><p><a href='../user/weibo.asp?userid=" & PostUserID &"' target='_blank'>" & UN.SelectSingleNode("@msgnum").text &"</a></p>广播</th><td><p><a href='../user/weibo.asp?userid=" & PostUserID &"&f=fans' target='_blank'>" & UN.SelectSingleNode("@fansnum").text &"</a></p>粉丝</td></table></div><div class='clubatten'><a href=""javascript:;"" onclick=""addatt(" & PostUserID &",'false')""><img src='../images/default/addgz.gif' alt='添加关注'/></a></div>"
                                Else
								 	echo "<div onmouseover=""popUserInfo(this," & n & ");""><div class=""userface""><a href='" & KS.GetSpaceUrl(PostUserID) & "' target='_blank'><img  width='90' height='90' onerror='this.src=""../images/face/boy.jpg""' src='" & UserFaceSrc &"' border='0'/></a></div></div>" 
                                   
								   echo "<div style='padding-left:10px;font-weight:bold'>" & PostUserName & " <a href=""javascript:;"" onclick=""addatt(" & PostUserID &",'false')"" style='color:#999;font-weight:normal'>+加关注</a></div>"
								End If
							   

                              If KS.ChkClng(BSetting(66))=0 Then
								   echo "级别:" 
								   If Not KS.IsNul(KS.A_G(UN.SelectSingleNode("@clubgradeid").text ,"color")) Then
								   echo "<span style='color:" & KS.A_G(UN.SelectSingleNode("@clubgradeid").text ,"color") &"'>" & KS.A_G(UN.SelectSingleNode("@clubgradeid").text ,"usertitle") & "</span>"
								   Else
								   echo KS.A_G(UN.SelectSingleNode("@clubgradeid").text ,"usertitle")
								   End If
								   
								   echo "<div style='margin;5px;height:20px;'><img src='" & KS.GetDomain & KS.Setting(66) & "/images/" & KS.A_G(UN.SelectSingleNode("@clubgradeid").text ,"ico") & "'></div>"
								   
								   If KS.Setting(48)="1" and IsBusiness Then
								   If UN.SelectSingleNode("@issfzrz").text="1" then
									echo "<a href='" & KS.GetDomain & "company/rz.asp?userid=" & UN.SelectSingleNode("@userid").text & "' target='_blank'><img src='"  & KS.GetDomain & "Images/default/medal_rz.gif' align='absmiddle' title='已实名认证'/></a><br/>"
								   Else
									echo "<a href='" & KS.GetDomain & "company/rz.asp?userid=" & UN.SelectSingleNode("@userid").text & "' target='_blank'><img src='"  & KS.GetDomain & "Images/default/medal_norz.gif' align='absmiddle' title='未实名认证'/></a><br/>"
								   End If
								   End If
								   
								   
								   echo "用户积分:" & UN.SelectSingleNode("@score").text &" 分<br/>"
								   echo "登录次数:" & UN.SelectSingleNode("@logintimes").text &" 次<br/>"
								   echo "注册时间:" & FormatDateTime(UN.SelectSingleNode("@regdate").text,2) &"<br/>"
								   echo "最后登录:" & FormatDateTime(UN.SelectSingleNode("@lastlogintime").text,2) &"<br/>"
								   Dim MedalArr,Medal:Medal=UN.SelectSingleNode("@medal").text
								   if Not KS.IsNul(Medal) Then  '勋章
									 MedalArr=split(medal,"@@@")
									 For i=0 tO Ubound(MedalArr)
									  echo "<img title='" &split(medalArr(i),"|")(1) & "' src='" & KS.Setting(3) & KS.Setting(66) & "/images/medal/" & split(medalArr(i),"|")(2) &"'>"
									 Next
								   End If
							  End If
							  End If
						  end Select
						case "replay" 
							 select case lcase(sTokenName)
							 case "showupfiles"
							   If KS.ChkClng(BSetting(36))=1 Then
								   If LoginTF=true Then
										If KS.IsNul(BSetting(17)) Or KS.FoundInArr(BSetting(17),KSUser.GroupID,",") Then
										  echo "<tr><td><iframe id=""upiframe"" name=""upiframe"" src=""../user/BatchUploadForm.asp?ChannelID=9994&Boardid=" & boardid & """ frameborder=""0"" width=""100%"" height=""20"" scrolling=""no""></iframe></td></tr>"
										End If
								   End If
								End If 
							case "username" echo ksuser.username
							case "changecategory"
							  If BSetting(23)="1" and BSetting(68)="1" and (CheckIsMaster or ksuser.username=PostUserName) Then
								   echo "<select class=""select"" id=""categoryid"" name=""categoryid"">"
								   echo " <option value=""-1"">--主题分类不变--</option>"
								   
								   KS.LoadClubBoardCategory
								   If IsObject(Application(KS.SiteSN&"_ClubBoardCategory")) Then
								  
								   Dim CNode
								   For Each CNode In Application(KS.SiteSN&"_ClubBoardCategory").DocumentElement.SelectNodes("row[@boardid=" & boardid & "]")
								   if KS.ChkClng(categoryid)=KS.ChkClng(CNode.SelectSingleNode("@categoryid").text) then
								   echo " <option selected value=""" & CNode.SelectSingleNode("@categoryid").text & """>" & CNode.SelectSingleNode("@categoryname").text & "</option>"
								   else
								   echo " <option value=""" & CNode.SelectSingleNode("@categoryid").text & """>" & CNode.SelectSingleNode("@categoryname").text & "</option>"
								   end if
								   Next
								   
								    End If
								   echo "</select> "
							  End If
						
						   
							   
							case "userface"
								 Dim UserFace
								 KSUser.UserLoginChecked
								 If Not KS.IsNUL(KSUser.GetUserInfo("UserFace")) Then
								  UserFace=KSUser.GetUserInfo("UserFace") : If Left(UserFace,1)<>"/" And Left(lcase(UserFace),4)<>"http" Then UserFace=KS.GetDomain & UserFace
								 Else
								  UserFace=KS.GetDomain & "images/face/boy.jpg"
								 End If 
								 echo UserFace
							end select
								  
					End Select
		End Sub
        
		Function GetContent(ByRef MStr)
		 If KS.ChkClng(ChannelID)<>0 And KS.ChkClng(InfoID)<>0 And N=1 Then  '绑定模型
				 Dim MRS,ModelNode,FieldXML,FieldNode
				 Call KSUser.LoadModelField(ChannelID,FieldXML,FieldNode)
				 Set FieldNode=FieldXML.DocumentElement.SelectNodes("fielditem[showonclubform=1][fieldtype!=0&&fieldtype!=13]")
				 If IsObject(FieldNode) Then
				 Set MRS=Conn.Execute("Select top 1 * From " & KS.C_S(ChannelID,2) & " Where PostID=" & TopicID)
				 If Not MRS.Eof Then
				    MStr=MStr & "<table cellspacing=""0"" cellpadding=""0"" class=""modeltable"">"
				    For Each ModelNode In FieldNode
					  MStr=MStr & "<tr><td width=""100"">" & ModelNode.SelectSingleNode("title").text &"：</td><td>" & MRS(trim(ModelNode.SelectSingleNode("@fieldname").text)) 
					    If ModelNode.SelectSingleNode("showunit").text="1" Then
						  MStr=MStr & " " & MRS(trim(ModelNode.SelectSingleNode("@fieldname").text)&"_unit")
						End If
							MStr=MStr & "</td></tr>"
					Next
						 MStr=MStr & "</table>"
						 GetContent=MRS("articlecontent")
					End If
					Set MRS=Nothing
				End If
			 Else
				 GetContent=TopicNode.SelectSingleNode("@content").text
			 End If
		End Function
		
		
		Public Sub Kesion()
		    Startime=timer() 
			If KS.Setting(56)="0" Then KS.Die "本站已关闭" & KS.Setting(61)
			LoginTF=KSUser.UserLoginChecked
			If Not KS.IsNul(KS.Setting(69)) Then
			  Dim QueryStr:QueryStr=Request.QueryString
			  Dim QArr:QArr=Split(Split(QueryStr,".")(0),"-")
			  If Ubound(Qarr)>=1 Then
			   ID=KS.ChkClng(Qarr(1))
			  Else
			   ID=KS.ChkClng(KS.S("ID"))
			  End If
			  If Ubound(QArr)>=2 Then  
			   CurrentPage = KS.ChkClng(Qarr(2))
			  Else
			   CurrentPage = KS.ChkClng(Request("page")) 
			  End If
			Else
		      ID=KS.ChkClng(KS.S("ID"))
			  CurrentPage = KS.ChkClng(Request("page")) 
			End If
			If CurrentPage<=0 Then CurrentPage=1
			If KS.Setting(114)="" Then Response.Write "请先到""基本信息设置->模板绑定""进行模板绑定操作!":response.end
				   FileContent = KSR.LoadTemplate(KS.Setting(160))
				   If KS.IsNul(FileContent) Then FileContent = "模板不存在!"
				   FCls.RefreshType = "guestdisplay" '设置刷新类型，以便取得当前位置导航等
				   FCls.RefreshFolderID = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
				   GetClubPopLogin FileContent
				   Call GetSubject()
				   If BoardID<>0  Then 
				    KS.LoadClubBoard()
				    Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]")
					If Node Is Nothing Then
					 KS.Die "非法参数!"
					End If
					 BSetting=Node.SelectSingleNode("@settings").text
		             master=Node.SelectSingleNode("@master").text
					 'FileContent=RexHtml_IF(FileContent) '先过滤无用的标签,减少标签解释
				   End If
				   
				    BSetting=BSetting & "$$$0$0$0$0$0$0$0$0$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$" :BSetting=Split(BSetting,"$")

					CheckIsMaster=check() '是否版主

					If verific=0 and CheckIsMaster=false Then KS.Die "<script>alert('对不起,该帖子还没有审核！');history.back();</script>"
					If DelTF=1 and CheckIsMaster=false Then KS.Die "<script>alert('对不起，帖子已删除!');location.href='" & KS.GetClubListUrl(boardid) & "';</script>"
					
					FileContent=Replace(FileContent,"{$GetInstallDir}",KS.Setting(3))
					FileContent=Replace(FileContent,"{$GetSiteUrl}",KS.GetDomain)
					FileContent=Replace(FileContent,"{$GetClubInstallDir}",KS.Setting(66))
					
					If BSetting(0)="0" Then  '不允许游客浏览时才进一步判断权限
						If LoginTF=true and PostUserName=KSUser.UserName Then
							 ShowTopic
						Else
						  Dim CheckResult:CheckResult=CheckPermissions(KSUser,BSetting,GuestTitle) '检查访问检查
						  If CheckResult="true" Then
							ShowTopic
						  Else
							ListTemplate=CheckResult : AllowShow=false
						  End If
						End If
					Else
					   ShowTopic
					End If
					
				   FileContent=Replace(FileContent,"{$GuestTitle}",GuestTitle)
                   FileContent=Replace(FileContent,"{$GetGuestList}",ListTemplate)
				   FileContent=KSR.KSLabelReplaceAll(FileContent)
				   FileContent=Replace(Replace(FileContent,"｛#","{"),"#｝","}")  '标签替换回来
				   FileContent=RexHtml_IF(FileContent)
				   FileContent=Replace(FileContent,"{#ExecutTime}","页面执行" & FormatNumber((timer()-startime),5,-1,0,-1) & "秒 powered by CMS")
				   KS.Echo  FileContent
		End Sub
		
		Sub ShowTopic()
		   FileContent=RexHtml_IF(FileContent) '先过滤无用的标签,减少标签解释
			 Dim PostBtnStr:PostBtnStr="<span style=""position:relative;z-index:1000"" onmouseover=""$('#postlist').show()"" onmouseout=""$('#postlist').hide()""><a href=""" & KS.Setting(3) & KS.Setting(66) & "/post.asp?bid=" & boardid & """><img src=""" & KS.Setting(3) & KS.Setting(66) & "/images/button_post.png""></a><div id=""postlist"" class=""submenu noli"">"
			 PostBtnStr=PostBtnStr&"<dl><img src=""" & KS.Setting(3) & KS.Setting(66) & "/images/new_post.gif"" align=""absmiddle""/> <a href=""" & KS.Setting(3) & KS.Setting(66) & "/post.asp?bid=" & boardid & """>发表帖子</a></dl>"
			 If KS.ChkClng(bsetting(64))>0 Then
			 PostBtnStr=PostBtnStr &"<dl><img src=""" & KS.Setting(3) & KS.Setting(66) & "/images/vote.gif"" align=""absmiddle""> <a href=""" & KS.Setting(3) & KS.Setting(66) & "/post.asp?bid=" & BoardID&"&posttype=1"">发起投票</a></dl>"
			 End If
			 PostBtnStr=PostBtnStr &"</div></span>"
	
					   FileContent=Replace(FileContent,"{$PostButtonAction}",PostBtnStr)
					   FileContent=Replace(FileContent,"{$BoardID}",BoardID)
					   FileContent=Replace(FileContent,"{$TopicID}",TopicID)
					   FileContent=Replace(FileContent,"{$PostTable}",PostTable)
					   FileContent=Replace(FileContent,"{$IsTop}",IsTop)
					   FileContent=Replace(FileContent,"{$Page}",currentpage)
					   AllowShow=true
					   FileContent=Replace(FileContent,"{$GuestTitle}","{@topic.subjectnohtml}")
					   FileContent=KSR.KSLabelReplaceAll(FileContent)
                       GetReplayList:If IsObject(Xml) Then Call GetTopicList(XML)
					   if instr(FileContent,"{#GetClubPopLogin}")<>0 Then GetClubPopLogin FileContent
					   SCan FileContent
					   If Session("PopTips")<>"" Then  Response.write "<script>$(document).ready(function(){popShowMessage('" &Session("PopTips") & "');});</script>": Session("PopTips")=""
					
			 KS.Die ""
		End Sub
		
		Sub GetSubject()
		  Dim Param
		  If Request("Move")<>"" Then
		    If Request("Move")="next" Then Param=" Where BoardID=" & KS.ChkClng(KS.S("BoardID")) & " and ID>" & ID & " Order By ID" Else Param=" Where  BoardID=" & KS.ChkClng(KS.S("BoardID")) & " and ID<" & ID & " Order By ID desc"
		  Else
		    Param=" Where ID=" & ID
		  End If
		  Set RST=Conn.Execute("Select top 1 ID,Verific,IsBest,IsTop,CategoryID,Subject,UserName,Hits,PostTable,PostType,ShowScore,TotalReplay,BoardID,DelTF,ChannelID,InfoID,IsClose From KS_GuestBook" & Param)
		  If RST.Eof Then
		   RST.Close:Set RST=Nothing
		   If Request("Move")<>"" Then
		    KS.Die("<script>alert('已没有记录了！');history.back();</script>")
		   Else
		    KS.Die("<script>alert('非法参数！');window.close();</script>")
		   End If
		  End If
		  ID       = RST("ID") : TopicID=ID : ChannelID=RST("ChannelID") : InfoID=RST("InfoID")
		  verific  = RST("Verific"):IsBest = KS.ChkClng(RST("IsBest")):IsTop = KS.ChkClng(RST("IsTop")) : IsClose=KS.ChkClng(RST("IsClose")) : CategoryID=KS.ChkClng(RST("CategoryID")):DelTF = KS.ChkClng(RST("DelTf")):PostUserName=RST("UserName")
		  Subject  = KS.FilterIllegalChar(RST("Subject")) : Subject  = replace(replace(subject,"{","｛#"),"}","#｝") '过滤科汛标签
		  GCls.Execute("Update KS_GuestBook Set Hits=Hits+1 Where ID=" & ID)
		  Hits     = rst("Hits"): PostTable = RST("PostTable") : PostType=RST("PostType")
		  ShowScore = KS.ChkClng(RST("ShowScore"))
		  TreplayNum= KS.ChkClng(RST("TotalReplay"))
		  TotalReplay=TreplayNum+1
		  FCls.RefreshFolderID = RST("BoardID")
		  BoardID=FCls.RefreshFolderID
		  RST.Close : Set RST=Nothing
		  If IsTop<>0 Then
		    If Not IsObject(Application(KS.SiteSN &"TopXML")) Then MustReLoadTopTopic
			Application(KS.SiteSN &"TopXML").DocumentElement.SelectSingleNode("row[@id=" & id&"]/@hits").text=hits
		  End If
		End Sub
		
		Sub GetReplayList()	
		 MaxPerPage=KS.ChkClng(BSetting(21)) : If MaxPerPage=0 Then MaxPerPage=10
		 Dim Param:Param=" DelTF=0 and topicid=" & ID
		 If Request.QueryString("UserName")<>"" Then Param=Param & " And UserName='" & KS.R(KS.S("UserName")) & "'"
		 If DataBaseType=1 Then
			    Dim Cmd : Set Cmd = Server.CreateObject("ADODB.Command")
				Set Cmd.ActiveConnection=conn
				Cmd.CommandText="KS_ClubsDisplay"
				Cmd.CommandType=4	
				CMD.Prepared = true 	
				Cmd.Parameters.Append cmd.CreateParameter("@rootid",3)
				Cmd.Parameters.Append cmd.CreateParameter("@pagenow",3)
				Cmd.Parameters.Append cmd.CreateParameter("@pagesize",3)
				Cmd.Parameters.Append cmd.CreateParameter("@totalusetable",200,1,20)
				Cmd.Parameters.Append cmd.CreateParameter("@param",200,1,110)
				'Cmd.Parameters.Append cmd.CreateParameter("@totalput",3,2,4)
				Cmd("@rootid")= ID
				Cmd("@pagenow")=CurrentPage
				Cmd("@pagesize")=MaxPerPage
				Cmd("@totalusetable")=PostTable
				If Not KS.IsNUL(Request.QueryString("UserName")) Then
				 Cmd("@param")=" and DelTF=0 and username='"+KS.S("UserName")+"'"
				Else
				 Cmd("@param")=" and DelTF=0"
				End If
				Set Rs=Cmd.Execute
				'rs.close  '注意：若要取得参数值，需先关闭记录集对象
				'TotalPut= cmd("@totalput")
				 TotalPut=GCls.Execute("Select Count(1) From " & PostTable& " Where " & Param)(0)
				'rs.open
				If Not RS.Eof Then 
				   Set XML=KS.RsToXml(RS,"row","")
				Else
					KS.AlertHintScript "没有记录了!"
				End If
				Rs.close()
				Set Rs=Nothing
				Set Cmd =  Nothing
			   Exit Sub
		Else
			 If TotalReplay=0 Then TotalReplay=1
			 SQLStr=KS.GetPageSQL(PostTable,"id",MaxPerPage,CurrentPage,0,Param,"*")
			 Dim RS:Set RS=conn.Execute(SQLStr)
			 IF RS.Eof And RS.Bof Then 
				  RS.Close:Set RS=Nothing: totalput=0: exit sub
			 Else
					TotalPut= GCls.Execute("Select Count(1) From " & PostTable& " Where " & Param)(0)
					Set XML=KS.RsToXml(RS,"row","")
					RS.Close:Set RS=Nothing
			End IF
		End If
		
	End Sub
		
	Sub GetTopicList(Xml)
		     If CurrentPage=1 Then N=0 Else N=MaxPerPage*(CurrentPage-1)
			 For Each Node In Xml.DocumentElement.SelectNodes("row")
			    If UserNames="" Then
				 UserNames="" & trim(Node.SelectSingleNode("@userid").text) & ""
				ElseIF KS.FoundInArr(UserNames,"" & Node.SelectSingleNode("@userid").text & "",",")=false Then
				 UserNames=UserNames & "," & trim(Node.SelectSingleNode("@userid").text) & ""
				End If
				If Pids="" Then
					Pids=Node.SelectSingleNode("@id").text
				Else
				    Pids=Pids & "," & Node.SelectSingleNode("@id").text
				End If
			 Next
			 UserNames=KS.FilterIds(UserNames) : If UserNames="" Then UserNames=0
			 If DataBaseType=1 Then
			    Dim Cmd : Set Cmd = Server.CreateObject("ADODB.Command")
				Set Cmd.ActiveConnection=conn
				Cmd.CommandText="KS_ClubsUserList"
				Cmd.CommandType=4	
				CMD.Prepared = true 	
				Cmd.Parameters.Append cmd.CreateParameter("@num",3)
				Cmd.Parameters.Append cmd.CreateParameter("@UserNames",202,1,8000)
				Cmd.Parameters.Append cmd.CreateParameter("@UserFields",202,1,300)
				Cmd("@num")=MaxPerPage
				Cmd("@UserNames")= UserNames 
				Cmd("@UserFields")=UserFields
				Set Rs=Cmd.Execute
				Set Cmd=Nothing
			 Else
				Set RS=GCls.Execute("Select top " & MaxPerPage & " " & UserFields &" From KS_User Where UserID in(" & UserNames & ")")
			 End If
			 If Not RS.Eof Then Set UserXml=KS.RsToXml(RS,"row","")
			 RS.Close:Set RS=Nothing
			
			 If Pids<>"" And KS.ChkClng(BSetting(44))<>0 Then
				Set RS=GCls.Execute("Select * From KS_GuestComment Where tid=" & id & " and pid in(" & pids & ") order by orderid,id desc")
				If Not RS.Eof Then
				 Set CommentXML=KS.RsToXml(rs,"row","")
				End If
				RS.Close :Set RS=Nothing
			 End If
	End Sub
		
	Function GetFloor(n)
			  select case n
			   case 1 GetFloor="楼主"
			   case 2 GetFloor="沙发"
			   case 3 GetFloor="藤椅"
			   case 4 GetFloor="板凳"
			   case 5 GetFloor="报纸"
			   case 6 GetFloor="地板"
			   case else
			   GetFloor=n & "楼"
			  end select
	 End function
	 
	 Private Function bbimg(strText)
		Dim s,re
        Set re=new RegExp
        re.IgnoreCase =true
        re.Global=True
		s=strText
		re.Pattern="<img(.[^>]*)([/| ])>"
		s=re.replace(s,"<img$1/>")
		re.Pattern="<img(.[^>]*)/>"
		s=re.replace(s,"<img$1 onclick=""window.open(this.src)"" style='max-width:600px;width:600px;width:expression(document.body.clientWidth>600 ?""600px"":""auto"");overflow:hidden;'/>")
		bbimg=s
	End Function
	
	
	
%>
 <!--#include file="../ks_cls/ubbfunction.asp"-->
<%		
	 function check()
	 	Dim KSLoginCls
		Set KSLoginCls = New LoginCheckCls1
		If KSLoginCls.Check=true Then
		  check=true
		  Exit function
		else
			Dim KSUser:Set KSUser=New UserCls
			LoginTF=KSUser.UserLoginChecked
			If Cbool(LoginTF)=false Then 
			  check=false
			  exit function
			elseif KSUser.GetUserInfo("ClubSpecialPower")="2" Or KSUser.GetUserInfo("ClubSpecialPower")="1" Then
			  check=true
			  exit function
			else
			   check=KS.FoundInArr(master, KSUser.UserName, ",")
			End If
		end if
	 End function	
	 
	 '随机获取广告,AdType广告类型  36 右侧广告,37 底部广告
	 Function GetAdByRnd(ByVal AdType)
	      Dim AdStr:AdStr=KS.Setting(AdType)
	      If KS.IsNul(AdStr) Then Exit Function
		  Dim AdArr:AdArr=Split(AdStr,"@")
		  Dim RandNum,N: N=Ubound(AdArr)+1
          Randomize
          RandNum=Int(Rnd()*N)
          GetAdByRnd=AdArr(RandNum)
	End Function
		
	Function ReplaceBadWord(str)
	  ReplaceBadWord=str
	End Function
					  
End Class
%>

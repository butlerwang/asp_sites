<!--#include file="../KS_Cls/Kesion.SpaceCls.asp"-->
<%

Class SpaceApp
        Public  Domain,FoundSpace,Param,PreviewTemplateID
        Private KS,UserName,UserID,KSR,Action,ID,Node,CurrPage,TotalPut,MaxPerPage,PageNum
		Private Template,TemplateSub,SubStr,BlogName,KSBCls
		Private Sub Class_Initialize()
		  MaxPerPage=10 : PageNum=1 : PreviewTemplateID=0
		  Set KS=New PublicCls
		  Set KSR = New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSR=Nothing
		 Call CloseConn()
		End Sub
		Public Sub Kesion()
		    If KS.SSetting(0)="0" Then FoundSpace=false : EXIT Sub
		    Dim QueryStrings:QueryStrings=Request.ServerVariables("QUERY_STRING")
			QueryStrings=KS.UrlDecode(QueryStrings)
			Call Show(QueryStrings)
		End Sub
		
		Sub Show(ByVal QueryStrings)
		if instr(QueryStrings,"&")<>0  then PreviewTemplateID=Split(QueryStrings,"&")(1) : QueryStrings=Split(QueryStrings,"&")(0)
		Dim QSArr:QSArr=Split(QueryStrings,"/")
		If Ubound(QSArr)>=0 Then
		 UserName=KS.DelSQL(QSArr(0))
		 If KS.ChkClng(UserName)=0 Then
		  Param=" Where UserName='" & UserName & "'"
		 Else
		  Param=" Where UserID=" & KS.ChkClng(UserName) & " or username='" & UserName& "'"
		 End If
		Else
		  Param=" Where [domain]='" & domain & "'"
		End If

		If Ubound(QSArr)>=1 Then Action=QSArr(1)
		If Ubound(QSArr)>=2 Then ID=KS.ChkClng(QSArr(2))
		If Ubound(QSArr)>=3 Then CurrPage=KS.ChkClng(QSArr(3))
		If CurrPage<=0 Then CurrPage=1
		
		Set KSBCls=New BlogCls
		Dim RS
		If DataBaseType=1 Then
			    Dim Cmd : Set Cmd = Server.CreateObject("ADODB.Command")
				Set Cmd.ActiveConnection=conn
				Cmd.CommandText="KS_ShowSpaces"
				Cmd.CommandType=4	
				CMD.Prepared = true 	
				Cmd.Parameters.Append cmd.CreateParameter("@Param",200,1,220)
				Cmd("@param")=param
				Set Rs=Cmd.Execute
				Set Cmd=Nothing
		Else
			Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select top 1 * From KS_Blog" &Param,conn,1,1
		End If
		If RS.Eof And RS.Bof Then
		 rs.close:set rs=nothing : FoundSpace=false
		 Exit Sub
		End If
		FoundSpace=true
		UserName=RS("UserName")
		Session("SpaceUserName")=UserName   '用于sql标签调用
		UserID=RS("UserID")
		Session("SpaceUserID")=UserID   '用于系统函数标签调用
		Domain=RS("Domain")
		If KS.C("AdminName") = "" Or KS.C("AdminPass") = "" Then
			If RS("Status")=0 Then
			 rs.close:set rs=nothing
			 KS.Die "<script>alert('该空间站点尚未审核!');window.close();</script>"
			elseif RS("Status")=2 then
			 rs.close:set rs=nothing
			 KS.Die "<script>alert('该空间站点已被管理员锁定!');window.close();</script>"
			end if
		End If
		If KS.FoundInArr(KS.U_G(Conn.Execute("Select top 1 GroupID From KS_User Where UserName='" & UserName & "'")(0),"powerlist"),"s01",",")=false Then 
		 RS.Close : Set RS=Nothing
		 if not conn.execute("select top 1 username from ks_enterprise where username='" & username & "'").eof then
		  KS.Die "<script>location.href='../company/show.asp?username=" & username & "';</script>"
		 else
			 KS.Die "<script>alert('对不起，该用户没有开通空间的权限!');window.close();</script>"
		 end if
		End If
		
		'============================记录访问次数及最近访客============================================
		conn.execute("update KS_Blog Set Hits=Hits+1 Where UserName='" & UserName & "'")
		If KS.C("UserName")<>"" And KS.C("UserName")<>UserName Then
		   Dim RSV:Set RSV=Server.CreateObject("adodb.recordset")
		   RSV.Open "Select top 1 * From KS_BlogVisitor Where UserName='" & UserName & "' and Visitors='" & KS.R(KS.C("UserName")) & "'",conn,1,3
		   If RSV.Eof And RSV.Bof Then
		     RSV.AddNew
			 RSV("UserName")=UserName
			 RSV("Visitors")=KS.C("UserName")
		   End If
		    RSV("AddDate")=Now
			RSV.Update
		    RSV.Close : Set RSV=Nothing
		 End If
		'============================结束记录============================================================
		 
		 Dim Xml:Set XML=KS.RsToXml(rs,"row","")
		 If Not IsObject(xml) Then KS.Die "error xml!"
		 Set Node=XML.DocumentElement.SelectSingleNode("row")
		 Set KSBCls.Node=Node
		 KSBCls.UserName=UserName
		 KSBCls.UserID=UserID
		 KSBCls.Domain=Domain
		 KSBCls.PreviewTemplateID=KS.ChkClng(PreviewTemplateID)
		 RS.Close : Set RS=Nothing
		 Dim TemplateID:TemplateID=Node.SelectSingleNode("@templateid").text
		 If KS.ChkClng(PreviewTemplateID)<>0 Then TemplateID=PreviewTemplateID
		 If Action<>"" Then template=Template & KSBCls.GetTemplatePath(TemplateID,"TemplateSub")
		 select case Lcase(action)
		   case "blog"
		      KSBCls.Title="博客"
			  BlogList
		   case "log" Call BlogLog
		   case "club"
		      KSBCls.Title="我的话题"
			   ClubList
		   case "album" 		    
		     KSBCls.Title="相册"
			 AlbumList
		   case "showalbum" Call ShowAlbum
		   case "group"
		     KSBCls.Title="圈子"
			 GroupList
		   case "friend"
		     KSBCls.Title="好友"
			 FriendList
		   case "xx"
		     KSBCls.Title="文集"
			 Call xxList
		   case "info"
		     KSBCls.Title="资料"
			 substr=substr & KSBcls.Location("首页 >> 个人档")
			 SubStr=SubStr & KSBCls.UserInfo(Template)
		   case "message"
		     KSBCls.Title="留言"
		     Call ShowMessage
		   case "intro"
		     KSBCls.Title="公司介绍"
			 SubStr=KSBcls.Location("首页 >> 公司简介")
			 Dim Irs:Set Irs=Conn.Execute("Select top 1 Intro From KS_EnterPrise Where UserName='" & UserName & "'")
			 if Not Irs.Eof Then
			 SubStr=SubStr & KS.HtmlCode(Irs(0))
		     Else
		       Irs.Close: Set Irs=Nothing
		       KS.AlertHintScript "对不起，该用户不是企业用户！"
			 End If
			 Irs.Close:Set IrS=Nothing
		   case "news" KSBCls.Title="公司动态" : GetNews
		   case "shownews" ShowNews
		   case "product"  ProductList
		   case "showproduct" ShowProduct
		   case "ryzs" KSBCls.Title="荣誉证书" : GetRyzs
		   case "job" JobList
		   case "showphoto" ShowPhoto
		   case else
		    KSBCls.Title="首页"
		    template=KSBCls.GetTemplatePath(TemplateID,"TemplateMain")
		 end select
		  template=Replace(Template,"{$BlogMain}",replace(SubStr,"{","｛#"))
		  template=KSBCls.ReplaceBlogLabel(Template)
		  KS.Echo KSBCls.LoadSpaceHead
		  KS.Echo Replace(Template,"｛#","{")
		  
		End Sub
		%>
		<!--#Include file="../ks_cls/ubbfunction.asp"-->
		<%
		'日志
		Sub BlogLog()
		  If ID=0 Then KS.Die "error logid!"
		  Dim RS,i
		  Set RS=Server.CreateObject("ADODB.RECORDSET")
		  If KS.C("AdminName") = "" Or KS.C("AdminPass") = "" Then
		   RS.Open "Select top 1 * from KS_BlogInfo Where ID=" & ID & " and Status=0",conn,1,1
		  Else
		   RS.Open "Select top 1 * from KS_BlogInfo Where ID=" & ID,conn,1,1
		  End If
		  If RS.EOF And RS.BOF Then
			KS.Die "<script>location.href='" & KS.GetDomain & "space/?" &userid & "/blog';</script>"
		  End If
		  if rs("istalk")=1 then
		  KSBCls.Title="查看" & UserName & "发表的新鲜事"
		  else
		  KSBCls.Title=rs("title")
		  end if
		  
		  substr=substr & KSBcls.Location("<span style=""float:right""><a href=""" & KS.GetDomain & "user/User_Blog.asp?Action=Add""><img src='" & KS.GetDomain & "user/images/icon7.png' border='0'/>写博文</a></span>首页 >> 查看" & UserName & "发表的博文 ")


		  conn.execute("update KS_BlogInfo Set Hits=Hits+1 Where ID=" & ID)
		  iF rs("IsTalk")="1" Then
		      SubStr=SubStr & "<strong><a href='../space/?" & userid & "' target='_blank'>" & UserName & "</a>说：</strong>" & KS.ReplaceInnerLink(Ubbcode(RS("Content"),i)) & "<BR/><BR/>"
		  Else
			  SubStr=substr & LFCls.GetConfigFromXML("space","/labeltemplate/label","log")
			   Dim EmotSrc:If RS("Face")<>"0" Then EmotSrc="<img src=""../User/images/face/" & RS("Face") & ".gif"" border=""0"">"
			   Dim Tags:Tags=RS("Tags")
			   If Not KS.IsNul(Tags) Then
			   Dim TagList,TagsArr:TagsArr=Split(Tags," ")
					if RS("Tags")<>"" then
					TagList="<div style='display:none'><form id='mytagform' target='_blank' action='../space/?" & username & "/blog' method='post'><input type='text' name='tag' id='tag'></form></div><div style='text-align:left'><strong>标签：</strong>"
					 For I=0 To Ubound(TagsArr)
					  If TagsArr(i)<>"" then
						TagList=TagList &"<a href=""javascript:void(0)"" onclick=""$('#tag').val('" & TagsArr(i) & "');$('#mytagform').submit();"">" & TagsArr(i) & "</a> "
					  end if
					 Next
					 TagList=TagList &"&nbsp;&nbsp;&nbsp;&nbsp;"
					end if
			   End If
				Dim MoreStr
				If Lcase(KS.C("UserName"))=Lcase(UserName) Then
				 MoreStr="<a href=""" & KS.GetDomain & "user/User_Blog.asp?action=Edit&id=" & id & """>编辑</a> | <a href=""" & KS.GetDomain & "user/user_blog.asp?action=Del&id=" & id &""" onclick=""return(confirm('确定删除博文吗？'))"">删除</a> | "
				End If
				MoreStr=MoreStr & "阅读次数("&RS("hits")&") | 回复数("& Conn.Execute("Select Count(ID) From KS_BlogComment Where LogID="  &id)(0) &")</div>"
				Dim ContentStr
				
				Dim JFStr:If RS("Best")="1" then JFStr="  <img src=""../images/jh.gif"" align=""absmiddle"">" else JFStr=""
	
				If KS.IsNul(RS("PassWord")) Then 
				
				 ContentStr=RS("Content")
				ElseIf KS.S("Pass")<>"" Then
				  If KS.S("Pass")=rs("password") then
				   ContentStr=RS("Content")
				  Else
				   SubStr="<br /><br />出错啦,您输入的日志密码有误!<a href='javascript:history.back(-1)'>返回</a><br/>"
				   exit sub
				  End if
				Else
				 SubStr="<br/><br/><br/><form method='post' action='" & KSBCls.GetLogUrl(RS) & "'>本篇文章已被主人加密码,请输入日志的查看密码：<input style='border-style:1 px solid;height:25px;line-height:24px ' class='textbox' type='password' name='pass' size='15'>&nbsp;<input type='submit' value=' 查看 '></form>"
				  exit sub
				End IF
			   SubStr=Replace(SubStr,"{$ShowLogTopic}",EmotSrc & RS("Title") & jfstr)
			   SubStr=Replace(SubStr,"{$ShowLogInfo}","[" & RS("AddDate") & "|by:" & RS("UserName") & "]")
			   SubStr=Replace(SubStr,"{$ShowLogText}",KS.ReplaceInnerLink(Ubbcode(ContentStr,i)))
			   SubStr=Replace(SubStr,"{$ShowLogMore}", TagList&MoreStr)
			   
			   SubStr=Replace(SubStr,"{$ShowTopic}",RS("Title"))
			   SubStr=Replace(SubStr,"{$ShowAuthor}",RS("UserName"))
			   SubStr=Replace(SubStr,"{$ShowAddDate}",RS("AddDate"))
			   SubStr=Replace(SubStr,"{$ShowEmot}",EmotSrc)
			   SubStr=Replace(SubStr,"{$ShowWeather}",KSBCls.GetWeather(RS))
			   SubStr=KSR.ReplaceEmot(SubStr)
			   SubStr=SubStr & "<div style=""padding-left:20px;text-align:left"">上一篇:" & ReplacePrevNextArticle(ID,"Prev")
			   SubStr=SubStr & "<br>下一篇:" & ReplacePrevNextArticle(ID,"Next") & "</div><br>"
		   End If
		   Dim Title:Title=RS("Title")

		   RS.Close:Set RS=Nothing
		
	maxperpage=5
	 Dim sqlstr:SqlStr="Select * From KS_BlogComment Where LogID=" & ID & " Order By AddDate Desc"
	 Set RS=Server.CreateObject("ADODB.RECORDSET")
	 RS.Open SqlStr,Conn,1,1
     IF Not Rs.Eof Then
		    totalPut = RS.RecordCount
		    If CurrPage > 1 and (CurrPage - 1) * MaxPerPage < totalPut Then
				RS.Move (CurrPage - 1) * MaxPerPage
			End If
			Call showContent(rs)
	End If
   If totalput>maxperpage Then
   substr=substr & showpage
   End If
   substr=substr &"<div class=""clear""></div>"
   SubStr=SubStr & "<script src=""writecomment.asp?UserID=" & UserID &"&ID=" & ID & "&UserName=" & UserName & "&Title=" & Title & """></script>"
  
  rs.close:set rs=nothing
End Sub

Sub ShowContent(rs)
     substr=substr & "<div style=""border-bottom:1px solid #f1f1f1;padding-bottom:2px;font-weight:bold;font-size:14px;text-align:left"">&nbsp;&nbsp;本文有 <font color=red>" & totalPut & " </font> 条评论，共分 <font color=red>" & pagenum & "</font> 页,第 <font color=red>" & CurrPage & "</font> 页</div>"
    substr=substr & "<table  width='99%' border='0' align='center' cellpadding='0' cellspacing='1'>"
  Dim FaceStr,Publish,i,n
     If CurrPage=1 Then
	  N=TotalPut
	 Else
	  N=totalPut-MaxPerPage*(CurrPage-1)
	 End IF
  Do While Not RS.Eof 
   FaceStr=KS.Setting(3) & "images/face/boy.jpg"

    Publish=RS("AnounName")
	If not Conn.Execute("Select top 1 UserFace From KS_User Where UserName='"& Publish & "'").eof Then
      FaceStr=Conn.Execute("Select top 1 UserFace From KS_User Where UserName='"& Publish & "'")(0)
	  If lcase(left(FaceStr,4))<>"http" and left(facestr,1)<>"/" then FaceStr=KS.Setting(3) & FaceStr
   End IF
	
   substr=substr & "<tr>"
   substr=substr & "<td width='70' rowspan='2' style='margin-top:3px;BORDER-BOTTOM: #999999 1px dotted;' valign='top'><img width=""50"" height=""52"" src=""" & facestr & """ border=""1"" class=""faceborder"" style=""margin-top:2px;margin-bottom:5px""></td>"
  ' substr=substr & "<td height='25' width=""70%"">"
  ' substr=substr & RS("Title")
  ' substr=substr  & "  </td><td width=""30"" align=""right""><font style='font-size:32px;font-family:""Arial Black"";color:#EEF0EE'> " & N & "</font></td>"
   'substr=substr & "</tr>"
   'substr=substr & "<tr>"
   substr=substr & "<td height='25'>" & ReplaceFace(RS("Content"))
   		 If Not KS.IsNul(RS("Replay")) Then
		 substr=substr&"<div style=""margin : 5px 20px; border : 1px solid #efefef; padding : 5px;background : #ffffee; line-height : normal;""><b>以下为space主人的回复:</b><br>" & RS("Replay") & "<br><div align=right>时间:" & rs("replaydate") &"</div></div>"
         End If
   substr=substr & "	 </td>"
   substr=substr & "</tr>"
   substr=substr & "<tr>"
   
   			 Dim MoreStr,KSUser,LoginTF
				 IF trim(KS.C("UserName"))=trim(RS("UserName")) Then
                 MoreStr="<a href='" & RS("HomePage") & "' target='_blank'>主页</a>| <a href='#'>顶部</a> | <a href='../User/user_message.asp?Action=CommentDel&id=" & RS("ID") & "' onclick=""return(confirm('确定删除该留言吗?'));"">删除</a> | <a href='../user/user_message.asp?id=" & RS("ID") & "&Action=ReplayComment' target='_blank'>回复</a>"
			 Else
                 MoreStr="<a href='" & RS("HomePage") & "' target='_blank'>主页</a>| <a href='#'>顶部</a> "
			 End If

   substr=substr & "<td align='right' colspan='2' style='margin-top:3px;BORDER-BOTTOM: #999999 1px dotted;'><font color='#999999'>(" & publish & " 发表于：" & RS("AddDate") &")</font>&nbsp;&nbsp;" & MoreStr & " </td>"
   substr=substr & "</tr>"
   N=N-1
   RS.MoveNext
		I = I + 1
	  If I >= MaxPerPage Then Exit Do
  loop
 substr=substr & "</table>"

End Sub

Function ReplaceFace(c)
		 Dim str:str="惊讶|撇嘴|色|发呆|得意|流泪|害羞|闭嘴|睡|大哭|尴尬|发怒|调皮|呲牙|微笑|难过|酷|非典|抓狂|吐|"
		 Dim strArr:strArr=Split(str,"|")
		 Dim K
		 For K=0 To 19
		  c=replace(c,"[e"&K &"]","<img title=""" & strarr(k) & """ src=""" & KS.Setting(3) & "images/emot/" & K & ".gif"">")
		 Next
		 ReplaceFace=C
End Function
		
		
		Function ReplacePrevNextArticle(NowID,TypeStr)
		    Dim SqlStr
			If Trim(TypeStr) = "Prev" Then
				   SqlStr = " SELECT Top 1 ID,Title From KS_BlogInfo Where istalk<>1 and UserName='" & UserName & "' And ID<" & NowID & " And Status=0 Order By ID Desc"
			ElseIf Trim(TypeStr) = "Next" Then
				   SqlStr = " SELECT Top 1 ID,Title From KS_BlogInfo Where istalk<>1 and UserName='" & UserName & "' And ID>" & NowID & " And Status=0 Order By ID Desc"
			Else
				ReplacePrevNextArticle = "":Exit Function
			End If
			 Dim RS:Set RS=Server.CreateObject("ADODB.Recordset")
			 RS.Open SqlStr, Conn, 1, 1
			 If RS.EOF And RS.BOF Then
				ReplacePrevNextArticle = "没有了"
			 Else
			  ReplacePrevNextArticle = "<a href=""" & KSBCls.GetCurrLogUrl(UserID,RS("ID")) & """ title=""" & RS("Title") & """>" & RS("title") & "</a>"
			 End If
			 RS.Close:Set RS = Nothing
	 End Function
	 
	 '我的话题
	 Sub ClubList()
	     MaxPerPage=20
	     substr=substr & KSBcls.Location("首页 >> 我的论坛话题")
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select id,subject,TotalReplay,BoardID,AddTime,LastReplayTime,LastReplayUser From KS_GuestBook Where deltf=0 and UserName='" & UserName & "' and verific=1 order by ID Desc",conn,1,1
		 If RS.Eof And RS.Bof Then
		  SubStr=SubStr & UserName & "还没有发表任何话题！" 
		 Else
				totalPut = RS.RecordCount
				If CurrPage > 1 and (CurrPage - 1) * MaxPerPage < totalPut Then
						RS.Move (CurrPage - 1) * MaxPerPage
				End If
				KS.LoadClubBoard
				Dim I:I=0
				SubStr=SubStr & "<table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""border""><tr height=""28"" class=""titlename"">"
				SubStr=SubStr & "<td align=""center"">主题</td><td align=""center"" nowrap>版块</td>"
				SubStr=SubStr & "<td align=""center"" width=""60"" nowrap>回复</td><td align=""center"" nowrap>最后发表</td></tr>"
				Do While Not RS.Eof
				   SubStr=SubStr & "<tr><td class='splittd'><img src='../images/default/arrow_r.gif' align='absmiddle' /> <a href='" &KS.GetClubShowUrl(rs("id")) & "' target='_blank'>" & replace(replace(rs("subject"),"{","｛"),"}","｝") & "</a><br/><span class='tips'>发表时间：" & rs("addTime") & "</span></td>"
				   Dim Node:Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & rs("boardid") &"]")
					if not node is nothing then
						SubStr=SubStr & "<td class='splittd' style='text-align:center'><a href='" & KS.GetClubListUrl(rs("boardid")) &"' target='_blank'>" & Node.SelectSingleNode("@boardname").text & "</a></td>"
					else
						 SubStr=SubStr & "<td class='splittd' style='text-align:center'>---</td>"
					end if
				    Set Node=Nothing
					SubStr=SubStr &"<td style='text-align:center' class='splittd'>" & rs("TotalReplay") & "</td>"
					SubStr=SubStr &"<td style='text-align:center' class='splittd'><a href='../space/?" & RS("LastReplayUser") & "' target='_blank'>" & RS("LastReplayUser") & "</a><br/><span class='tips'>" & rs("LastReplayTime") & "</span></td>"
						   
				   SubStr=SubStr &"</tr>" 
				   I=i+1
				   If I>=MaxPerPage Then Exit Do
				RS.MoveNext
				Loop
				SubStr=SubStr & "</table>"
			End If
		 RS.Close:Set RS=Nothing
		 SubStr=SubStr & vbcrlf & ShowPage
	 End Sub
	 
	
	 
	 '博文列表
	 Sub BlogList()
	     Dim Loc
		 If KS.C("UserName")=UserName Then Loc="<span style=""float:right""><a href=""" & KS.GetDomain & "user/User_Blog.asp?Action=Add""><img src='" & KS.GetDomain & "user/images/icon7.png' border='0'/>写博文</a></span>"
		 Loc=Loc & "首页 >> 博文"
	     substr=substr & KSBcls.Location(loc)
		 MaxPerPage =KSBCls.GetUserBlogParam(UserName,"ListBlogNum"): If MaxPerPage=0 Then MaxPerPage=10
		  
		 Dim ClassID:ClassID=KS.ChkClng(KS.S("ClassID"))
		 Dim Param:Param=" IsTalk<>1 and UserID=" & UserID
		 Dim KeyWord:KeyWord=KS.S("Date") '日历搜索
		 Dim Key:Key=KS.R(KS.S("Key")) '表单搜索
		 Dim Tag:Tag=KS.R(KS.S("Tag")) 
		 If IsDate(KeyWord) Then
			   Param=Param & " And i.AddDate>=#" & KeyWord & " 00:00:00# and i.AddDate<=#" &KeyWord & " 23:59:59#"
		 End If
		 If ClassID<>0 Then Param=Param & " And i.ClassID=" & ClassID
		 If Key <>"" Then Param=Param & " And i.Title Like '%" & Key & "%'"
		 If Tag <>"" Then Param=Param & " And i.Tags Like '%" & Tag & "%'"
		 
		 If KS.S("Date")<>"" Then substr=substr & "<h2>搜索日期:<font color=red>" & KS.S("Date") & "</font>的博文</h2></br>"
		 If Tag<>"" Then substr=substr & "<h2>搜索Tag:<font color=red>" & Tag& "</font>的博文</h2></br>"
		 iF Key<>"" Then substr=substr & "<h2>搜索标题含有""<font color=red>" & Key & "</font>""的博文</h2></br>"
		 iF ClassID<>0 Then substr=substr & "<h2>搜索自定义分类ID""<font color=red>" & ClassID & "</font>""的博文</h2></br>"
		 
		 Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		 RSObj.Open "Select i.*,t.typename from KS_BlogInfo i inner join ks_blogtype t on i.typeid=t.typeid Where " & Param & " and i.Status=0 Order By i.AddDate Desc",Conn,1,1
		         If RSObj.EOF and RSObj.Bof  Then
				  	  If Key<>"" Then
						substr=substr &"<div style=""border: #efefef 1px dotted;text-align:center"" colspan=3><p>找不到日志标题含有<font color=red>""" & key & """</font>的博文!</p></div>"
						Else
						  If KeyWord="" And ClassID=0 Then
							substr=substr & "<div style=""border: #efefef 1px dotted;text-align:center"" colspan=3><p>您还没有写博文,<a href=""" & KS.GetDomain & "user/user_blog.asp?action=Add"">点此开始写博文</a>！</p></div>"
						  ElseIf ClassID<>0 Then
							substr=substr & "<div style=""border: #efefef 1px dotted;text-align:center"" colspan=3><p>找不到该分类的日志!</p></div>"
						  Else
							substr=substr & "<div style=""border: #efefef 1px dotted;text-align:center"" colspan=3><p>日期：<font color=red>" & KeyWord & "</font>,您没有写博文！</p></div>"
						  End If
					   End if
				 Else
							totalPut =conn.execute("Select count(1) from KS_BlogInfo i inner join ks_blogtype t on i.typeid=t.typeid Where " & Param & " and i.Status=0")(0)
							If CurrPage > 1 and (CurrPage - 1) * MaxPerPage < totalPut Then
										RSObj.Move (CurrPage - 1) * MaxPerPage
							End If
							call showlog(RSObj)
			End If
		 RSObj.Close:Set RSObj=Nothing
		 SubStr=SubStr & vbcrlf & ShowPage
	End Sub
    Sub showlog(RS)
		 Dim I,Url,Num
		 Num=(CurrPage-1)*MaxPerPage
		 Do While Not RS.Eof 
		   Url=KSBCls.GetLogUrl(rs)
		   substr=substr & "<div class=""loglist"">"
		   substr= substr & "<div class='t'><a class='title' href='" & Url & "'>" & rs("title") & "</a> <span class='tips'>" & rs("adddate") & "</span></div>"
		   
		   substr=substr &"<div class='intro'>" &KS.Gottopic(ks.losehtml(ks.ClearBadChr(ubbcode(rs("content"),1))),190) & "..."
		   if not ks.isnul(rs("photourl")) then
		    substr=substr & "<div class='pic'><a href='" & url & "' target='_blank' title='" & rs("title") & "'><img border='0' src='" & rs("photourl") & "' alt='" & rs("title") & "'/></a></div>"
		   end if
		   substr= substr & "</div><span class='tips'>分类：[<a target='_blank' href='" & KS.GetDomain & "space/morelog.asp?classid=" & rs("typeid") &"'>" & rs("typename") & "</a>] <a href=""" & Url  & """>阅读全文("&RS("hits")&")</a>  <a href=""" & Url  & "#Comment"">回复("& rs("totalput") &")</a></span>"
           substr=substr &"</div>"
		  RS.MoveNext
		 I = I + 1
		  If I >= MaxPerPage Then Exit Do
		 Loop
	End Sub
	
	 '好友列表
	 Sub FriendList()
	     MaxPerPage=20
	     substr=substr & KSBcls.Location("首页 >> " & UserName & "的好友")
	     substr=substr & "           <table border=""0"" align=""center"" class=""border"" width=""100%"">" & vbcrlf
		 Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		 RSObj.Open "select friend,u.userface,u.username from ks_friend f inner join ks_user u on f.friend=u.username where f.username='" & username & "' and f.accepted=1",Conn,1,1
		                 If RSObj.EOF and RSObj.Bof  Then
						  substr=substr & "<tr><td style=""BORDER: #efefef 1px dotted;text-align:center"" colspan=3>没有加好友！</td></tr>"
						 Else
							   totalPut = RSObj.RecordCount
							   If CurrPage >1 and (CurrPage - 1) * MaxPerPage < totalPut Then
									RSObj.Move (CurrPage - 1) * MaxPerPage
							   End If
								call showfriend(RSObj)
				           End If
		 RSObj.Close:Set RSObj=Nothing
		 substr=substr &  "    </table>" & vbcrlf
		 substr=substr & ShowPage
		End Sub
		
		sub showfriend(RS)
		    Dim I,k
			  Do While Not RS.Eof 
                 substr=substr & "<tr height=""20""> " &vbNewLine
				 for k=1 to 4
				 	   Dim UserFaceSrc:UserFaceSrc=RS("UserFace")
					   if lcase(left(userfacesrc,4))<>"http" then userfacesrc="../" & userfacesrc
                        substr=substr &"<td width=""25%"" style=""border: #efefef 1px dotted;"" align=""center""><a target=""_blank"" href=""" & KS.GetDomain & "space/?" & RS("userid") & """><img width=""80"" height=""80"" src=""" & UserFaceSrc & """ border=""0""></a><div align=""center""><a target=""_blank"" href=""blog.asp?username=" & RS("username") & """ target=""_blank"">" &RS(0) & "</a></div><a href=""javascript:void(0)"" onclick=""ksblog.addF(event,'" & rs("UserName") & "');""><img src=""images/adfriend.gif"" border=""0"" align=""absmiddle"" title=""加为好友"">好友</a> <a href=""javascript:void(0)"" onclick=""ksblog.sendMsg(event,'" & rs("username") & "')""><img src=""images/sendmsg.gif"" border=""0"" align=""absmiddle"" title=""发小纸条"">消息</a></td>" & vbnewline
			     RS.MoveNext
			     I = I + 1
				 If I >= MaxPerPage or rs.eof Then Exit for
				 next 
				 do while k<4
				  substr=substr & "<td width=""25%"">&nbsp</td>"
				  k=K+1
				 loop
                 substr=substr & "</tr> " & vbcrlf
				If I >= MaxPerPage Then Exit Do
			 Loop
	end sub
	 
	 '圈子列表
	 Sub GroupList()
	     substr=substr & KSBcls.Location("首页 >> 圈子")
		 MaxPerPage =10
		 substr=substr &"  <table border=""0"" class=""border"" align=""center"" width=""100%"">" & vbcrlf
		 Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		 RSObj.Open "select * from KS_team where username='" & username & "' and verific=1 order by id desc",Conn,1,1
		   If RSObj.EOF and RSObj.Bof  Then
			substr=substr & "<tr><td style=""border: #efefef 1px dotted;text-align:center"" colspan=3>没有创建圈子！</td></tr>"
		   Else
				totalPut = RSObj.RecordCount
				If CurrPage >1 and (CurrPage - 1) * MaxPerPage < totalPut Then
						RSObj.Move (CurrPage - 1) * MaxPerPage
				End If
				ShowGroup(RSObj)
		   End If
		 RSObj.Close:Set RSObj=Nothing
		 substr=substr &  "    </table>" & vbcrlf
		 substr=substr & ShowPage
	 End Sub
			 
	 Sub ShowGroup(RS)		 
		 Dim I
		 Do While Not RS.Eof 
		   substr=substr & "<tr style=""margin:2px;border-bottom:#9999CC dotted 1px;"">"
		   substr=substr & "<td width=""20%"" style=""border-bottom:#9999CC dotted 1px;"">"& vbcrlf
		   substr=substr & " <table style=""BORDER-COLLAPSE: collapse"" borderColor=#c0c0c0 cellSpacing=0 cellPadding=0 border=1>"
		   substr=substr &"	<tr>"
		   substr=substr & "		<td><a href=""group.asp?id=" & rs("id") & """ title=""" & rs("teamname") & """ target=""_blank""><img src=""" & rs("photourl") & """ width=""110"" height=""80"" border=""0""></a></td>"
		   substr=substr & "	 </tr>"
		   substr=substr & " </table>"
		   substr=substr & "</td>"
		   substr=substr & " <td style=""border-bottom:#9999CC dotted 1px;""><a class=""teamname"" href=""group.asp?id=" & rs("id") & """ title=""" & rs("teamname") & """ target=""_blank""> " & rs("TeamName") & "</a><br><font color=""#a7a7a7"">创建者：" & rs("username") & "</font><br><font color=""#a7a7a7"">创建时间:" &rs("adddate") & "</font><br>主题/回复：" & conn.execute("select count(id) from ks_teamtopic where teamid=" & rs("id") & "  and parentid=0")(0) & "/" & conn.execute("select count(id) from ks_teamtopic where teamid=" & rs("id"))(0) & "&nbsp;&nbsp;&nbsp;成员:" & conn.execute("select count(username)  from ks_teamusers where status=3 and teamid=" & rs("id"))(0) & "人  </td>"
		   substr=substr & "</tr>"
		   substr=substr & "<tr><td height='2'></td></tr>"
			rs.movenext
			I = I + 1
		  If I >= MaxPerPage Then Exit Do
		 Loop
	 End Sub
	 
	 Sub AlbumList()
	     SubStr=SubStr & KSBcls.Location("首页 >> 相册")
		 MaxPerPage =9
		 SubStr=SubStr & "  <div class=""albumlist"">" & vbcrlf
		 SubStr=SubStr & "   <ul>" & vbcrlf
		 Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		 RSObj.Open "Select * from KS_Photoxc Where username='" & username & "' and status=1 order by id desc",Conn,1,1
		  If RSObj.EOF and RSObj.Bof  Then
			 substr=substr & "<div style=""border: #efefef 1px dotted;text-align:center"">没有创建相册！</div>"
		  Else
							totalPut = RSObj.RecordCount
							
							If CurrPage>1 And (CurrPage - 1) * MaxPerPage < totalPut Then
								RSObj.Move (CurrPage - 1) * MaxPerPage
							Else
								CurrPage = 1
							End If
							 Dim I,k,Url
							 Do While Not RSObj.Eof 
									  substr=substr & "<li>" &vbcrlf
									          If KS.SSetting(21)="1" Then
											   Url="showalbum-" & RSObj("userid") & "-" & rsobj("id")
											  Else
											   Url="../space/?" & RSObj("userid") &"/showalbum/" &RSObj("id")
											  End If
											  substr=substr &"<div class=""albumbg""><a href=""" & url &""" target=""_blank""><img style=""margin-left:-4px;margin-top:5px"" src=""" &RSObj("photourl") &""" width=""120"" height=""90"" border=0></a></div><B><a href=""" & Url &""">" &RSObj("xcname") &"</a></B> (" & RSObj("xps") & ")<font color=red>[" & GetStatusStr(RSObj("flag")) &"]</font>" & vbcrlf
											  substr=substr &"</li>"
											RSObj.movenext
											I = I + 1
										  If I >= MaxPerPage Then Exit Do
										 Loop
				           End If
		 
		 substr=substr &  "    </ul></div>" & vbcrlf & ShowPage
		 
		 RSObj.Close:Set RSObj=Nothing
	 End Sub
	 

	 
	'查看相片
	 Sub ShowAlbum()
	   	SubStr=SubStr & KSBcls.Location("首页 >> 查看相册")

	   If ID=0 Then KS.Die "error xcid!"
	    Dim RSXC:Set RSXC=Server.CreateObject("ADODB.RECORDSET")
		RSXC.OPEN "Select top 1 * from ks_photoxc where id=" & id,conn,1,3
		if rsxc.eof and rsxc.bof then
		  rsxc.close:set rsxc=nothing
		  KS.Die "<script>alert('参数传递出错!');history.back();</script>"
		end if
	   If KS.C("AdminName") = "" Or KS.C("AdminPass") = "" Then
		If RSxc("Status")=0 Then
		 KS.Die "<script>alert('该相册尚未审核!');window.close();</script>"
		elseif RSxc("Status")=2 then
		 KS.Die "<script>alert('该相册已被管理员锁定!');window.close();</script>"
		end if
	   End If
	   KSBCls.Title=rsxc("xcname")
	   Select Case rsxc("flag")
		   Case 1,2
		    If rsxc("Flag")=2 and KS.C("UserName")="" then
			  SubStr=SubStr &"<br><br>此相册设置会员可见，请先<a href=""../User/"" target=""_blank"">登录</a>！"
			Else
			  GetAlbumBody rsxc("xcname")
		    End If
		  Case 3
		    If KS.S("Password")=rsxc("password") or Session("xcpass")=rsxc("password") then
			   If Not KS.IsNul(KS.S("Password")) Then Session("xcpass")=KS.S("Password")
			   GetAlbumBody rsxc("xcname")
			else
		      SubStr=SubStr &"<form action=""../space/?" & username &"/showalbum/" & id& """ method=""post"" name=""myform"" id=""myform"">请输入查看密码：<input type=""password"" name=""password"" size=""12"" style='border-style: solid; border-width: 1'>&nbsp;<input type='submit' value=' 查看 '></form>"
		   end if
		  Case 4
		    If KS.C("UserName")=rsxc("username") then
			  GetAlbumBody rsxc("xcname")
			else
			  SubStr=SubStr &"<br><br><li>该相册设为稳私，只有相册主人才有权利浏览!</li><li>如果你是相册主人，<a href=""../User/""  target=""_blank"">登录</a>后即可查看!</li>"
			end if
		 End Select
		 rsxc("hits")=rsxc("hits")+1
		 rsxc.update
		 rsxc.close:set rsxc=nothing
	 End Sub
	  Sub GetAlbumBody(xcname)
	             Dim TotalNum,RS,prevurl,nexturl
				 Set RS=Server.CreateObject("ADODB.RECORDSET")
				 RS.Open "Select * from KS_Photozp Where xcid=" & id &" Order By ID Desc",conn,1,1
				 If RS.EOF And RS.BOF Then
				    RS.Close : Set RS=Nothing
					SubStr=SubStr &"<p>该相册下没有照片！</p>"
				 Else
				        TotalNum=RS.Recordcount
				        If CurrPage>TotalNum Or CurrPage<=0 Then CurrPage=1
				        RS.Move(CurrPage-1)
						Conn.Execute("Update KS_PhotoZP Set Hits=hits+1 Where id=" & rs("id"))
						If KS.SSetting(21)="1" Then
						   prevurl="showalbum-" & userid & "-" & id & "-" & CurrPage-1
						   nexturl="showalbum-" & userid & "-" & id & "-" & CurrPage+1
						Else
						   prevurl="../space/?" & userid & "/showalbum/" & id & "/" & CurrPage-1
						   nexturl="../space/?" & userid & "/showalbum/" & id & "/" & CurrPage+1
						End If
						SubStr=SubStr &"<script type='text/javascript'>function getOffset(e){var target=e.target;if(target.offsetLeft==undefined){target=target.parentNode;}var pageCoord=getPageCoord(target);var eventCoord={x:window.pageXOffset+e.clientX,y:window.pageYOffset+e.clientY};var offset={offsetX:eventCoord.x-pageCoord.x,offsetY:eventCoord.y-pageCoord.y};return offset;}function getPageCoord(element){var coord={x:0,y:0};while(element){coord.x+=element.offsetLeft;coord.y+=element.offsetTop;element=element.offsetParent;}return coord;}</script>"&vbcrlf
						SubStr=SubStr &"<style type=""text/css"">.ruler{position:relative;}</style>"&vbcrlf
						SubStr=SubStr &"<div style='height:50px;line-height:50px;text-align:center'>（键盘左右键翻页）<a style='padding:3px;border:1px solid #cccccc' href='" & prevurl & "'>上一张</a> 第<font color=red>" & currpage & "</font>/" & TotalNum & "张 <a style='padding:3px;border:1px solid #cccccc' href='" & nexturl& "'>下一张</a> <a style='padding:3px;border:1px solid #cccccc' href=""" & RS("PhotoUrl") & """ target=""_blank"">查看原图</a></div><div style='padding-bottom:20px;text-align:center'><strong>相册名称:</strong>" & xcname &" <strong>浏览:</strong><font color=red>" & rs("hits") & "</font>次 <strong>大小:</strong>" & round(rs("photosize") /1024,2)  & " KB <strong>上传时间:</strong>" & rs("adddate") & "</div><div style='text-align:center'><img class=""ruler"" onmouseover=""upNext(this)"" id=""bigimg"" src='" & RS("PhotoUrl") & "' alt=""" & rs("descript") & """ style='border:1px solid #efefef' onload=""if (this.width>450) this.width=450;""/></div><div style='padding-top:20px;text-align:center'>" & rs("descript") & "</div>"
						substr=substr & "<script>" & vbcrlf
						substr=substr &" function upNext(bigimg){"&vbcrlf
						substr=substr &"var lefturl		= '" & prevurl & "';	var righturl	= '" & nexturl & "';"&vbcrlf
						substr=substr &"var imgurl		= righturl;var width	= bigimg.width;	var height	= bigimg.height;"&vbcrlf
						substr=substr &"bigimg.onmousemove=function(e){"&vbcrlf
						substr=substr &"var e=window.event || e,"&vbcrlf
						substr=substr &" posX=(e.offsetX==undefined) ? getOffset(e).offsetX : e.offsetX ;"&vbcrlf
						substr=substr &"if(posX<bigimg.width/2){"&vbcrlf
						substr=substr &"bigimg.style.cursor	= 'url(../images/default/arr_left.cur),auto';"&vbcrlf
						substr=substr &"imgurl				= lefturl;}else{"&vbcrlf
						substr=substr &"bigimg.style.cursor	= 'url(../images/default/arr_right.cur),auto';"&vbcrlf
						substr=substr &"imgurl				= righturl;}" &vbcrlf
						substr=substr &"}"&vbcrlf
						substr=substr &"bigimg.onmouseup=function(){top.location=imgurl;}}</script>"
						
		   		       RS.Close:Set  RS=Nothing
			    End If
				 SubStr=SubStr & "<script>document.onkeydown=chang_page;function chang_page(event){var e=window.event||event;var eObj=e.srcElement||e.target;var oTname=eObj.tagName.toLowerCase();if(oTname=='input' || oTname=='textarea' || oTname=='form')return;	event = event ? event : (window.event ? window.event : null);if(event.keyCode==37||event.keyCode==33){location.href='" & prevurl &"'}	if (event.keyCode==39 ||event.keyCode==34){location.href='" & nexturl & "'}}</script>"
		End Sub
		Function GetStatusStr(val)
           Select Case Val
		    Case 1:GetStatusStr="公开"
			Case 2:GetStatusStr="会员"
			Case 3:GetStatusStr="密码"
			Case 4:GetStatusStr="隐私"
		   End Select
			GetStatusStr="<font color=red>" & GetStatusStr & "</font>"
		End Function
		
		'留言
		Sub ShowMessage()
		 
		 SubStr=SubStr & KSBcls.Location("首页 >> 留言板(<a href=""#write"">签写留言</a>)")
		 
		 SubStr=substr & "" &  GetWriteMessage() 
		 MaxPerPage =8
		 SubStr=SubStr &"  <table border=""0"" align=""center""  width=""100%"">" & vbcrlf
		 Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		 RSObj.Open "Select * from KS_BlogMessage Where UserName='" & UserName & "' and status=1 Order By AddDate Desc",Conn,1,1
		         If RSObj.EOF and RSObj.Bof  Then
				  RSObj.Close :Set RSObj=Nothing
				 SubStr=SubStr &"<tr><td style=""background:#FBFBFB;padding:10px;border: #efefef 1px dotted;text-align:center"">还没有人给主人留言哦!</td></tr></table>"
				 	 ExiT Sub
				 Else
					   totalPut = Conn.Execute("Select count(1) From KS_BlogMessage Where UserName='" & UserName & "' and status=1")(0)
						If CurrPage >1 and (CurrPage - 1) * MaxPerPage < totalPut Then
										RSObj.Move (CurrPage - 1) * MaxPerPage
						End If
						Call showguest(RSObj)
                End If
		 
		 substr=substr &  "</table>" & vbcrlf & ShowPage
		 RSObj.Close:Set RSObj=Nothing
		
		
		End Sub
		
		Sub ShowGuest(rs)
		 Dim I,CommentStr,n
		  CommentStr="<br/><div style='border-bottom:1px solid #f1f1f1;padding-bottom:3px;font-weight:bold;font-size:14px'>&nbsp;&nbsp;共有 <font color=red>" & totalPut & " </font> 条留言信息，共分 <font color=red>" & pagenum & "</font> 页,第 <font color=red>" & CurrPage & "</font> 页</div>"
			If CurrPage=1 Then
			 N=TotalPut
			 Else
			 N=totalPut-MaxPerPage*(CurrPage-1)
			 End IF
		  Dim RSU,FaceStr,Publish,MoreStr,Rname
		  Do While Not RS.Eof 
		   FaceStr=KS.Setting(3) & "images/face/boy.jpg"
		
			Publish=KS.R(RS("AnounName"))
			Set RSU=Conn.Execute("Select top 1 UserFace,userid,UserName,RealName From KS_User Where UserName='"& Publish & "'")
			If Not RSU.Eof Then
			  FaceStr=rsu(0) : Rname=KS.CheckXSS(rsu(3)) : If KS.IsNul(Rname) Then Rname=RSU(2)
			  Publish="<a href='" & KS.GetSpaceUrl(RSU(1)) & "' target='_blank'>" & Rname & "</a>"
			  If lcase(left(FaceStr,4))<>"http" And Left(FaceStr,1)<>"/" then FaceStr=KS.Setting(3) & FaceStr
			Else
			  Publish="<a href='#'>" & Publish & "</a>"
		    End IF
			RSU.Close

		   CommentStr=CommentStr & "<tr><td valign='top' style='padding-bottom:10px;margin-top:3px;BORDER-BOTTOM: #999999 1px dotted;'><table border='0' width='100%' cellspacing='0' cellpadding='0'><tr><td rowspan='2' width='80' align='center'><img width=""60"" height=""60"" src=""" & facestr & """ border=""1"" /></td><td> <span style='color:#999'>第" & N & "楼 " & publish & " 留言于：" & RS("AddDate") &"</span>" 
			 IF KS.C("UserName")=UserName Then
			  MoreStr="<a href='#'>顶部</a> | <a href='../User/user_message.asp?Action=MessageDel&id=" & RS("ID") & "' onclick=""return(confirm('确定删除该留言吗?'))"">删除</a> | <a href='../user/user_message.asp?id=" & RS("ID") & "&Action=ReplayMessage' target='_blank'>回复</a>"
             Else
			  MoreStr="<a href='#'>顶部</a>"
			 End If
		   CommentStr=CommentStr & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & MoreStr & ""
		   
		   CommentStr=CommentStr &"<br/>"
		   If Not KS.IsNUL(RS("Title")) Then
		   CommentStr=CommentStr & RS("Title") & "<br/>"
		   End If
		   CommentStr=CommentStr & Replace(RS("Content"),chr(10),"<br/>")
		   
		    If Not KS.IsNul(RS("Replay")) Then
			 CommentStr=CommentStr&"<div style=""margin : 5px 20px; border : 1px solid #efefef; padding : 5px;background : #ffffee; line-height : normal;""><b>以下为space主人的回复:</b><br>" & RS("Replay") & "<br><div align=right>时间:" & rs("replaydate") &"</div></div>"
			End If
				 
		   CommentStr=CommentStr & "	 </td></tr></table></td>"
		   CommentStr=CommentStr & "</tr>"
		
		   N=N-1
		   RS.MoveNext
				I = I + 1
			  If I >= MaxPerPage Then Exit Do
		  loop
		 SubStr=SubStr & CommentStr
		End Sub
		
		Function GetWriteMessage()
		
		 If KS.SSetting(25)="0" And KS.IsNul(KS.C("UserName")) Then
		  GetWriteMessage="<div style=""margin:20px""><strong>温馨提示：</strong>只有会员才可以留言,如果是会员请先<a href=""javascript:ShowLogin()"">登录</a>,不是会员请点此<a href=""../user/reg/"" target=""_blank"">注册</a>。</div>"
		 Else
		 GetWriteMessage = "<div style=""clear:both""></div><iframe src=""about:blank"" name=""siframe"" width=""0"" height=""0""></iframe><a name=""write""></a><table width=""98%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"" class=""comment_write_table"">"
		 GetWriteMessage = GetWriteMessage & "<form name=""msgform"" action=""" & KS.GetDomain &"plus/ajaxs.asp?action=MessageSave"" method=""post"" target=""siframe"">"
		 GetWriteMessage = GetWriteMessage & "<input type=""hidden"" value=""" & UserName & """ name=""UserName"">"
		 GetWriteMessage = GetWriteMessage & "<input type=""hidden"" value=""" & UserId & """ name=""UserId"">"
		 GetWriteMessage = GetWriteMessage & "<input type=""hidden"" value="""" name=""scontent"">"
		 GetWriteMessage = GetWriteMessage & "<tr><td class=""comment_write_title""><span style='font-weight:bold;font-size:14px'>发表您的留言:</span><br/><textarea class=""msgtextarea"" cols='70' rows='4' id=""Content"" onfocus=""if (this.value=='既然来了，就顺便留句话儿吧...') this.value='';"" name=""Content"" onblur=""if (this.value=='') this.value='既然来了，就顺便留句话儿吧...';"">既然来了，就顺便留句话儿吧...</textarea><br/>昵称："
		
		GetWriteMessage = GetWriteMessage & "   <input name=""AnounName"""
		If KS.C("UserName")<>"" Then GetWriteMessage = GetWriteMessage & " readonly"
		GetWriteMessage = GetWriteMessage & " maxlength=""100"" type=""text"" value=""" & KS.C("UserName") & """ id=""AnounName"" style=""background:#FBFBFB;color:#999;border:1px solid #ccc;width:100px;height:20px;line-height:20px""/>&nbsp;<font color=red>*</font> <span>验证码 </span><script>writeVerifyCode("""&KS.GetDomain&""",1)</script><br/><input type=""submit"" onclick=""return(CheckForm());""  name=""SubmitComment"" value=""OK了，提交留言"" class=""btn""/>"
		GetWriteMessage = GetWriteMessage & "    </td>"
		GetWriteMessage = GetWriteMessage & "  </tr>"
		GetWriteMessage = GetWriteMessage & "  </form>"
		GetWriteMessage = GetWriteMessage & "</table>"
		End If
		End Function 
		
		Sub GetNews()
		 Dim SQL,i,param
		 SubStr=KSBcls.Location("首页 >> 公司动态")
		 Dim RS:Set RS=Conn.Execute("Select classid,classname from ks_userclass where username='" & UserName & "' and typeid=4 order by orderid")
		 If Not RS.Eof Then SQL=RS.GetRows(-1)
		 RS.Close:Set RS=Nothing
		 If IsArray(SQL) tHEN
		     SubStr=SubStr &"<h3>按分类查看</h3><img width='50' src='images/search.png' align='absmiddle'>"
			 if ID=0 Then
			  SubStr=SubStr &"<a href='../space/?" & userid & "/" & Action & "/'><font color=red>全部文章</font></a>&nbsp;&nbsp;"
			 else
			  SubStr=SubStr &"<a href='../space/?" & userid & "/" & Action & "/'>全部文章</a>&nbsp;&nbsp;"
			 end if
			 For I=0 To Ubound(SQL,2)
			   if ID=SQL(0,I) then
			   SubStr=SubStr & "<a href='../space/?" & userid & "/" & action & "/" & SQL(0,i) & "'><font color=red>" & SQL(1,i) & "(" & conn.execute("select count(id) from ks_enterprisenews where classid=" & sql(0,i))(0) &")</font></a>&nbsp;&nbsp;&nbsp;&nbsp;"
			   else
			   SubStr=SubStr & "<a href='../space/?" & userid & "/" & action & "/" & SQL(0,i) & "'>" & SQL(1,i) & "(" & conn.execute("select count(id) from ks_enterprisenews where classid=" & sql(0,i))(0) &")</a>&nbsp;&nbsp;&nbsp;&nbsp;"
			   end if
			 Next
		 End If
		 if ID=0 Then
		 SubStr=SubStr &"<h3>所有新闻</h3>"
		 Else
		 SubStr=SubStr &"<h3>" & Conn.Execute("Select ClassName From KS_UserClass Where ClassID=" & ID)(0) & "</h3>"
		 End If
		 MaxPerPage=10
		 param=" Where UserName='" & UserName & "'"
		 If ID<>0 Then Param=Param & " and classid=" & id
		 SubStr=SubStr & "<table border='0' width='98%' align='center' cellspacing='2' cellpadding='2'>"
		 Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select ID,Title,AddDate From KS_EnterPriseNews " & Param &" order by adddate desc",conn,1,1
		 If RS.EOF and RS.Bof  Then
			 SubStr=SubStr & "<tr><td style=""border: #efefef 1px dotted;text-align:center;height:80px;"" colspan=3><p>没有发布动态文章,请<a href='../user/user_EnterPriseNews.asp?Action=Add' target='_blank'><font color=red>点此发布</font></a>！</p></td></tr>"
		Else
			  totalPut = RS.RecordCount
				If CurrPage>  1 and (CurrPage - 1) * MaxPerPage < totalPut Then
						RS.Move (CurrPage - 1) * MaxPerPage
				End If
				SQL=RS.GetRows(-1)
				 Dim K,N,Total,url
				 Total=Ubound(SQL,2)+1
				 For I=0 To Total
					If KS.SSetting(21)="1" Then Url="show-news-" & userid & "-" & sql(0,n) & KS.SSetting(22) Else Url="../space/?" & userid & "/shownews/" & sql(0,n)
					SubStr=SubStr &"<tr>"
					SubStr=SubStr & "<td style=""border-bottom: #efefef 1px dotted;height:22""><img src='../images/default/arrow_r.gif' align='absmiddle'> <a href='" & url & "' target='_blank'>" & SQL(1,N) & "</a>&nbsp;" & sql(2,n)
					SubStr=SubStr & "</td>"
					N=N+1
					If N>=Total Or N>=MaxPerPage Then Exit For
				   SubStr=SubStr &"</tr>"
				 Next
		 End If
		  SubStr=SubStr &"</table>" 
		  SubStr=SubStr & "<div id=""kspage"">" & ShowPage() & "</div>"
		  
		End Sub
		
		'显示新闻详情
		Sub ShowNews()
		 Dim SQL,i,RS,PhotoUrl,url
		 SubStr=KSBcls.Location("首页 >> 公司动态 >> 查看新闻")
		 SubStr=SubStr & "<table border='0' width='98%' align='center' cellspacing='2' cellpadding='2'>"
		 Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 * From KS_EnterPriseNews Where UserName='" & UserName & "' and ID=" & ID  ,conn,1,1
		 If RS.EOF and RS.Bof  Then
		     RS.Close:Set RS=Nothing
			 KS.Die "<script>alert('参数传递出错！');window.close();</script>"
		 Else
		   KSBCls.Title=rs("Title")
		   SubStr=SubStr &"<tr><td align='center' style='color:#ff6600;font-weight:bold;font-size:14px'><div style=""font-weight:bold;text-align:center"">" & rs("title") & "</div></td></tr>"
		   SubStr=SubStr & "<tr><td><div style=""text-align:center"">作者：" & UserName & "&nbsp;&nbsp;&nbsp;&nbsp;时间:" & RS("AddDate") & "</div>"
		   SubStr=SubStr & "<hr size=1><div>" & KS.HTMLCode(rs("content")) & "</div></td></tr>"
		   If KS.SSetting(21)="1" Then Url="news-" & username  Else Url="../space/?" & username & "/news"
		   SubStr=SubStr &"<tr><td><div style='text-align:center'><a href='" & Url & "'>[返回公司动态]</a></div></td></tr>"
		 End If
		 SubStr=SubStr &"</table>"   
         RS.Close:Set RS=Nothing
		End Sub
		
		'产品列表
		Function ProductList()
		 Dim SQL,i,param,classUrl
		 SubStr=KSBcls.Location("首页 >> 产品展示")
		 Dim RS:Set RS=Conn.Execute("Select classid,classname from ks_userclass where username='" & UserName & "' and typeid=3 order by orderid")
		 If Not RS.Eof Then SQL=RS.GetRows(-1)
		 RS.Close:Set RS=Nothing
		 If IsArray(SQL) tHEN
		     SubStr=SubStr &"<h3>按分类查看</h3><img width='50' src='images/search.png' align='absmiddle'>"
			 If KS.SSetting(21)="1" Then classUrl="product-" & userid Else classUrl="../space/?" & Userid & "/product"
			 if ID=0 Then
			  SubStr=SubStr & "<a href='" & classUrl & "'><font color=red>全部产品</font></a>&nbsp;&nbsp;"
			 else
			  SubStr=SubStr &"<a href='" & classUrl & "'>全部产品</a>&nbsp;&nbsp;"
			 end if
			 For I=0 To Ubound(SQL,2)
			   If KS.SSetting(21)="1" Then classUrl="product-" & userid & "-" & SQL(0,I) & ks.SSetting(22) Else classUrl="../space/?" & Userid & "/product/" & SQL(0,i)
			   if ID=SQL(0,I) then
			   SubStr=SubStr & "<a href='" & ClassURL & "'><font color=red>" & SQL(1,i) & "(" & conn.execute("select count(id) from ks_product where verific=1 and classid=" & sql(0,i))(0) &")</font></a>&nbsp;&nbsp;&nbsp;&nbsp;"
			   else
			   SubStr=SubStr & "<a href='" & ClassURL & "'>" & SQL(1,i) & "(" & conn.execute("select count(id) from ks_product where verific=1 and classid=" & sql(0,i))(0) &")</a>&nbsp;&nbsp;&nbsp;&nbsp;"
			   end if
			 Next
		 End If
		 if ID=0 Then
		 SubStr=SubStr &"<h3>所有产品</h3>"
		 Else
		 SubStr=SubStr &"<h3>" & Conn.Execute("Select top 1 classname from ks_userclass where classid=" &ID)(0) & "</h3>"
		 End If
		 MaxPerpage=12
		 param=" Where verific=1 and Inputer='" & UserName & "'"
		 If ID<>0 Then Param=Param & " and classid=" & id
		 SubStr=SubStr & "<table border='0' width='98%' align='center' cellspacing='2' cellpadding='2'>"
		 Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select ID,Title,PhotoUrl From KS_Product " & Param &" order by adddate desc",conn,1,1
		 If RS.EOF and RS.Bof  Then
			SubStr=SubStr & "<tr><td style=""border: #efefef 1px dotted;text-align:center;height:80px;"" colspan=3><p>没有发布产品展示,请<a href='../user/user_myshop.asp?Action=Add' target='_blank'><font color=red>点此发布</font></a>！</p></td></tr>"
		Else
			  totalPut = RS.RecordCount
				If CurrPage> 1 and (CurrPage - 1) * MaxPerPage < totalPut Then
						RS.Move (CurrPage - 1) * MaxPerPage
				End If
				SQL=RS.GetRows(MaxperPage)
				Dim K,N,Total,PhotoUrl,Url
				 Total=Ubound(SQL,2)+1
				 For I=0 To Total
				   SubStr=SubStr &"<tr>"
				   For K=1 To 4
					PhotoUrl=SQL(2,N)
					If KS.SSetting(21)="1" Then Url="show-product-" &userid & "-" & sql(0,n) & KS.SSetting(22) Else url="../space/?" & Userid & "/showproduct/" & sql(0,n)
					If PhotoUrl="" Or IsNull(PhotoUrl) Then PhotoUrl="../images/nopic.gif"
					SubStr=SubStr & "<td align='center'>" 
					SubStr=SubStr & "<a href='" & Url & "' target='_blank'><Img border='0' src='" & PhotoUrl & "' alt='" & SQL(1,N) & "' width='130' height='90' /></a><div style='text-align:center'><a href='" & Url & "'>" & KS.Gottopic(SQL(1,N),20) & "</a></div>"
					SubStr=SubStr & "</td>"
					N=N+1
					If N>=Total Or N>=MaxPerPage Then Exit For
				   Next
				   SubStr=SubStr &"</tr>"
				   If N>=Total  Or N>=MaxPerPage Then Exit For
				 Next
		 End If
		 SubStr=SubStr &"</table>" 
		 SubStr=SubStr & ShowPage()  
		End Function
		
		'查看产品详情
		Function ShowProduct()
		 Dim SQL,i,RS,PhotoUrl
		 SubStr=KSBcls.Location("首页 >> 产品展示 >> 查看产品详情")
		 SubStr=SubStr & "<table border='0' width='98%' align='center' cellspacing='2' cellpadding='2'>"
		 Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 * From KS_Product Where inputer='" & UserName & "' and ID=" & ID ,conn,1,1
		 If RS.EOF and RS.Bof  Then
		     RS.Close:Set RS=Nothing
			 KS.Die "<script>alert('参数传递出错！');window.close();</script>"
		 Else
		   KSBCls.Title=RS("title")
		   photourl=RS("BigPhoto")
		   If PhotoUrl="" Or IsNull(photourl) Then photourl="../images/nopic.gif"
		   SubStr=SubStr &"<tr><td align='center' style='color:#ff6600;font-weight:bold;font-size:14px'>" & rs("Title") & "</td></tr>"
		   SubStr=SubStr & "<tr><td align='center'><img  style='max-width:600px;width:600px;width:expression(document.body.clientWidth>600?""600px"":""auto"");overflow:hidden;' src='" & photourl &"' border='0'></td></tr>"
		   SubStr=SubStr & "<tr><td><h3>基本参数</h3></td></tr>"
		   SubStr=SubStr & "<tr><td>生 产 商：" & RS("ProducerName") & "</td></tr>"
		   SubStr=SubStr & "<tr><td>产品分类：" & KS.C_C(RS("tid"),1) & "</td></tr>"
		   SubStr=SubStr & "<tr><td>产品型号：" & RS("ProModel") & "</td></tr>"
		   SubStr=SubStr & "<tr><td>品牌/商标：" & RS("TrademarkName") & "</td></tr>"
		   SubStr=SubStr & "<tr><td>生 产 商：" & RS("ProducerName") & "</td></tr>"
		   SubStr=SubStr & "<tr><td>市 场 价：￥" & RS("price") & " 元</td></tr>"
		   SubStr=SubStr & "<tr><td>商 城 价：￥" & RS("price_member") & " 元</td></tr>"
		   If KS.C_S(5,21)="1" Then
		   SubStr=SubStr & "<tr><td>在线购买：<a target='_blank' href=""" & KS.GetItemURL(5,rs("Tid"),rs("ID"),rs("Fname"))   & """><img src='" & KS.GetDomain & "images/ProductBuy.gif' align='absmiddle' border='0'/></a></td></tr>"
		   End If
		   SubStr=SubStr & "<tr><td><h3>详细介绍</h3></td></tr>"
		   SubStr=SubStr & "<tr><td align=""left"">" & bbimg(KS.HtmlCode(RS("proIntro"))) & "</td></tr>"
		 End If
		 SubStr=SubStr &"</table>"   
         RS.Close:Set RS=Nothing
		End Function
		
		Private Function bbimg(strText)
		Dim s,re
        Set re=new RegExp
        re.IgnoreCase =true
        re.Global=True
		s=strText
		re.Pattern="<img(.[^>]*)([/| ])>"
		s=re.replace(s,"<img$1/>")
		re.Pattern="<img(.[^>]*)/>"
		s=re.replace(s,"<img$1  onclick=""window.open(this.src)"" style='display:block;max-width:600px;width:600px;width:expression(document.body.clientWidth>600?""600px"":""auto"");overflow:hidden;'/>")
		bbimg=s
	End Function
		
		'招聘
		Sub JobList()
		   SubStr=KSBcls.Location("首页 >> 企业招聘")
		 If KS.C_S(10,21)="0" Then 
		   Dim Jrs:set Jrs=Conn.Execute("Select top 1 Job From ks_Enterprise where username='" & UserName & "'")
		   If Not Jrs.Eof Then
		    SubStr=SubStr & KS.HTMLCode(Jrs(0))
		   Else
		    Jrs.Close: Set Jrs=Nothing
		    KS.AlertHintScript "对不起，该用户不是企业用户！"
		   End If
		   Jrs.Close
		   Set Jrs=Nothing
		   Exit Sub
		 End If
		 
		 SubStr=SubStr &"<h3>招聘信息</h3>"
		 MaxPerPage=5
		 Dim Param,rs,sql
		 param=" Where status=1 and UserName='" & UserName & "'"
		 SubStr=SubStr & "<table border='0' width='98%' align='center' cellspacing='2' cellpadding='2'>"
		 Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select ID,JobTitle,province,city,workexperience,num,salary,refreshtime,status,intro,sex From KS_Job_ZW " & Param &" order by refreshtime desc",conn,1,1
		 If RS.EOF and RS.Bof  Then
			 SubStr=SubStr & "<tr><td style=""border: #efefef 1px dotted;text-align:center;height:80px;"" colspan=3><p>没有发布招聘信息,请<a href='../User/User_JobCompanyZW.asp?Action=Add' target='_blank'><font color=red>点此发布</font></a>！</p></td></tr>"
		Else
			  totalPut = RS.RecordCount
				If CurrPage>  1 and (CurrPage - 1) * MaxPerPage < totalPut Then
						RS.Move (CurrPage - 1) * MaxPerPage
				End If
				SQL=RS.GetRows(MaxperPage)
				Dim I,K,N,Total,PhotoUrl,url
				 Total=Ubound(SQL,2)
				 For I=0 To Total
				     SubStr=SubStr &"<tr><td style='line-height:180%;padding-top:6px;padding-bottom:8px;border-bottom:1px solid #cccccc;'>"
					 SubStr=SubStr & "<font color=#ff6600>岗位名称：" & sql(1,i) & "</font>&nbsp;&nbsp;<a href='" & JLCls.GetZWUrl(SQL(0,I)) & "' target='_blank'>浏览详情</a><br>工作地点：" & SQL(2,I) & "&nbsp;" & SQL(3,I) & "&nbsp;&nbsp;招聘人数：" & SQL(5,I) & " 人<BR>"
					 SubStr=SubStr& "发布日期：" & sql(7,i) & "&nbsp;&nbsp;性别要求：" & SQL(10,I) & "<br>详细介绍：" & SQL(9,I) & "</td>"
				     SubStr=SubStr &"</tr>"
				 Next
		 End If
		 SubStr=SubStr &"</table>"
		 SubStr=SubStr & ShowPage
		End Sub
		
		'荣誉证书
		Sub GetRyzs()
		Dim SQL,i,param,RS
		 Substr=KSBcls.Location("首页 >> 荣誉证书")
		 SubStr=SubStr &"<h3>荣誉证书</h3>"

		 param=" Where status=1 and UserName='" & UserName & "'"
		 SubStr=SubStr & "<table style='margin-bottom:5px' border='0' width='98%' align='center' cellspacing='1' cellpadding='0' bgcolor='#FFFFFF'>"
		 SubStr=SubStr & "<tr bgcolor='#F3F3F3' align='center'><td width='20%' height='20'>证收照片</td><td width='24%'>证书名称</td><td width='21%'>发证机构</td><td width='17%'>生效日期</td><td width='18%'>截止日期</td></tr>"
		 Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select ID,Title,FZJG,sxrq,jzrq,photourl From KS_EnterPriseZS " & Param &" order by adddate desc",conn,1,1
		 If RS.EOF and RS.Bof  Then
			SubStr=SubStr & "<tr><td style=""border: #efefef 1px dotted;text-align:center;height:80px;"" colspan=6><p>没有发布荣誉证书,请<a href='../user/user_EnterPriseZS.asp?Action=Add' target='_blank'><font color=red>点此发布</font></a>！</p></td></tr>"
		Else
			  totalPut = RS.RecordCount
				If CurrPage>  1 and (CurrPage - 1) * MaxPerPage < totalPut Then
						RS.Move (CurrPage - 1) * MaxPerPage
				End If
				SQL=RS.GetRows(MaxperPage)
				Dim K,N,Total,PhotoUrl,url,BeginDateStr,EndDateStr
		 Total=Ubound(SQL,2)
		 For I=0 To Total
		   if i mod 2=0 then
		    SubStr=SubStr &"<tr bgcolor='#ffffff'>"
		   else
		    SubStr=SubStr & "<tr bgcolor='#f6f6f6'>"
		   end if
		    PhotoUrl=SQL(5,i)
			If PhotoUrl="" Or IsNull(PhotoUrl) Then PhotoUrl="../images/nopic.gif"
			BeginDateStr=SQL(3,I) :	If Not IsDate(BeginDateStr) Then BeginDateStr=Now
			EndDateStr =SQL(4,I) : If Not IsDate(EndDateStr) Then EndDateStr=Now
		    SubStr=SubStr & "<td width='150' style='height:80px;text-align:center;padding-top:6px;padding-bottom:8px;'>" 
			SubStr=SubStr & "<a href='" & PhotoUrl & "' target='_blank'><Img border='0' src='" & PhotoUrl & "' width='85' height='60'></a>"
			SubStr=SubStr & "</td>"
			SubStr=SubStr & "<td style='text-align:center;line-height:150%;' >" & sql(1,i) & "</td>"
			SubStr=SubStr & "<td style='text-align:center;line-height:150%;' >" & sql(2,i) & "</td>"
			SubStr=SubStr & "<td style='text-align:center;line-height:150%;' >" & year(BeginDateStr) & "年" & month(BeginDateStr) & "月</td>"
			SubStr=SubStr & "<td style='text-align:center;line-height:150%;' >" & year(EndDateStr) & "年" & month(EndDateStr) & "月</td>"
		    SubStr=SubStr &"</tr>"
		 Next
		 End If
		 SubStr=SubStr &"</table>" 
		 SubStr=SubStr & ShowPage  
		End Sub
		
		'显示图片
		Function ShowPhoto()
		 Dim SQL,n,RS,PhotoUrlArr,PhotoUrl,t
		 substr=KSBcls.Location("首页 >> 作品展示 >> 查看作品")
		 substr=substr & "<table border='0' width='98%' align='center' cellspacing='2' cellpadding='2'>"
		 Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select * From KS_Photo Where Inputer='" & UserName & "' and ID=" & ID  ,conn,1,1
		 If RS.EOF and RS.Bof  Then
		     RS.Close:Set RS=Nothing
			 KS.Die "<script>alert('参数传递出错！');window.close();</script>"
		 Else
		   KSBCls.Title = rs("title")
		   photourlArr=Split(RS("PicUrls"),"|||")
		   n=CurrPage
		   if n<0 then n=0
		   t=Ubound(PhotoUrlArr)
		   If N>=t Then n=0
		   If t=0 Then t=1
		   PhotoUrl=Split(PhotoUrlArr(N),"|")(1)
		   substr=substr & "<tr><td align='center' class='divcenter_work_on'><div class='fpic'><a href='../space/?" & UserName & "/showphoto/" & ID & "/" & n+1 &"'><img  onload=""var myImg = document.getElementById('myImg'); if (myImg.width >580 ) {myImg.width =580 ;};"" id=""myImg"" src='" & photourl &"' title='查看下一张' border='0'></A></div></td></tr>"
		   substr=substr &"<tr><td height='35' align='center'>浏览：<Script Src='../item/GetHits.asp?Action=Count&m=2&GetFlag=0&ID=" & ID & "'></Script> 总得票：<Script Src='../item/GetVote.asp?m=2&ID=" & ID & "'></Script> 投票：<a href='../item/Vote.asp?m=2&ID=" & ID & "'>投它一票</a></td></tr>"
           substr=substr & "<tr><td height='35' align='center'>第" & N+1 & "/" & t & "张 <a href='../space/?" & UserID & "/showphoto/" & ID &"/0'><img src='images/picindex.gif' border='0'></a>&nbsp;<a href='../space/?" & UserID & "/showphoto/" & id &"/" & N-1 & "'><img src='images/picpre.gif' border='0'></a>&nbsp;<a href='../space/?" & UserID & "/showphoto/" & id &"/" & N+1 & "'><img src='images/picnext.gif' border='0'></a>&nbsp;<a href='../space/?" & UserID & "/showphoto/" & id &"/" & t-1 & "'><img src='images/picend.gif' border='0'></a></td></tr>"
		   substr=substr & "<tr><td><span class=""writecomment""><Script Language=""Javascript"" Src=""../plus/Comment.asp?Action=Write&ChannelID=2&InfoID=" &id & """></Script></span></td></tr>"
		   substr=substr & "<tr><td>&nbsp;<Img src='images/topic.gif' align='absmiddle'> <strong>作品评论：</strong><br><span class=""showcomment""><script src=""../ks_inc/Comment.page.js"" language=""javascript""></script><script language=""javascript"" defer>Page(1,2,'" & ID & "','Show','../');</script><div id=""c_" & ID & """></div><div id=""p_" & ID & """ align=""right""></div> </span></td></tr>"
		 End If
		 substr=substr &"</table>"   
		End Function
		
		
		'信息集
		Sub xxList()
		If KS.IsNUL(Request.ServerVariables("QUERY_STRING")) Then KS.Die "error"
		Dim QueryParam:QueryParam=Request.ServerVariables("QUERY_STRING")&"////"
		Dim Channelid:ChannelID=KS.ChkClng(Split(QueryParam,"/")(2))
		if channelid=0 then channelid=1
		Dim SQL,K,OPStr,RSC:Set RSC=Conn.Execute("Select ChannelID,itemName From KS_Channel Where ChannelStatus=1 and channelid<>6  And ChannelID<>9 And ChannelID<>10 order by channelid")
		SQL=RSc.GetRows(-1)
		RSc.Close:set RSc=Nothing
		For K=0 To Ubound(SQL,2)
		 if channelid=sql(0,k) then
		 OpStr=OpStr & "<option value='../space/?" & userid & "/xx/" & SQL(0,K) & "' selected>" & SQL(1,K) & "</option>"
		 else
		 OpStr=OpStr & "<option value='../space/?" & userid & "/xx/" & SQL(0,K) & "'>" & SQL(1,K) & "</option>"
		 end if
		Next
	    substr= KSBcls.Location("首页 >> 信息集")
		Substr=Substr& "<div class=""xxcategory"">信息分类<select name='channelid' onchange=""location.href=this.value"">" & opstr & "</select></div>"
		 MaxPerPage =20
		 substr=substr & "  <table border=""0"" class=""border"" align=""center"" width=""100%"">" & vbcrlf

		 Dim Sqlstr
		 Select Case KS.C_S(ChannelID,6) 
		  Case 1
		   SQLStr="Select ID,Title,Tid,ReadPoint,InfoPurview,Fname,Changes,AddDate from " & KS.C_S(ChannelID,2) & " Where Inputer='" & UserName & "'  and deltf=0 and verific=1 Order By AddDate Desc"
		  Case 2
  		   SQLStr="Select ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate,photourl from " & KS.C_S(ChannelID,2) & " Where Inputer='" & UserName & "'  and deltf=0 and verific=1 Order By AddDate Desc"
          Case 3
  		   SQLStr="Select ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate from " & KS.C_S(ChannelID,2) & " Where Inputer='" & UserName & "'  and deltf=0 and verific=1 Order By AddDate Desc"
		  Case 4
  		   SQLStr="Select ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate from " & KS.C_S(ChannelID,2) & " Where Inputer='" & UserName & "'  and deltf=0 and verific=1 Order By AddDate Desc"
		  Case 5
  		   SQLStr="Select ID,Title,Tid,0,0,Fname,0,AddDate from " & KS.C_S(ChannelID,2) & " Where Inputer='" & UserName & "'  and deltf=0 and verific=1 Order By AddDate Desc"
		  Case 7
  		   SQLStr="Select ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate from " & KS.C_S(ChannelID,2) & " Where Inputer='" & UserName & "'  and deltf=0 and verific=1 Order By AddDate Desc"
		  Case 8
  		   SQLStr="Select ID,Title,Tid,0,0,Fname,0,AddDate from " & KS.C_S(ChannelID,2) & " Where inputer='" & UserName & "'  and deltf=0 and verific=1 Order By AddDate Desc"
		 End Select

		 Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		 RSObj.Open SqlStr,Conn,1,1
		         If RSObj.EOF and RSObj.Bof  Then
					substr=substr & "<tr><td style=""border: #efefef 1px dotted;text-align:center"" colspan=3><p>找不到您要的信息！</p></td></tr>"
				 Else
							totalPut = RSObj.RecordCount
							 If CurrPage>1 and (CurrPage - 1) * MaxPerPage < totalPut Then
										RSObj.Move (CurrPage - 1) * MaxPerPage
							 End If
						call showxx(RSObj,channelid)
				End If
		 
		 substr=substr &  "            </table>" & vbcrlf
		 substr=substr & showpage
		 RSObj.Close:Set RSObj=Nothing
	End Sub	
	
	Sub showxx(rs,channelid)
		if KS.C_S(ChannelID,6) =2 then       '图片显示不同
		   substr=substr & GetUserPhoto(RS,MaxPerPage,ChannelID)
		Else
			 Dim K,SQL
			 Do While Not RS.Eof
				substr=substr & "<tr><td style=""border-bottom: #efefef 1px dotted;height:22""><img src=""../images/default/arrow_r.gif"" align=""absmiddle""> [" & KS.GetClassNP(RS(2)) & "] <a href='" & KS.GetItemUrl(channelid,RS(2),RS(0),RS(5)) & "' target='_blank'>" & RS(1) & "</a>&nbsp;&nbsp;(" & RS(7) & ")</td></tr>"
				K=K+1
				If K>=MaxPerPage Then Exit Do
				RS.MoveNext
			 Loop
		 End if
	End Sub
	'===========9-30========================
			Function GetUserPhoto(RS,totalPut,ChannelID)
		    Dim I,K,Url
			Dim PerLineNum:PerLineNum=4   '每行显示作品数
			  Do While Not RS.Eof 
              GetUserPhoto=GetUserPhoto & "<tr height=""20""> " &vbNewLine
			  
			  For K=1 To PerLineNum
			  If ChannelID=2 Then
			   Url="../space/?" & UserName & "/showphoto/" & RS(0)
			  Else
			   Url=KS.GetItemUrl(channelid,RS(2),RS(0),RS(5))
			  End If
              GetUserPhoto=GetUserPhoto & "  <td style=""BORDER-RIGHT: #efefef 1px dotted; BORDER-LEFT: #efefef 1px dotted; BORDER-BOTTOM: #efefef 1px dotted;text-align:center""><a href=""" & Url & """ target=""_blank""><img style='border:1px #efefef solid' width=120 height=80 src=""" & rs("photourl") & """ border=""0""></a><br><a href=""" & Url & """ target=""_blank"">" & KS.Gottopic(RS(1),15) & "</a></td>" & vbnewline
             RS.MoveNext
			    I = I + 1
				If rs.eof or I >= totalPut Then Exit For
			  Next
			   For K=K+1 To PerLineNum
            GetUserPhoto=GetUserPhoto & "   <td width=120 style=""BORDER-RIGHT: #efefef 1px dotted; BORDER-LEFT: #efefef 1px dotted; BORDER-BOTTOM: #efefef 1px dotted;text-align:center"">&nbsp;</td> " & vbcrlf
			   Next
            GetUserPhoto=GetUserPhoto & "   </tr> " & vbcrlf
				If I >= totalPut Then Exit Do
			 Loop

		End Function
		
		'通用分页
		Public Function ShowPage()
		         Dim I, PageStr
				 PageStr = ("<div class=""fenye""><table border='0' align='right'><tr><td><div class='showpage' style='height:28px'>")
					if (CurrPage>1) then pageStr=PageStr & "<a href=""../space/?" & userid & "/" &action & "/" & ID & "/" & CurrPage-1 & """ class=""prev"">上一页</a>"
				   if (CurrPage<>PageNum) then pageStr=PageStr & "<a href=""../space/?" & userid & "/" &action & "/" & ID & "/" & CurrPage+1 & """ class=""next"">下一页</a>"
				   pageStr=pageStr & "<a href=""../space/?" & userid & "/" &action & "/" & id & """ class=""prev"">首 页</a>"
				
				    If (totalPut Mod MaxPerPage) = 0 Then
						pagenum = totalPut \ MaxPerPage
					Else
						pagenum = totalPut \ MaxPerPage + 1
					End If
					Dim startpage,n,j
					 if (CurrPage>=7) then startpage=CurrPage-5
					 if PageNum-CurrPage<5 Then startpage=PageNum-10
					 If startpage<=0 Then startpage=1
					 n=0
					 For J=startpage To PageNum
						If J= CurrPage Then
						 PageStr=PageStr & " <a href=""#"" class=""curr""><font color=red>" & J &"</font></a>"
						Else
						 PageStr=PageStr & " <a class=""num"" href=""../space/?" & userid & "/" &action & "/" &id & "/" & J&""">" & J &"</a>"
						End If
						n=n+1 : if n>=10 then exit for
					 Next
					
					 PageStr=PageStr & " <a class=""next"" href=""../space/?" & userid & "/" &action & "/" &id & "/" & PageNum&""">末页</a>"
					 pageStr=PageStr & " <span>共" & totalPut & "条记录,分" & PageNum & "页</span></div></td></tr></table>"
				     PageStr = PageStr & "</div>"
			         ShowPage = PageStr
	     End Function
End Class
%>

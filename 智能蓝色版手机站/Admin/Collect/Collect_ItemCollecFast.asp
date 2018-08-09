<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%> 
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.CollectCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<%

Response.Buffer = True
Server.ScriptTimeout = 9999
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"

Dim KSCls
Set KSCls = New Collect_ItemCollectFast
KSCls.Kesion()
Set KSCls = Nothing

Class Collect_ItemCollectFast
        Private KS
		Private KMCObj
		Private ConnItem
		Private ItemNum, ListNum, PageNum, NewsSuccesNum, NewsFalseNum
		Private Rs, Sql, RsItem, SqlItem, FoundErr, ErrMsg, ItemEnd, ListEnd
		
		'项目变量
		Private ItemID, ItemName, ChannelID, strChannelDir, ClassID, SpecialID, LoginType, LoginUrl, LoginPostUrl, LoginUser, LoginPass, LoginFalse
		Private ListStr, LsString, LoString, ListPageType, LPsString, LPoString, ListPageStr1, ListPageStr2, ListPageID1, ListPageID2, ListPageStr3, HsString, HoString, HttpUrlType, HttpUrlStr
		Private TsString, ToString, CsString, CoString, DateType, DsString, DoString, AuthorType, AsString, AoString, AuthorStr, CopyFromType, FsString, FoString
		Private CopyFromStr, KeyType, KsString, KoString, KeyStr, NewsPageType, NPsString, NPoString, NewsPageStr, NewsPageEnd
		Private ItemCollecDate, PaginationType, MaxCharPerPage, ReadLevel, Stars, ReadPoint, Hits, UpDateType, UpDateTime, PicNews, Rolls, Comment, Recommend, Popular
		Private FnameType, TemplateID, Script_Iframe, Script_Object, Script_Script, Script_Div, Script_Class, Script_Span, Script_Img, Script_Font, Script_A, Script_Html, CollecListNum, CollecNewsNum, IntoBase, BeyondSavePic, CollecOrder, Verific, InputerType, Inputer, EditorType, Editor, ShowComment, Script_Table, Script_Tr, Script_Td,ThumbType,TbsString,TboString,CharsetCode,RepeatInto,Node,Tp_Lists,Tp_Listo,Tp_Srcs,Tp_Srco,Tp_Is,Tp_Io,Tp_AddressList_Code,Tp_PicUrls,tp_str,SaveFileName
		
		Private InfoPageArrayCode ,InfoPageArray,Testi,NewsNextPageStr,FieldNameList,FieldNameArr,i,Tp_Url,tp_intro
		
		Private Shop_BigPhotoBeginStr,Shop_BigPhotoEndStr,Shop_BigPhotoSrcBeginStr,Shop_BigPhotoSrcEndStr,Shop_UnitBeginStr,Shop_UnitEndStr,Shop_OriginPriceBeginStr,Shop_OriginPriceEndStr,Shop_PriceBeginStr,Shop_PriceEndStr,Shop_MarketPriceBeginStr,Shop_MarketPriceEndStr,Shop_MemberPriceBeginStr,Shop_MemberPriceEndStr,Shop_ProModelBeginStr,Shop_ProModelEndStr,Shop_ProSpecificatBeginStr,Shop_ProSpecificatEndStr,Shop_ProducerNameBeginStr,Shop_ProducerNameEndStr,Shop_TrademarkNameBeginStr,Shop_TrademarkNameEndStr

		
		'过滤变量
		Private Arr_Filters, FilterStr, Filteri
		
		'采集相关的变量
		Private ContentTemp, NewsPageNext, NewsPageNextCode, Arr_i, NewsUrl, NewsCode
		
		'文章保存变量
		Private ArticleID, Title, Content, Author, CopyFrom, Key, IncludePic, UploadFiles, DefaultPicUrl
		
		'其它变量
		Private LoginData, LoginResult, OrderTemp
		Private Arr_Item, CollecTest, Content_View, CollecNewsAll,Arr_Field
		Private StepID
		
		'历史记录
		Private Arr_Historys, His_Title, His_CollecDate, His_Result, His_Repeat, His_i
		
		'执行时间变量
		Private StartTime, OverTime
		
		'图片统计
		Private Arr_Images, ImagesNum, ImagesNumAll
		
		'列表
		Private ListUrl, ListCode, NewsArrayCode, NewsArray, ListArray, ListPageNext
		
		'小图片
		Private ThumbArrayCode,ThumbArray,ThumbUrl
		
		'安装路径
		Private strInstallDir, CacheTemp
		
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KMCObj=New CollectPublicCls
		  Set ConnItem = KS.ConnItem()
		End Sub
        Private Sub Class_Terminate()
		 Call KS.CloseConnItem()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KMCObj=Nothing
		End Sub
		Sub Kesion()
		
		strInstallDir =KS.Setting(3)
		
		'缓存路径
		CacheTemp = KS.SiteSN
		
		'数据初始化
		CollecListNum = 0
		CollecNewsNum = 0
		ArticleID = 0
		ItemNum = CLng(Trim(Request("ItemNum")))
		ListNum = CLng(Trim(Request("ListNum")))
		NewsSuccesNum = CLng(Trim(Request("NewsSuccesNum")))
		NewsFalseNum = CLng(Trim(Request("NewsFalseNum")))
		ImagesNumAll = CLng(Trim(Request("ImagesNumAll")))
		ListPageNext = Trim(Request("ListPageNext"))
		FoundErr = False
		ItemEnd = False
		ListEnd = False
		ErrMsg = ""
		
		Call SetCache
		
		If ItemEnd <> True Then
		   If (ItemNum - 1) > UBound(Arr_Item, 2) Then
			  ItemEnd = True
		   Else
			  Call SetItems
		   End If
		End If
		
		If ItemEnd <> True Then
		   If ListPageType = 0 Then
			  If ListNum = 1 Then
				 ListUrl = ListStr
			  Else
				 ListEnd = True
			  End If
		   ElseIf ListPageType = 1 Then
			  If ListNum = 1 Then
				 ListUrl = ListStr
			  Else
				 If ListPageNext = "" Or ListPageNext = "Error" Then
					ListEnd = True
				 Else
					ListPageNext = Replace(ListPageNext, "{$ID}", "&")
					ListUrl = ListPageNext
				 End If
			  End If
		   ElseIf ListPageType = 2 Then  '索引方式
					If ListNum = 1 Then
					 ListUrl = ListStr
					Else
						If ListPageID1 > ListPageID2 Then
						   If (ListPageID1 - ListNum + 1) < ListPageID2 Or (ListPageID1 - ListNum + 1) < 0 Then
							  ListEnd = True
						   Else
							  ListUrl = Replace(ListPageStr2, "{$ID}", CStr(ListPageID1 - ListNum + 1))
						   End If
						Else
						   If (ListPageID1 + ListNum - 1) > ListPageID2 Then
							  ListEnd = True
						   Else
							  ListUrl = Replace(ListPageStr2, "{$ID}", CStr(ListPageID1 + ListNum - 1))
						   End If
						End If
				   End If
		   ElseIf ListPageType = 3 Then
			  ListArray = Split(ListPageStr3, "|")
			  If (ListNum - 1) > UBound(ListArray) Then
				 ListEnd = True
			  Else
				 ListUrl = ListArray(ListNum - 1)
			  End If
		   End If
		   If ListNum > CollecListNum And CollecListNum <> 0 Then
			  ListEnd = True
		   End If
		End If
		
		If ItemEnd = True Then
		   ErrMsg = "<br>采集任务全部完成"
		   ErrMsg = ErrMsg & "<br>成功采集： " & NewsSuccesNum & "  篇,失败： " & NewsFalseNum & "  篇,图片：" & ImagesNumAll & "  张"
		   ChannelID=KS.G("ChannelID")
		   Call DelCache
			'定时任务,关闭
			if Session("taskf")="task" Then
			   KS.Echo "<script>setTimeout('window.close();',3000);</script>"
			End If
		Else
		   If ListEnd = True Then
			  ItemNum = ItemNum + 1
			  ListNum = 1
			  ErrMsg = "<br>" & ItemName & "  项目所有列表采集完成，正在整理数据请稍后..."
			  ErrMsg = ErrMsg & "<meta http-equiv=""refresh"" content=""3;url=Collect_ItemCollecFast.asp?ChannelID=" & Channelid & "&ItemNum=" & ItemNum & "&ListNum=" & ListNum & "&NewsSuccesNum=" & NewsSuccesNum & "&NewsFalseNum=" & NewsFalseNum & "&ImagesNumAll=" & ImagesNumAll & """>"
		   End If
		End If
		
		Call TopItem
		If ItemEnd = True Or ListEnd = True Then
		   If ItemEnd <> True Then
			  Call SetCache_His
		   End If
		   Call KMCObj.WriteCollectSucced(ErrMsg,ChannelID)
		Else
		   FoundErr = False
		   ErrMsg = ""
		   Call StartCollection
		   Call FootItem2
		End If
		Response.Flush
		End Sub
		'==================================================
		'过程名：StartCollection
		'作  用：开始采集
		'参  数：无
		'==================================================
		Sub StartCollection()
		Dim Rs
		'第一次采集时登录
		If LoginType = 1 And ListNum = 1 Then
		   LoginData = KMCObj.UrlEncoding(LoginUser & "&" & LoginPass)
		   LoginResult = KMCObj.PostHttpPage(LoginUrl, LoginPostUrl, LoginData)
		   If InStr(LoginResult, LoginFalse) > 0 Then
			  FoundErr = True
			  ErrMsg = ErrMsg & "<br><li>在登录网站时发生错误，请确保登录信息的正确性！</li>"
		   End If
		End If
		Set Rs = Server.CreateObject("Adodb.Recordset")
			 Rs.Open "Select top 1 ID From KS_Class Where ID='" & ClassID & "'", conn, 1, 1
			If Rs.EOF And Rs.BOF Then
				  FoundErr = True
				  ErrMsg = ErrMsg & "<br>系统检测到栏目ID[<font color=red>" & ClassID & "</font>]在主数据库中已删除，请修改项目属性的所属栏目后，再采集"
				  Call KMCObj.WriteCollectSuccedStart(ErrMsg)
			   Response.End
			End If
			Rs.Close
			Set Rs = Nothing
			
		If FoundErr <> True Then
		   ListCode = KMCObj.GetHttpPage(ListUrl,CharsetCode)

		   Call GetListPage
		   If ListCode = "Error" Then
			  FoundErr = True
			  ErrMsg = ErrMsg & "<br><li>在获取列表：" & ListUrl & "网页源码时发生错误！</li>"
		   Else
		      
			  ListCode = KMCObj.GetBody(ListCode, LsString, LoString, False, False)
			  If ListCode = "Error" Or ListCode = "" Then
				 FoundErr = True
				 ErrMsg = ErrMsg & "<br><li>在截取：" & ListUrl & "的列表时发生错误！</li>"
			  End If
		   End If
		End If
		
		
		If FoundErr <> True Then
		   NewsArrayCode = KMCObj.GetArray(ListCode, HsString, HoString, False, False)
		   ThumbArrayCode = KMCObj.GetArray(ListCode, tbsString, tboString, False, False)
		   If NewsArrayCode = "Error" Then
			  FoundErr = True
			  ErrMsg = ErrMsg & "<br><li>在分析：" & ListUrl & "列表时发生错误！</li>"
		   ElseIf ThumbType=1 and ThumbArrayCode="Error" Then
			  FoundErr = True
			  ErrMsg = ErrMsg & "<br><li>在分析：" & ListUrl & "列表缩略图发生错误！</li>"
		   Else
			  NewsArray = Split(NewsArrayCode, "$Array$")
			  ThumbArray=Split(ThumbArrayCode,"$Array$")
			  For Arr_i = 0 To UBound(NewsArray)
				 If HttpUrlType = 1 Then
					NewsArray(Arr_i) = Trim(Replace(HttpUrlStr, "{$ID}", NewsArray(Arr_i)))
				  
				 Else
					NewsArray(Arr_i) = Trim(KMCObj.DefiniteUrl(NewsArray(Arr_i), ListUrl))
				  
				 End If
				 NewsArray(Arr_i) = KMCObj.CheckUrl(NewsArray(Arr_i))
			  Next


			  
			  If CollecOrder = True Then
				 For Arr_i = 0 To Fix(UBound(NewsArray) / 2)
					OrderTemp = NewsArray(Arr_i)
					NewsArray(Arr_i) = NewsArray(UBound(NewsArray) - Arr_i)
					NewsArray(UBound(NewsArray) - Arr_i) = OrderTemp
				 Next
			  End If
		   End If
		End If
		
		If FoundErr <> True Then
		   Call TopItem2
		   CollecNewsAll = 0
		   For Arr_i = 0 To UBound(NewsArray)
			  If CollecNewsAll >= CollecNewsNum And CollecNewsNum <> 0 Then Exit For
			  CollecNewsAll = CollecNewsAll + 1
			  '变量初始化
			  UploadFiles = ""
			  DefaultPicUrl = ""
			  IncludePic = 0
			  ImagesNum = 0
			  NewsCode = ""
			  FoundErr = False
			  ErrMsg = ""
			  His_Repeat = False
			  NewsUrl = NewsArray(Arr_i)
			  If ThumbType=1 Then
			  ThumbUrl=ThumbArray(Arr_i)
			  End If
			  Title = ""
			  PageNum = 1
			  '………………………………………………
			  If Response.IsClientConnected Then
				 Response.Flush
			  Else
				 Response.End
			  End If
			  '………………………………………………
		
			  If CollecTest = False Then
				 His_Repeat = CheckRepeat(NewsUrl)
			  Else
				 His_Repeat = False
			  End If
			  If His_Repeat = True Then
				 FoundErr = True
			  End If
					  
			   
			  If FoundErr <> True Then
				 NewsCode = KMCObj.GetHttpPage(NewsUrl,CharsetCode)
				 If NewsCode = "Error" Then
					FoundErr = True
					ErrMsg = ErrMsg & "<br>在获取：" & NewsUrl & "源码时发生错误！"
					Title = "获取网页源码失败"
				 End If
			  End If
		
			  If FoundErr <> True Then
				 Title = KMCObj.GetBody(NewsCode, TsString, ToString, False, False)
				 If Title = "Error" Or Title = "" Then
					FoundErr = True
					ErrMsg = ErrMsg & "<br>在分析：" & NewsUrl & "的文章标题时发生错误"
					Title = "<br>标题分析错误"
				 End If
				 If FoundErr <> True Then
					Content = KMCObj.GetBody(NewsCode, CsString, CoString, False, False)
					If Content = "Error" Or Content = "" Then
					   FoundErr = True
					   ErrMsg = ErrMsg & "<br>在分析：" & NewsUrl & "的文章正文时发生错误"
					   Title = Title & "<br>正文分析错误"
					End If
				 End If
				 If FoundErr <> True Then
				 
				    
					
				If KS.C_S(ChannelID,6)="2" Then  '图片模型
				    Tp_AddressList_Code=KMCObj.GetBody(NewsCode, Tp_Lists, Tp_Listo, False, False)
					Tp_Url=KMCObj.GetBody(Tp_AddressList_Code, Tp_srcs, Tp_srco, False, False)
					If Tp_Url="Error" Then
					 Tp_Url=KMCObj.GetBody(NewsCode, Tp_srcs, Tp_srco, False, False)
					End If
					If Tp_Url="Error" Then
					   FoundErr = True
					   ErrMsg = ErrMsg & "<br>在分析：" & NewsUrl & "的单张图片地址发生错误"
					   Title = Title & "<br>单张图片地址分析错误"
					End If
					If Tp_is<>"" and Tp_Io<>"" Then
					 Tp_Intro=KMCObj.GetBody(Tp_AddressList_Code, Tp_is, Tp_io, False, False)
					End If
					If Tp_Intro="Error" Then
					  Tp_Intro=KMCObj.GetBody(NewsCode, Tp_is, Tp_io, False, False)
					End If
					If Tp_Intro="Error" Then Tp_Intro=""
					If CollecTest = False And BeyondSavePic = 1 Then '存图
						SaveFileName = Year(Now()) & Month(Now()) & Day(Now()) & KS.MakeRandom(10) & Mid(Tp_Url, InStrRev(Tp_Url, "."))
						Call KS.SaveBeyondFile(KS.GetUpFilesDir & "/" & SaveFileName,KMCObj.DefiniteUrl(Tp_Url, NewsUrl))
						Tp_Url=KS.Setting(2) & KS.GetUpFilesDir & "/" & SaveFileName
					End If
					tp_str=Replace(Tp_Intro,"|","")&"|" & Tp_Url & "|" & Tp_Url
					If UploadFiles="" Then
					  UploadFiles=Tp_Url
					Else
					  UploadFiles=UploadFiles & "|" & Tp_Url
					End If
					
					 '图片分页
					 If NewsPageType = 1 Then
						NewsPageNext = KMCObj.GetBody(NewsCode, NPsString, NPoString, False, False)
						If NewsPageNext = "Error" Then  '正文没有分页
						
				        Else
								 InfoPageArrayCode = KMCObj.GetArray(NewsPageNext, NewsPageStr, NewsPageEnd, False, False)
								 If InfoPageArrayCode = "Error" Then
									 FoundErr = True
									 ErrMsg = ErrMsg & "<br><li>在分析：图片分页时发生错误，请检查分页链接的开始代码和结束代码！</li>"
								  Else
									 InfoPageArray = Split(InfoPageArrayCode, "$Array$")
									 If IsArray(InfoPageArray) = True Then
										For Testi = 0 To UBound(InfoPageArray)
											'分页的正确地址
											InfoPageArray(Testi) = KMCObj.DefiniteUrl(InfoPageArray(Testi), NewsUrl) 
												  
											NewsPageNextCode = KMCObj.GetHttpPage(InfoPageArray(Testi),CharsetCode)
											'截取图片地址列表内容
											Tp_AddressList_Code=KMCObj.GetBody(NewsPageNextCode, Tp_Lists, Tp_Listo, False, False)
											If Tp_AddressList_Code = "Error" Then
													' Exit For
											Else
											 PageNum = PageNum + 1
                                             Tp_Url=KMCObj.GetBody(Tp_AddressList_Code, Tp_srcs, Tp_srco, False, False)
											 If Tp_Url="Error" Then
												Tp_Url=KMCObj.GetBody(NewsCode, Tp_srcs, Tp_srco, False, False)
											 End If
											 If Tp_Url="Error" Then
												 FoundErr = True
												 ErrMsg = ErrMsg & "<br>在分析：" & NewsUrl & "的单张图片地址发生错误"
												 Title = Title & "<br>单张图片地址分析错误"
											 End If
											 If Tp_is<>"" and Tp_Io<>"" Then
												 Tp_Intro=KMCObj.GetBody(Tp_AddressList_Code, Tp_is, Tp_io, False, False)
											 End If
											 If Tp_Intro="Error" Then
												  Tp_Intro=KMCObj.GetBody(NewsCode, Tp_is, Tp_io, False, False)
											 End If
											 If Tp_Intro="Error" Then Tp_Intro=""
													
											If CollecTest = False And BeyondSavePic = 1 Then '存图
												SaveFileName = Year(Now()) & Month(Now()) & Day(Now()) & KS.MakeRandom(10) & Mid(Tp_Url, InStrRev(Tp_Url, "."))
												Call KS.SaveBeyondFile(KS.GetUpFilesDir & "/" & SaveFileName,KMCObj.DefiniteUrl(Tp_Url, NewsUrl))
												Tp_Url=KS.Setting(2) & KS.GetUpFilesDir & "/" & SaveFileName
											End If
											tp_str=tp_str & "|||" & Replace(Tp_Intro,"|","")&"|" & Tp_Url & "|" & Tp_Url
										 	If UploadFiles="" Then
												  UploadFiles=Tp_Url
											Else
												  UploadFiles=UploadFiles & "|" & Tp_Url
											End If													 
													 
													 
										   End If 
										Next
										 
									  Else
										FoundErr = True
										ErrMsg = ErrMsg & "<br><li>在分析：" & NewsUrl & "图片分页列表时发生错误！</li>"
									 End If
								  End If
						End if
			         End If
					
			  ElseIf KS.C_S(ChannelID,6)=5 Then  '商城系统		
				'采集大图
				Tp_AddressList_Code=KMCObj.GetBody(NewsCode, Shop_BigPhotoBeginStr, Shop_BigPhotoEndStr, False, False)
				Tp_Url=KMCObj.GetBody(Tp_AddressList_Code, Shop_BigPhotoSrcBeginStr, Shop_BigPhotoSrcEndStr, False, False)
				If Tp_Url<>"Error" And Tp_Url<>"" Then
					Tp_Url=KMCObj.DefiniteUrl(Tp_Url, NewsUrl)
					If CollecTest = False And BeyondSavePic = 1 Then '存图
							SaveFileName = Year(Now()) & Month(Now()) & Day(Now()) & KS.MakeRandom(10) & Mid(Tp_Url, InStrRev(Tp_Url, "."))
							Call KS.SaveBeyondFile(KS.GetUpFilesDir & "/" & SaveFileName,KMCObj.DefiniteUrl(Tp_Url, NewsUrl))
							Tp_Url=KS.Setting(2) & KS.GetUpFilesDir & "/" & SaveFileName
					End If
				End If
			  Else
				    '源代码中获取分页URL  
					 '正文分页
					 If NewsPageType = 1 Then
						NewsPageNext = KMCObj.GetBody(NewsCode, NPsString, NPoString, False, False)
						'response.write NewsPageNext& "<br><hr color=red>"
						If NewsPageNext = "Error" Then  '正文没有分页
						
				        Else
								 InfoPageArrayCode = KMCObj.GetArray(NewsPageNext, NewsPageStr, NewsPageEnd, False, False)
								 If InfoPageArrayCode = "Error" Then
									 'FoundErr = True        '为防止部分文章没有分页出错,屏蔽掉这句
									 'ErrMsg = ErrMsg & "<br><li>在分析：新闻正文分页时发生错误，请检查分页链接的开始代码和结束代码！</li>"
								  Else
										InfoPageArray = Split(InfoPageArrayCode, "$Array$")
										 If IsArray(InfoPageArray) = True Then
											For Testi = 0 To UBound(InfoPageArray)
											      '分页的正确地址
												  InfoPageArray(Testi) = KMCObj.DefiniteUrl(InfoPageArray(Testi), NewsUrl) 
												  
												  NewsPageNextCode = KMCObj.GetHttpPage(InfoPageArray(Testi),CharsetCode)
												  '截取内容
												  ContentTemp=KMCObj.GetBody(NewsPageNextCode, CsString, CoString, False, False)
												  
												  NewsNextPageStr = KMCObj.GetBody(NewsPageNextCode, NPsString, NPoString, true, true)
												 ' response.write InfoPageArray(Testi) &"<br>" &  NewsNextPageStr & "<hr>"
												  if NewsNextPageStr="Error" Then  '载取分页字符串没成功时，改变结束标记重新载取
												   NewsNextPageStr=KMCObj.GetBody(ContentTemp, NPsString, CoString,true, true)
												  End IF
												  
												  IF NewsNextPageStr<>"Error" Then 
												   ContentTemp=Replace(ContentTemp,NewsNextPageStr,"")         '替换分页部分
												  End IF
												  
												  If ContentTemp = "Error" Then
													' Exit For
												  Else
													PageNum = PageNum + 1
													 IF PaginationType=0 Then      ' 不分页
													  Content=Content&ContentTemp
													 ElseIF PaginationType=1 Then  '自动分页
													   Content=Content&ContentTemp
													 ElseIf PaginationType=2 Then  '原文分页方式
													  Content = Content & "[NextPage]" & ContentTemp
													 End IF
												  End If 
											Next
											 
										 Else
											FoundErr = True
											ErrMsg = ErrMsg & "<br><li>在分析：" & NewsUrl & "新闻列表时发生错误！</li>"
										 End If
								  End If
						End if
						Content=Replace(Content,NewsPageNext,"")
			         End If
				 End If
					
					
						IF PaginationType=1 Then             '调用自动分页函数
						 Content=KMCObj.SplitNewsPage(Content,MaxCharPerPage)
						End IF
				 
					'过滤
					Call Filters
					Title = KMCObj.FpHtmlEnCode(Title)
					Call FilterScript
					Content = KMCObj.UBBCode(Content, strInstallDir, strChannelDir)
				 End If
			  End If
		 
		
		
			  If FoundErr <> True Then
				 '时间
				 If UpDateType = 0 Then
					UpDateTime = Now()
				 ElseIf UpDateType = 1 Then
					If DateType = 0 Then
					   UpDateTime = Now()
					Else
					   UpDateTime = KMCObj.GetBody(NewsCode, DsString, DoString, False, False)
					   If Not IsDate(UpDateTime) Then
					   UpDateTime = Replace(Replace(Replace(Trim(Replace(UpDateTime, "&nbsp;", " ")),"年","-"),"月","-"),"日","-")
					   End If
					   If IsDate(UpDateTime) = True Then
						  UpDateTime = CDate(UpDateTime)
					   Else
						  UpDateTime = Now()
					   End If
					End If
				 ElseIf UpDateType = 2 Then
				 Else
					UpDateTime = Now()
				 End If
						  
				 '作者
				 If AuthorType = 1 Then
					Author = KMCObj.GetBody(NewsCode, AsString, AoString, False, False)
				 ElseIf AuthorType = 2 Then
					Author = AuthorStr
				 Else
					Author = "佚名"
				 End If
				 Author = KMCObj.FpHtmlEnCode(Author)
				 If Author = "" Or Author = "Error" Then
					Author = "佚名"
				 Else
					If Len(Author) > 255 Then
					   Author = Left(Author, 255)
					End If
				 End If
				   
				 '来源
				 If CopyFromType = 1 Then
					CopyFrom = KMCObj.GetBody(NewsCode, FsString, FoString, False, False)
				 ElseIf CopyFromType = 2 Then
					CopyFrom = CopyFromStr
				 Else
					CopyFrom = "不详"
				 End If
				 
				 CopyFrom = KMCObj.FpHtmlEnCode(KS.ScriptHtml(CopyFrom, "A", 3))
				 If CopyFrom = "" Or CopyFrom = "Error" Then
						CopyFrom = "不详"
				 Else
					If Len(CopyFrom) > 255 Then
					   CopyFrom = Left(CopyFrom, 255)
					End If
				 End If
		
				 '关键字
				 If KeyType = 0 Then
					Key = Title
					Key = KMCObj.CreateKeyWord(Key, 2)
				 ElseIf KeyType = 1 Then
					Key = KMCObj.GetBody(NewsCode, KsString, KoString, False, False)
					Key = KMCObj.FpHtmlEnCode(Key)
					'Key = KMCObj.CreateKeyWord(Key, 2)
				 ElseIf KeyType = 2 Then
					Key = KeyStr
					Key = KMCObj.FpHtmlEnCode(Key)
					If Len(Key) > 253 Then
					   Key = "," & Left(Key, 253) & ","
					Else
					   Key = "," & Key & ","
					End If
				 End If
				 If Key = "" Or Key = "Error" Then
					Key = ""
				 End If

				 '转换图片相对地址为绝对地址/保存
				 If CollecTest = False And BeyondSavePic = 1 Then
				   Content = KMCObj.ReplaceSaveRemoteFile(UploadFiles, Content, strInstallDir, strChannelDir, True, NewsUrl)
				   if ThumbUrl<>"" then
				    '缩略图
					SaveFileName = Year(Now()) & Month(Now()) & Day(Now()) & KS.MakeRandom(10) & Mid(ThumbUrl, InStrRev(ThumbUrl, "."))
					Call KS.SaveBeyondFile(KS.GetUpFilesDir & "/" & SaveFileName,KMCObj.DefiniteUrl(ThumbUrl, NewsUrl))
					ThumbUrl=KS.Setting(2) & KS.GetUpFilesDir & "/" & SaveFileName
				   end if
				 Else
				   Content = KMCObj.ReplaceSaveRemoteFile(UploadFiles, Content, strInstallDir, strChannelDir, False, NewsUrl)
				 End If
				 '转换swf文件地址
				 Content = KMCObj.ReplaceSwfFile(Content, NewsUrl)
		  
				 '图片统计、文章图片属性设置
				 If UploadFiles <> "" Then
					If InStr(UploadFiles, "|") > 0 Then
					   Arr_Images = Split(UploadFiles, "|")
					   ImagesNum = UBound(Arr_Images) + 1
					   DefaultPicUrl = Arr_Images(0)
					Else
					   ImagesNum = 1
					   DefaultPicUrl = UploadFiles
					End If
		
					If BeyondSavePic <> 1 Then
					   UploadFiles = ""
					End If
				 Else
					ImagesNum = 0
					DefaultPicUrl = ""
					IncludePic = 0
				 End If
				 ImagesNumAll = ImagesNumAll + ImagesNum
			  End If
		
			  If FoundErr <> True Then
				 If CollecTest = False Then
					Call SaveArticle
					SqlItem = "INSERT INTO KS_History(ItemID,ChannelID,ClassID,SpecialID,ArticleID,Title,CollecDate,NewsUrl,Result) VALUES ('" & ItemID & "','" & ChannelID & "','" & ClassID & "','" & SpecialID & "','" & ArticleID & "','" &  left(Title,255) & "','" & Now() & "','" &  left(NewsUrl,255) & "',True)"
					if Title<>"" then ConnItem.Execute (SqlItem)
					Content = Replace(Content, "[InstallDir_ChannelDir]", strInstallDir & strChannelDir & "/")
				 End If
				 NewsSuccesNum = NewsSuccesNum + 1
				 ErrMsg = ErrMsg & "No:<font color=red>" & NewsSuccesNum + NewsFalseNum & "</font><br>"
				 ErrMsg = ErrMsg & "数据标题："
				 ErrMsg = ErrMsg & "<font color=red>" & Title & "</font><br>"
				 
				 If DefaultPicUrl<>"" Then
				 ErrMsg = ErrMsg & "缩略图片：" & DefaultPicUrl &"<br>"
				 End If
				 
				 ErrMsg = ErrMsg & "更新时间：" & UpDateTime & "<br>"
				 ErrMsg = ErrMsg & "采集页面：<a href=" & NewsUrl & " target=_blank>" & NewsUrl & "</a><br>"
				 ErrMsg = ErrMsg & "其它信息：分页--" & PageNum & " 页，图片--" & ImagesNum & " 张<br>"
				 ErrMsg = ErrMsg & "关键字Tags：" & Key & ""
				Call InnerJS(Arr_I,UBound(NewsArray)+1,ErrMsg)
			  Else
				 NewsFalseNum = NewsFalseNum + 1
				 If His_Repeat = True Then
					ErrMsg = ErrMsg & "No:<font color=red>" & NewsSuccesNum + NewsFalseNum & "</font><br>"
					ErrMsg = ErrMsg & "目标数据：<font color=red>"
					If His_Result = True Then
					   ErrMsg = ErrMsg & His_Title
					Else
					   ErrMsg = ErrMsg & NewsUrl
					End If
					ErrMsg = ErrMsg & "</font> 的记录已存在，不给予采集。<br>"
					ErrMsg = ErrMsg & "采集时间：" & His_CollecDate & "<br>"
					ErrMsg = ErrMsg & "采集结果："
					If His_Result = False Then
					   ErrMsg = ErrMsg & "失败"
					   ErrMsg = ErrMsg & "<br>失败原因：" & Title
					Else
					   ErrMsg = ErrMsg & "成功"
					End If
					ErrMsg = ErrMsg & "<br>提示信息：如想再次采集，请先将该历史记录<font color=red>删除</font><br>"
				 End If
				 If CollecTest = False And His_Repeat = False Then
					SqlItem = "INSERT INTO KS_History(ItemID,ChannelID,ClassID,SpecialID,Title,CollecDate,NewsUrl,Result) VALUES ('" & ItemID & "','" & ChannelID & "','" & ClassID & "','" & SpecialID & "','" & left(replace(Title,"'",""),255) & "','" & Now() & "','" & left(NewsUrl,255) & "',False)"
					if Title<>"" then ConnItem.Execute (SqlItem)
				 End If
			     Call ShowMsg(ErrMsg)
			     Response.Flush  '刷新
			  End If
		   Next
		Else
		   errmsg="<b>&nbsp;&nbsp;&nbsp;&nbsp;发现错误：</b>" & errmsg & "<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color=red>3秒后自动转向采集下一页！</font>"
		   Call ShowMsg(ErrMsg)
		End If

		End Sub
		
		'==================================================
		'过程名：TopItem
		'作  用：显示导航信息
		'参  数：无
		'==================================================
		Sub TopItem()
		Response.Write "<html>"
		Response.Write "<head>"
		Response.Write "<title>采集系统</title>"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		Response.Write "<link rel=""stylesheet"" type=""text/css"" href=""../Include/Admin_Style.css"">"
		Response.Write "</head>"
		Response.Write "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
		Response.Write "<div class=""topdashed sort"">采集系统采集管理</div>"
		End Sub
		
		
		Sub TopItem2()
		
		Response.Write "<br>"
		Response.Write "<table width=""100%"" height=""20"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		Response.Write "  <tr>"
		 Response.Write "   <td width=""50%;"" align=""right""><span style=""color:red;""><strong><font id=""CollectEndArea"">系统正在采集</font></strong></span></td>"
		Response.Write "    <td width=""50%;"" valign=""top"">&nbsp;<span style=""color:red;""><strong><font id=""ShowInfoArea"">&nbsp;</font></strong></span></td>"
		Response.Write "  </tr>"
		Response.Write "</table>"
		Response.Write "<table width=""98%"" border=""0"" style=""text=align:center"" cellpadding=""0"" cellspacing=""1"">"
		 Response.Write "   <tr>"
		 Response.Write "     <td height=""45"" colspan=""2"" aling=""left"">本次运行：" & UBound(Arr_Item, 2) + 1 & " 个项目,正在采集第 <font color=red>" & ItemNum & "</font> 个项目  <font color=red>" & ItemName & "</font>  的第   <font color=red>" & ListNum & "</font> 页列表,该列表待采集数据  <font color=red>" & UBound(NewsArray) + 1 & "</font> 条。"
			  If CollecNewsNum <> 0 Then Response.Write "限制 <font color=red>" & CollecNewsNum & "</font> 篇。"
		 Response.Write "     <br>采集统计：成功采集--" & NewsSuccesNum & "  条数据，失败--" & NewsFalseNum & "  条数据，图片--" & ImagesNumAll & "　张。<a href=""Collect_Main.asp?ChannelID=" & ChannelID &""">停止采集</a>"
		 Response.Write "     </td>"
		 Response.Write "   </tr>"
		Response.Write "</table>"
		Response.Write "<script language=""JavaScript"">"
		Response.Write "var ForwardShow=true;"
		Response.Write "function ShowPromptInfo()"
		Response.Write "{"
		Response.Write "    var TempStr=document.all.ShowInfoArea.innerText;"
		Response.Write "    if (ForwardShow==true)"
		Response.Write "    {"
		Response.Write "        if (TempStr.length>4) ForwardShow=false;"
		Response.Write "        document.all.ShowInfoArea.innerText=TempStr+'.';"
		Response.Write "    }"
		Response.Write "    else"
		 Response.Write "   {"
		 Response.Write "       if (TempStr.length==2) ForwardShow=true;"
		 Response.Write "       document.all.ShowInfoArea.innerText=TempStr.substr(0,TempStr.length-1);"
		 Response.Write "   }"
		Response.Write "}"
		Response.Write "window.setInterval('ShowPromptInfo()',200);</script>"
		
		    Response.Write "<div id='tips'>"
			Response.Write "<br><br><br><table id=""BarShowArea"" width=""400"" border=""0"" align=""center"" cellspacing=""1"" cellpadding=""1"">"
			Response.Write "<tr> "
			Response.Write "<td bgcolor=000000>"
			Response.Write " <table width=""400"" border=""0"" cellspacing=""0"" cellpadding=""1"">"
			Response.Write "<tr> "
			Response.Write "<td bgcolor=ffffff height=9><img src=""../images/114_r2_c2.jpg"" width=0 height=10 id=img2 name=img2 align=absmiddle></td></tr></table>"
			Response.Write "</td></tr></table>"
			Response.Write "<table width=""550"" border=""0"" align=""center"" cellspacing=""1"" cellpadding=""1""><tr> "
			Response.Write "<td align=center> <span id=txt2 name=txt2 style=""font-size:9pt"">0</span><span id=txt4 style=""font-size:9pt"">%</span></td></tr> "
			Response.Write "<tr><td align=center><span id=txt3 name=txt3 style=""font-size:9pt"">0</span></td></tr>"
			Response.Write "</table>"
			Response.Write "<table align=""center"" style=""margin-top:30px;border: double #E7E7E7;overflow: auto; width:80%""  border=""0"" cellpadding=""0"" cellspacing=""0"">"
			Response.Write " <tr>"
			Response.Write "   <td height=""100"" id=""fsohtml"">系统正在初始化数据...</td>"
			Response.Write "   </tr>"
			Response.Write "</table>"
			Response.Write "</div>"
		StartTime = Timer()
		End Sub
		
		Sub InnerJS(NowNum,TotalNum,msg)
		  msg=Replace(Replace(Replace(msg, Chr(13) & Chr(10), ""),"'","\'"),"""","\""")
		  With Response
				.Write "<script>"
				.Write "fsohtml.innerHTML='" & msg & "';" & vbCrLf
				.Write "img2.width=" & Fix((NowNum / TotalNum) * 400) & ";" & vbCrLf
				.Write "txt2.innerHTML=""采集进度:" & FormatNumber(NowNum / TotalNum * 100, 2, -1) & """;" & vbCrLf
				.Write "txt3.innerHTML=""正在采集第 <font color=red>" & ItemNum & "</font> 个项目的第   <font color=red>" & ListNum & "</font> 页列表,本页需要采集 <font color=red><b>" & TotalNum & "</b></font> 条数据,<font color=red><b>在此过程中请勿刷新此页面！！！</b></font> 系统正在采集第 <font color=red><b>" & NowNum & "</b></font> 条"";" & vbCrLf
				.Write "img2.title=""(" & NowNum & ")"";</script>" & vbCrLf
				.Flush
		  End With

		End Sub

		'==================================================
		'过程名：FootItem2
		'作  用：显示该列表采集时间等信息
		'参  数：无
		'==================================================
		Sub FootItem2()
		   OverTime = Timer()
		   With Response
		        If CollecTest = False Then
				.Write "<meta http-equiv=""refresh"" content=""3;url=Collect_ItemCollecFast.asp?ItemNum=" & ItemNum & "&ListNum=" & ListNum + 1 & "&NewsSuccesNum=" & NewsSuccesNum & "&NewsFalseNum=" & NewsFalseNum & "&ImagesNumAll=" & ImagesNumAll & "&ListPageNext=" & ListPageNext & """>"
				End If
				if founderr<>true then
					.Write "<script>"
					If CollecTest = False Then
					.Write "fsohtml.innerHTML='执行时间：" & CStr(FormatNumber((OverTime - StartTime) * 1000, 2)) & " 毫秒,数据整理中，3秒后继续......3秒后如果还没反应请点击 <a href=""Collect_ItemCollecFast.asp?ItemNum=" & ItemNum & "&ListNum=" & ListNum + 1 & "&NewsSuccesNum=" & NewsSuccesNum & "&NewsFalseNum=" & NewsFalseNum & "&ImagesNumAll=" & ImagesNumAll & "&ListPageNext=" & ListPageNext & """><font color=red>这里</font></a> 继续<br>';" & vbCrLf
					else
					.Write "fsohtml.innerHTML='执行时间：" & CStr(FormatNumber((OverTime - StartTime) * 1000, 2)) & " 毫秒';" & vbCrLf
					end if
					.Write "img2.width=400;" & vbCrLf
					.Write "txt2.innerHTML=""采集进度:100"";" & vbCrLf
					.Write "txt3.innerHTML="""";" & vbCrLf
					.Write "img2.title='';</script>" & vbCrLf
				end if
				.Flush
		  End With
		End Sub
		
		'==================================================
		'过程名：ShowMsg
		'作  用：显示信息
		'参  数：无
		'==================================================
		Sub ShowMsg(Msg)
		   Dim strTemp
		   if founderr<>true then
		   strTemp = "<script>document.getElementById('tips').style.display='none';</script>"
		   end if
		   strtemp = strTemp & "<table width=""90%"" border=""0"" bgcolor=""#efefef"" align=""center"" cellpadding=""2"" cellspacing=""1"">"
		   strTemp = strTemp & "   <tr>"
		   strTemp = strTemp & "      <td height=""22"" colspan=""2"" bgcolor=""#ffffff"" align=""left"">"
		   strTemp = strTemp & Msg
		   strTemp = strTemp & "      </td>"
		   strTemp = strTemp & "   </tr><br>"
		   strTemp = strTemp & "</table>"
		   Response.Write strTemp
		End Sub
		
		'==================================================
		'过程名：SetCache
		'作  用：存取缓存
		'参  数：无
		'==================================================
		Sub SetCache()
		   Dim myCache
		   Set myCache = New ClsCache
		
		   '项目信息
		   myCache.name = CacheTemp & "items"
		   If myCache.valid Then
			  Arr_Item = myCache.value
		   Else
			  ItemEnd = True
		   End If
		   
		   '自定义字段
		   myCache.name = CacheTemp & "field"
		   If myCache.valid Then
			  Arr_Field = myCache.value
		   End If
		
		   '过滤信息
		   myCache.name = CacheTemp & "filters"
		   If myCache.valid Then
			  Arr_Filters = myCache.value
		   End If
		
		   '历史记录
		   myCache.name = CacheTemp & "Historys"
		   If myCache.valid Then
			  Arr_Historys = myCache.value
		   End If
		
		   '其它信息
		   myCache.name = CacheTemp & "collectest"
		   If myCache.valid Then
			  CollecTest = myCache.value
		   Else
			  CollecTest = False
		   End If
		   myCache.name = CacheTemp & "contentview"
		   If myCache.valid Then
			  Content_View = myCache.value
		   Else
			  Content_View = False
		   End If
		
		   Set myCache = Nothing
		End Sub
		
		Sub DelCache()
		   Dim myCache
		   Set myCache = New ClsCache
		   myCache.name = CacheTemp & "items"
		   Call myCache.clean
		   myCache.name = CacheTemp & "filters"
		   Call myCache.clean
		   myCache.name = CacheTemp & "Historys"
		   Call myCache.clean
		   myCache.name = CacheTemp & "collectest"
		   Call myCache.clean
		   myCache.name = CacheTemp & "contentview"
		   Call myCache.clean
		   Set myCache = Nothing
		End Sub
		
		'==================================================
		'过程名：SetItems
		'作  用：获取项目信息
		'参  数：无
		'==================================================
		Sub SetItems()
			  Dim ItemNumTemp
			  ItemNumTemp = ItemNum - 1
			  ItemID = Arr_Item(0, ItemNumTemp)
			  ItemName = Arr_Item(1, ItemNumTemp)
			  ChannelID = Arr_Item(2, ItemNumTemp)     '频道ID
			  strChannelDir = Arr_Item(3, ItemNumTemp) '频道目录
			  ClassID = Arr_Item(4, ItemNumTemp)         '栏目
			  SpecialID = Arr_Item(5, ItemNumTemp)     '专题
			  LoginType = Arr_Item(9, ItemNumTemp)
			  LoginUrl = Arr_Item(10, ItemNumTemp)       '登录
			  LoginPostUrl = Arr_Item(11, ItemNumTemp)
			  LoginUser = Arr_Item(12, ItemNumTemp)
			  LoginPass = Arr_Item(13, ItemNumTemp)
			  LoginFalse = Arr_Item(14, ItemNumTemp)
			  ListStr = Arr_Item(15, ItemNumTemp)         '列表地址
			  LsString = Arr_Item(16, ItemNumTemp)        '列表
			  LoString = Arr_Item(17, ItemNumTemp)
			  ListPageType = Arr_Item(18, ItemNumTemp)
			  LPsString = Arr_Item(19, ItemNumTemp)
			  LPoString = Arr_Item(20, ItemNumTemp)
			  ListPageStr1 = Arr_Item(21, ItemNumTemp)
			  ListPageStr2 = Arr_Item(22, ItemNumTemp)
			  ListPageID1 = Arr_Item(23, ItemNumTemp)
			  ListPageID2 = Arr_Item(24, ItemNumTemp)
			  ListPageStr3 = Arr_Item(25, ItemNumTemp)
			  HsString = Arr_Item(26, ItemNumTemp)
			  HoString = Arr_Item(27, ItemNumTemp)
			  HttpUrlType = Arr_Item(28, ItemNumTemp)
			  HttpUrlStr = Arr_Item(29, ItemNumTemp)
		
			  TsString = Arr_Item(30, ItemNumTemp)       '标题
			  ToString = Arr_Item(31, ItemNumTemp)
			  CsString = Arr_Item(32, ItemNumTemp)       '正文
			  CoString = Arr_Item(33, ItemNumTemp)
			  DateType = Arr_Item(34, ItemNumTemp)   '作者
			  DsString = Arr_Item(35, ItemNumTemp)
			  DoString = Arr_Item(36, ItemNumTemp)
			  AuthorType = Arr_Item(37, ItemNumTemp)   '作者
			  AsString = Arr_Item(38, ItemNumTemp)
			  AoString = Arr_Item(39, ItemNumTemp)
			  AuthorStr = Arr_Item(40, ItemNumTemp)
			  CopyFromType = Arr_Item(41, ItemNumTemp) '来源
			  FsString = Arr_Item(42, ItemNumTemp)
			  FoString = Arr_Item(43, ItemNumTemp)
			  CopyFromStr = Arr_Item(44, ItemNumTemp)
			  KeyType = Arr_Item(45, ItemNumTemp)         '关键词
			  KsString = Arr_Item(46, ItemNumTemp)
			  KoString = Arr_Item(47, ItemNumTemp)
			  KeyStr = Arr_Item(48, ItemNumTemp)
			  NewsPageType = Arr_Item(49, ItemNumTemp)         '文章分页
			  NPsString = Arr_Item(50, ItemNumTemp)            '文章分页代码开始
			  NPoString = Arr_Item(51, ItemNumTemp)            '文章分页代码结束
			  NewsPageStr = Arr_Item(52, ItemNumTemp)          '文章分页链接的开始标记
			  NewsPageEnd = Arr_Item(53, ItemNumTemp)          '文章分页链接的结束标记
			  PaginationType = Arr_Item(55, ItemNumTemp)
			  MaxCharPerPage = Arr_Item(56, ItemNumTemp)
			  ReadLevel = Arr_Item(57, ItemNumTemp)
			  Stars = Arr_Item(58, ItemNumTemp)
			  ReadPoint = Arr_Item(59, ItemNumTemp)
			  Hits = Arr_Item(60, ItemNumTemp)
			  UpDateType = Arr_Item(61, ItemNumTemp)
			  UpDateTime = Arr_Item(62, ItemNumTemp)
			  PicNews = Arr_Item(63, ItemNumTemp)
			  Rolls = Arr_Item(64, ItemNumTemp)
			  Comment = Arr_Item(65, ItemNumTemp)
			  Recommend = Arr_Item(66, ItemNumTemp)
			  Popular = Arr_Item(67, ItemNumTemp)
			  FnameType = Arr_Item(68, ItemNumTemp)          '生成的扩展名
			  TemplateID = Arr_Item(69, ItemNumTemp)         '生成的模板
			  Script_Iframe = Arr_Item(70, ItemNumTemp)
			  Script_Object = Arr_Item(71, ItemNumTemp)
			  Script_Script = Arr_Item(72, ItemNumTemp)
			  Script_Div = Arr_Item(73, ItemNumTemp)
			  Script_Class = Arr_Item(74, ItemNumTemp)
			  Script_Span = Arr_Item(75, ItemNumTemp)
			  Script_Img = Arr_Item(76, ItemNumTemp)
			  Script_Font = Arr_Item(77, ItemNumTemp)
			  Script_A = Arr_Item(78, ItemNumTemp)
			  Script_Html = Arr_Item(79, ItemNumTemp)
			  CollecListNum = Arr_Item(80, ItemNumTemp)
			  CollecNewsNum = Arr_Item(81, ItemNumTemp)
			  IntoBase = Arr_Item(82, ItemNumTemp)
			  BeyondSavePic = Arr_Item(83, ItemNumTemp)
			  CollecOrder = Arr_Item(84, ItemNumTemp)
			  Verific = Arr_Item(85, ItemNumTemp)
			  InputerType = Arr_Item(86, ItemNumTemp)
			  Inputer = Arr_Item(87, ItemNumTemp)
			  EditorType = Arr_Item(88, ItemNumTemp)
			  Editor = Arr_Item(89, ItemNumTemp)
			  ShowComment = Arr_Item(90, ItemNumTemp)
			  Script_Table = Arr_Item(91, ItemNumTemp)
			  Script_Tr = Arr_Item(92, ItemNumTemp)
			  Script_Td = Arr_Item(93, ItemNumTemp)
			  ThumbType = Arr_Item(94, ItemNumTemp)
			  TbsString = Arr_Item(95, ItemNumTemp)
			  TboString = Arr_Item(96, ItemNumTemp)
		      CharsetCode =Arr_Item(97,ItemNumTemp) '编码
			  RepeatInto=Arr_Item(98,ItemNumTemp)
			  
			  If KS.C_S(ChannelID,6)=2 Then  '图片模型
			   If IsObject(Application("CollectFieldRules")) Then
				   Set Node=Application("CollectFieldRules").DocumentElement.SelectSingleNode("row[@fieldname='Tp_List'&&@itemid=" & ItemID& "&&@channelid="&channelid&"]/@beginstr")
				   if not Node Is Nothing Then
					Tp_lists=Node.Text
				   End If
				   Set Node=Application("CollectFieldRules").DocumentElement.SelectSingleNode("row[@fieldname='Tp_List']/@endstr")
				   if not Node Is Nothing Then
					Tp_listo=Node.Text
				   End If
				   Set Node=Application("CollectFieldRules").DocumentElement.SelectSingleNode("row[@fieldname='Tp_Src']/@beginstr")
				   if not Node Is Nothing Then
					Tp_Srcs=Node.Text
				   End If
				   Set Node=Application("CollectFieldRules").DocumentElement.SelectSingleNode("row[@fieldname='Tp_Src']/@endstr")
				   if not Node Is Nothing Then
					Tp_Srco=Node.Text
				   End If
				   Set Node=Application("CollectFieldRules").DocumentElement.SelectSingleNode("row[@fieldname='Tp_I']/@beginstr")
				   if not Node Is Nothing Then
					Tp_Is=Node.Text
				   End If
				   Set Node=Application("CollectFieldRules").DocumentElement.SelectSingleNode("row[@fieldname='Tp_I']/@endstr")
				   if not Node Is Nothing Then
					Tp_Io=Node.Text
				   End If
		         End If
			  ElseIf KS.C_S(ChannelID,6)=5 Then
			     FieldNameList="Shop_BigPhoto,Shop_BigPhotoSrc,Shop_Unit,Shop_OriginPrice,Shop_Price,Shop_MarketPrice,Shop_MemberPrice,Shop_ProModel,Shop_ProSpecificat,Shop_ProducerName,Shop_TrademarkName"
				 FieldNameArr=Split(FieldNameList,",")
				  If IsObject(Application("CollectFieldRules")) Then
				    For I=0 To Ubound(FieldNameArr)
					   Set Node=Application("CollectFieldRules").DocumentElement.SelectSingleNode("row[@fieldname='" & FieldNameArr(i) & "']/@beginstr")
					   if not Node Is Nothing Then
						Execute(FieldNameArr(i)&"beginstr=Node.Text")
					   End If
					   Set Node=Application("CollectFieldRules").DocumentElement.SelectSingleNode("row[@fieldname='" & FieldNameArr(i) & "']/@endstr")
					   if not Node Is Nothing Then
						Execute(FieldNameArr(i)&"endstr=Node.Text")
					   End If
					
					Next
				  End If
			  End If
			  
			  If InputerType = 1 Then
				 Inputer = KMCObj.FpHtmlEnCode(Inputer)
			  Else
				 Inputer = KS.C("AdminName")
			  End If
			  If EditorType = 1 Then
				 Editor = KMCObj.FpHtmlEnCode(Editor)
			  Else
				 Editor = KS.C("AdminName")
			  End If
		End Sub
		
		'==================================================
		'过程名：GetListPage
		'作  用：获取列表下一页
		'参  数：无
		'==================================================
		Sub GetListPage()
		   If ListPageType = 1 Then
			  ListPageNext = KMCObj.GetPage(ListCode, LPsString, LPoString, False, False)
			 ' ListPageNext = KMCObj.FpHtmlEnCode(ListPageNext)
			  If ListPageNext <> "Error" And ListPageNext <> "" Then
				 If ListPageStr1 <> "" Then
					ListPageNext = Replace(ListPageStr1, "{$ID}", ListPageNext)
				 Else
					ListPageNext = KMCObj.DefiniteUrl(ListPageNext, ListUrl)
				 End If
				 ListPageNext = Replace(ListPageNext, "&", "{$ID}")
			  End If
		   Else
			  ListPageNext = "Error"
		   End If
		End Sub
		
		'==================================================
		'过程名：SaveArticle
		'作  用：保存文章
		'参  数：无
		'==================================================
		Sub SaveArticle()
		    'on error resume next
			Dim FsoType,CBody,Images
			FsoType = conn.Execute("select top 1 FsoType from KS_class where id='" & ClassID & "'")(0)
		   Set Rs = Server.CreateObject("adodb.recordset")
		   If RepeatInto<>"1" Then
		   Sql = "select top 1 * from "& KS.C_S(ChannelID,2) & " where Title='" & Title & "' and Tid='" & ClassID & "'"
		   Else
		   Sql = "select top 1 * from "& KS.C_S(ChannelID,2) & " where 1=0"
		   End If

		   If IntoBase <>0 Then          '直接插入数据库
			 Rs.Open Sql, conn, 1, 3
		   Else
			 Rs.Open Sql, ConnItem, 1, 3
		   End If
		   If Rs.EOF And rs.bof Then
		   Rs.AddNew

		   Rs("Tid") = ClassID
		   Rs("Keywords") = Key
		   Rs("Title") = Title
		  If KS.IsNul(Content) Then Content="&nbsp;"
		  Select Case  KS.C_S(ChannelID,6)
		   Case 1
		   Rs("TitleType") = ""
		   Rs("ShowComment") = ShowComment
		   Rs("TitleFontColor") = ""
		   Rs("TitleFontType") = ""
		   Rs("ArticleContent") = Content
		   Rs("Intro")=Left(KS.LoseHtml(Content),255)
		   Rs("JSID") = ""
		   Rs("Changes") = 0
           If ThumbType=1 Then
             Rs("PhotoUrl")=ThumbUrl
			 Rs("PicNews")=1
		   Else       
			   If PicNews=1 And DefaultPicUrl<>"" Then
				Rs("PhotoUrl")=DefaultPicUrl
				Rs("PicNews")=1
			   Else
			    Rs("PicNews") = 0
			   End If
		   End If
		   Rs("Author") = Author
		   Rs("Origin") = KS.GotTopic(CopyFrom,50)
		   Images=Images&Content&ThumbUrl&DefaultPicUrl
		  Case 2
		   Rs("ShowStyle")=1
		   Rs("PageNum")=10
		   Rs("Author") = Author
		   Rs("Origin") = KS.GotTopic(CopyFrom,50)
		   Rs("PictureContent")= Content
		   If ThumbUrl<>"" Then
		   Rs("PhotoUrl")=ThumbUrl
		   Else
		   Rs("PhotoUrl")=DefaultPicUrl
		   End If
		   Rs("PicUrls")=tp_str
		   Images=Images&Content&ThumbUrl&DefaultPicUrl&tp_str
		  Case 5
			 If ThumbUrl<>"" Then
			   Rs("PhotoUrl")=ThumbUrl
			 Else
			   Rs("PhotoUrl")=DefaultPicUrl
			 End If

		    Rs("ProID")=KS.GetInfoID(ChannelID)
			Rs("ProIntro")= Content
			If Tp_Url="Error" Or Tp_Url="" Then
			Rs("BigPhoto")=Rs("PhotoUrl")
			Else
			 Rs("BigPhoto")=Tp_Url
			End If
			'计量单位
			Rs("Unit")=GetCollectValue(Shop_UnitBeginStr,Shop_UnitEndStr,"件")
			'会员价格
			Cbody=GetCollectValue(Shop_MemberPriceBeginStr,Shop_MemberPriceEndStr,0)
			If IsNumeric(Cbody) Then Rs("Price_Member")=Cbody  Else Rs("Price_Member")=0
			'市场价格
			Cbody=GetCollectValue(Shop_MarketPriceBeginStr,Shop_MarketPriceEndStr,0)
			If IsNumeric(Cbody) Then Rs("Price_Market")=Cbody  Else Rs("Price_Market")=0
			'原始价格
			Cbody=GetCollectValue(Shop_OriginPriceBeginStr,Shop_OriginPriceEndStr,0)
			If IsNumeric(Cbody) Then Rs("Price_Original")=Cbody  Else Rs("Price_Original")=0
			'商城价
			Cbody=GetCollectValue(Shop_PriceBeginStr,Shop_PriceEndStr,0)
			If IsNumeric(Cbody) Then Rs("Price")=Cbody  Else Rs("Price")=0
			'型号
			Rs("ProModel")=GetCollectValue(Shop_ProModelBeginStr,Shop_ProModelEndStr,"")
			'规格
			Rs("ProSpecificat")=GetCollectValue(Shop_ProSpecificatBeginStr,Shop_ProSpecificatEndStr,"")
			'生产商
			Rs("ProducerName")=GetCollectValue(Shop_ProducerNameBeginStr,Shop_ProducerNameEndStr,"")
			'商标
			Rs("TrademarkName")=GetCollectValue(Shop_TrademarkNameBeginStr,Shop_TrademarkNameEndStr,"")
            Images=Images&Content&ThumbUrl&DefaultPicUrl&Tp_Url
		  End Select
		   
		   Rs("Rank") = Stars           '阅读星级
		   Rs("Hits") = Hits
		   Rs("AddDate") = UpDateTime   '更新时间
		   Rs("TemplateID") = KS.C_C(ClassID,5) '模板
		   Rs("Fname") = KS.GetFileName(FsoType, UpDateTime, FnameType)
		   Rs("Inputer") = left(Inputer,50)
		   Rs("Recommend") = Recommend
		   Rs("Rolls") = Rolls
		   Rs("Strip")=0
		   Rs("Popular") = Popular
		   Rs("Verific") = Verific      '审核与否
		   Rs("Slide") = 0
		   Rs("Comment") = Comment
		  
		   If IsArray(Arr_Field) Then
			  For StepID=0 To Ubound(Arr_Field,2)
			     If Arr_Field(0,StepID)<>0 Then
					If KS.ChkClng(Arr_Field(5,StepID))=ItemID Then
					   If KS.ChkClng(Arr_Field(4,StepID))=0 Then
						  If Arr_Field(2,StepID)<>"" And Arr_Field(3,StepID)<>"" Then '开始标记和结束标记都不为空
						   Cbody=KMCObj.GetBody(NewsCode, Arr_Field(2,StepID),Arr_Field(3,StepID), False, False)
						   If Cbody <> "Error" and Cbody <> "" Then
						   rs(Arr_Field(1,StepID))=KMCObj.FpHtmlEnCode(Cbody)
						   End If
						  Else
						   rs(Arr_Field(1,StepID))=Arr_Field(2,StepID)
						  End If
					   Else
							  '======================列表采集开始======================
							  Dim DiyField:DiyField=KMCObj.GetArray(ListCode, Arr_Field(2,StepID), Arr_Field(3,StepID), False, False)
							  Dim DiyFieldArr,DiyFieldArrLen
							  If DiyField<>"Error" And DiyField<>"" Then
								  DiyFieldArr=Split(DiyField,"$Array$")
								  DiyFieldArrLen=Ubound(DiyFieldArr)
								  rs(Arr_Field(1,StepID))=DiyFieldArr(Arr_I)
							  End If
							  '========================================================
					  End IF
				   End If
				End If
			  Next
			End IF
		   
		   Rs.Update
		   rs.movelast
		   if rs("fname")="ID" Then
		    rs("fname")=rs("id") & fnametype
			rs.update
		   end if
		   
		   
 
		   If IntoBase = 1 or IntoBase=2 Then          '当直接插入数据库,同时往主数据库写数据
		     Call KS.FileAssociation(ChannelID,RS("ID"),Images ,0)
		     Call LFCls.AddItemInfo(ChannelID,RS("ID"),Title,ClassID,left(KS.LoseHtml(Content),255),Key,Rs("PhotoUrl"),UpDateTime,KS.C("AdminName"),rs("Hits"),rs("HitsByDay"),rs("HitsByWeek"),rs("HitsByMonth"),rs("Recommend"),rs("Rolls"),rs("Strip"),rs("Popular"),rs("Slide"),rs("IsTop"),rs("Comment"),rs("Verific"),RS("Fname"))
			 If Verific=1 and IntoBase=2 Then 
					If (KS.C_S(Channelid,7)=1 or KS.C_S(ChannelID,7)=2) Then
					 Dim KSRObj:Set KSRObj=New Refresh
					 Dim DocXML:Set DocXML=KS.RsToXml(RS,"row","root")
				     Set KSRObj.Node=DocXml.DocumentElement.SelectSingleNode("row")
					  KSRObj.ModelID=ChannelID
					  KSRObj.ItemID = KSRObj.Node.SelectSingleNode("@id").text 
					  Call KSRObj.RefreshContent()
					  Set KSRobj=Nothing
					End If
			End If
			
		   End If
		   
		   End If
		   Rs.Close
		   Set Rs = Nothing
		End Sub
		
		Function GetCollectValue(BeginStr,EndStr,DefaultValue)
		    Dim Cbody
		    If BeginStr<>"" And  EndStr<>"" Then
			 Cbody=KMCObj.GetBody(NewsCode, BeginStr,EndStr, False, False)
			 If Cbody <> "Error" and Cbody <> "" Then
			 Cbody=KMCObj.FpHtmlEnCode(Cbody)
			 Else
			  Cbody="件"
			 End If
			Else
			 Cbody=DefaultValue
			End If
			GetCollectValue=Cbody
		End Function
		
		
		'==================================================
		'过程名：Filters
		'作  用：过滤
		'==================================================
		Sub Filters()
		If IsNull(Arr_Filters) = True Or IsArray(Arr_Filters) = False Then
		   Exit Sub
		End If
		
		   For Filteri = 0 To UBound(Arr_Filters, 2)
			  FilterStr = ""
			  If Arr_Filters(1, Filteri) = ItemID Or Arr_Filters(10, Filteri) = True Then
				 If Arr_Filters(3, Filteri) = 1 Then '标题过滤
					If Arr_Filters(4, Filteri) = 1 Then
					   Title = Replace(Title, Arr_Filters(5, Filteri), Arr_Filters(8, Filteri))
					ElseIf Arr_Filters(4, Filteri) = 2 Then
					   FilterStr = KMCObj.GetBody(Title, Arr_Filters(6, Filteri), Arr_Filters(7, Filteri), True, True)
					   Do While FilterStr <> "Error"
						  Title = Replace(Title, FilterStr, Arr_Filters(8, Filteri))
						  FilterStr = KMCObj.GetBody(Title, Arr_Filters(6, Filteri), Arr_Filters(7, Filteri), True, True)
					   Loop
					End If
				 ElseIf Arr_Filters(3, Filteri) = 2 Then '正文过滤
					If Arr_Filters(4, Filteri) = 1 Then
					   Content = Replace(Content, Arr_Filters(5, Filteri), Arr_Filters(8, Filteri))
					ElseIf Arr_Filters(4, Filteri) = 2 Then
					   FilterStr = KMCObj.GetBody(Content, Arr_Filters(6, Filteri), Arr_Filters(7, Filteri), True, True)
					   Do While FilterStr <> "Error"
						  Content = Replace(Content, FilterStr, Arr_Filters(8, Filteri))
						  FilterStr = KMCObj.GetBody(Content, Arr_Filters(6, Filteri), Arr_Filters(7, Filteri), True, True)
					   Loop
					End If
				 End If
			  End If
		   Next
		End Sub
		
		'==================================================
		'过程名：FilterScript
		'作  用：脚本过滤
		'==================================================
		
		Sub FilterScript()
		   If Script_Iframe = True Then
			  Content = KS.ScriptHtml(Content, "Iframe", 1)
		   End If
		   If Script_Object = True Then
			  Content = KS.ScriptHtml(Content, "Object", 2)
		   End If
		   If Script_Script = True Then
			  Content = KS.ScriptHtml(Content, "Script", 2)
		   End If
		   If Script_Div = True Then
			  Content = KS.ScriptHtml(Content, "Div", 3)
		   End If
		   If Script_Table = True Then
			  Content = KS.ScriptHtml(Content, "table", 3)
		   End If
		   If Script_Tr = True Then
			  Content = KS.ScriptHtml(Content, "tr", 3)
		   End If
		   If Script_Td = True Then
			  Content = KS.ScriptHtml(Content, "td", 3)
		   End If
		   If Script_Span = True Then
			  Content = KS.ScriptHtml(Content, "Span", 3)
		   End If
		   If Script_Img = True Then
			  Content = KS.ScriptHtml(Content, "Img", 3)
		   End If
		   If Script_Font = True Then
			  Content = KS.ScriptHtml(Content, "Font", 3)
		   End If
		   If Script_A = True Then
			  Content = KS.ScriptHtml(Content, "A", 3)
		   End If
		   If Script_Html = True Then
			  Content = KMCObj.nohtml(Content)
		   End If
		End Sub
		
		Function CheckRepeat(strUrl)
		   CheckRepeat = False
		   If IsArray(Arr_Historys) = True Then
			  For His_i = 0 To UBound(Arr_Historys, 2)
				If Arr_Historys(0, His_i) = strUrl Then
					CheckRepeat = True
					His_Title = Arr_Historys(1, His_i)
					His_CollecDate = Arr_Historys(2, His_i)
					His_Result = Arr_Historys(3, His_i)
					Exit For
				 End If
			  Next
		   End If
		   
		End Function
		
		Sub SetCache_His()
		   '历史记录
		   SqlItem = "select NewsUrl,Title,CollecDate,Result From KS_History"
		   Set RsItem = Server.CreateObject("adodb.recordset")
		   RsItem.Open SqlItem, ConnItem, 1, 1
		   If Not RsItem.EOF Then
			  Arr_Historys = RsItem.GetRows()
		   End If
		   RsItem.Close
		   Set RsItem = Nothing
		
		   Dim myCache
		   Set myCache = New ClsCache
		   myCache.name = CacheTemp & "Historys"
		   Call myCache.clean
		   If IsArray(Arr_Historys) = True Then
			  myCache.add Arr_Historys, DateAdd("n", 1000, Now)
		   End If
		End Sub
End Class
%>
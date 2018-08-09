<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%> 
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.CollectCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Collect_ItemModify5
KSCls.Kesion()
Set KSCls = Nothing

Class Collect_ItemModify5
        Private KS
		Private KMCObj
		Private ConnItem
		Private ItemID, Action
		Private RsItem, SqlItem, SqlF, RsF, FoundErr, ErrMsg
		Private LoginType, LoginUrl, LoginPostUrl, LoginUser, LoginPass, LoginFalse, LoginResult, LoginData
		Private ListStr, LsString, LoString, ListPageType, LPsString, LPoString, ListPageStr1, ListPageStr2, ListPageID1, ListPageID2, ListPageStr3, HsString, HoString, HttpUrlType, HttpUrlStr
		Private TsString, ToString, CsString, CoString, DateType, DsString, DoString, UpDateTime, AuthorType, AsString, AoString, AuthorStr, CopyFromType, FsString, FoString, CopyFromStr, KeyType, KsString, KoString, KeyStr, NewsPageType, NPsString, NPoString, NewsPageStr, NewsPageEnd,CharsetCode
		Private NewsPageNext, NewsPageNextCode, ContentTemp
		Private UrlTest, ListUrl, ListCode
		Private NewsUrl, NewsCode, NewsArrayCode, NewsArray
		Private Title, Content, Author, CopyFrom, Key
		Private Arr_Filters, Filteri, FilterStr
		Private UpDateType,Tp_Lists,Tp_Listo,Tp_Srcs,Tp_Srco,Tp_Is,Tp_Io
		
		Private InfoPageStr
		Private InfoPageArrayCode,InfoPageArray,Testi

		
		Private UploadFiles, strInstallDir, strChannelDir
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
			strInstallDir = Trim(Request.ServerVariables("SCRIPT_NAME"))
			strInstallDir = Left(strInstallDir, InStrRev(LCase(strInstallDir), "/") - 1)
			strInstallDir = Left(strInstallDir, InStrRev(LCase(strInstallDir), "/"))
			strChannelDir = "Test"
			FoundErr = False
			
			ItemID = Trim(Request("ItemID"))
			Action = Trim(Request("Action"))
			
			If ItemID = "" Then
			   FoundErr = True
			   ErrMsg = ErrMsg & "●参数错误，项目ID不能为空\n"
			Else
			   ItemID = CLng(ItemID)
			End If
			
			If Action = "SaveEdit" And FoundErr <> True Then
			   Call SaveEdit
			End If
			
			If FoundErr <> True Then
			   Call GetTest
			End If
			If FoundErr <> True Then
			   Call Main
			Else
			   Call KS.AlertHistory(ErrMsg,-1)
			End If
			End Sub
			Sub Main()
			Response.Write "<html>"
			Response.Write "<head>"
			Response.Write "<title>采集系统</title>"
			Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			Response.Write "<link rel=""stylesheet"" type=""text/css"" href=""../Include/Admin_Style.css"">"
			Response.Write "</head>" & vbCrLf
			Response.Write "<div class='topdashed'>"& KMCObj.GetItemLocation(5,ItemID) &"</div>"
		
			Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
			Response.Write "<tr align=""center"">"
			Response.Write "    <td colspan=""2"" valign=""bottom"">"
			Response.Write "    <font size=""3"">" & Title & "</font></td>"
			Response.Write "</tr>"
			Response.Write "  <tr align=""center"">"
			Response.Write "    <td colspan=""2"">"
			Response.Write "        作者：" & Author & "&nbsp;&nbsp;来源：" & CopyFrom & "&nbsp;&nbsp;更新时间：" & UpDateTime
			Response.Write "    </td>"
			Response.Write "  </tr>"
			Response.Write "  <tr>"
			Response.Write "    <td colspan=""2"" align=""center"" valign=""top"">"
			Response.Write "      <div style=""border: double #E7E7E7;height:345; overflow: auto; width:95%"" align=""center"">"
			Response.Write "      <table width=""95%"" height=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""5"">"
			 Response.Write "       <tr>"
			 Response.Write "         <td height=""200"" valign=""top""><p>" & Content & "</p>"
			 Response.Write "         </td>"
			 Response.Write "       </tr>"
			 Response.Write "     </table>"
			 Response.Write "     </div>"
			 Response.Write "     <div align=""center"" style=""height:25""><b>关键字：" & Key & "</b></div>"
			 Response.Write "   </td>"
			 Response.Write " </tr>"
			Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
			Response.Write "<form method=""post"" action=""Collect_ItemAttribute.asp"" name=""form1"">"
			Response.Write "    <tr>"
			Response.Write "      <td colspan=""2"" style=""text-align:center"">"
			Response.Write "        <input name=""Action"" type=""hidden"" id=""Action"" value=""SaveEdit"">"
			Response.Write "        <input name=""ItemID"" type=""hidden"" id=""ItemID"" value=""" & ItemID & """>"
			Response.Write "        <input name=""Cancel"" class='button' type=""button"" id=""Cancel"" value="" 上&nbsp;一&nbsp;步 "" onClick=""window.location.href='javascript:history.go(-1)'"">"
			Response.Write "        &nbsp;"
			Response.Write "        <input  type=""submit"" class='button' name=""Submit"" value=""  下&nbsp;一&nbsp;步 ""></td>"
			Response.Write "    </tr>"
		  	Response.Write "</form>"
			Response.Write "</table>"
			
			if NewsPageType=1 And isarray(InfoPageArray) Then
					Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
				   Response.Write "  <tr>"
					Response.Write "   <td height=""22"" align=""center""><font color=red>分析得到的正文分页URL，请检查是否正确:</font>"
					Response.Write "<select name=pagelist style=""width:380"">"	
					For Testi = 0 To UBound(InfoPageArray)
					   Response.Write "<option>" & InfoPageArray(Testi) & "</option>"
					Next
				
				  Response.Write " </select></td>"
				  Response.Write "</tr>"
				  Response.Write "</table>"
		  End IF
		  
		  	Response.Write "</body>"
			Response.Write "</html>"
			End Sub
			'==================================================
			'过程名：SaveEdit
			'作  用：保存设置
			'参  数：无
			'==================================================
			Sub SaveEdit()
			
			TsString = Request.Form("TsString")
			ToString = Request.Form("ToString")
			CsString = Request.Form("CsString")
			CoString = Request.Form("CoString")
			
			DateType = Trim(Request.Form("DateType"))
			DsString = Request.Form("DsString")
			DoString = Request.Form("DoString")
			
			AuthorType = Trim(Request.Form("AuthorType"))
			AsString = Request.Form("AsString")
			AoString = Request.Form("AoString")
			AuthorStr = Request.Form("AuthorStr")
			
			CopyFromType = Trim(Request.Form("CopyFromType"))
			FsString = Request.Form("FsString")
			FoString = Request.Form("FoString")
			CopyFromStr = Request.Form("CopyFromStr")
			
			KeyType = Trim(Request.Form("KeyType"))
			KsString = Request.Form("KsString")
			KoString = Request.Form("KoString")
			KeyStr = Request.Form("KeyStr")
			
			
			NewsPageType = Trim(Request.Form("NewsPageType"))
			NPsString = Request.Form("NpsString")
			NPoString = Request.Form("NpoString")
			NewsPageStr = Request.Form("NewsPageStr")
			NewsPageEnd = Request.Form("NewsPageEnd")
		
			UrlTest = Trim(Request.Form("UrlTest"))
			
			
			If ItemID = "" Then
			   FoundErr = True
			   ErrMsg = ErrMsg & "●参数错误，项目ID不能为空\n"
			Else
			   ItemID = CLng(ItemID)
			End If
			If UrlTest = "" Then
			   FoundErr = True
			   ErrMsg = ErrMsg & "●参数错误，数据传递时发生错误\n"
			Else
				  NewsUrl = UrlTest
			End If
			If TsString = "" Then
			   FoundErr = True
			   ErrMsg = ErrMsg & "●标题开始标记不能为空\n"
			End If
			If ToString = "" Then
			   FoundErr = True
			   ErrMsg = ErrMsg & "●标题结束标记不能为空\n"
			End If
			If CsString = "" Then
			   FoundErr = True
			   ErrMsg = ErrMsg & "●正文开始标记不能为空\n"
			End If
			If CoString = "" Then
			   FoundErr = True
			   ErrMsg = ErrMsg & "●正文结束标记不能为空\n"
			End If
			
			If Request("ChannelID")<>"" Then
			  If KS.C_S(KS.ChkCLng(Request("ChannelID")),6)=2 Then
					If Request.Form("Tp_listBeginStr")="" Then
					   FoundErr = True
					   ErrMsg = ErrMsg & "●图片地址列表开始标记不能为空\n"
					End If
					If Request.Form("Tp_listEndStr")="" Then
					   FoundErr = True
					   ErrMsg = ErrMsg & "●图片地址列表结束标记不能为空\n"
					End If
					If Request.Form("Tp_SrcBeginStr")="" Then
					   FoundErr = True
					   ErrMsg = ErrMsg & "●单张图片地址开始标记不能为空\n"
					End If
					If Request.Form("Tp_SrcEndStr")="" Then
					   FoundErr = True
					   ErrMsg = ErrMsg & "●单张图片地址结束标记不能为空\n"
					End If
			  End If
			End If
			
			
			If DateType = "" Then
			   FoundErr = True
			   ErrMsg = ErrMsg & "●请设置时间类型\n"
			Else
			   DateType = CLng(DateType)
			   If DateType = 0 Then
			   ElseIf DateType = 1 Then
				  If DsString = "" Or DoString = "" Then
					 FoundErr = True
					 ErrMsg = ErrMsg & "●请将时间的开始/结束标记填写完整\n"
				  End If
			   Else
				  FoundErr = True
				  ErrMsg = ErrMsg & "●参数错误，请从有效链接进入\n"
			   End If
			End If
			
			If AuthorType = "" Then
			   FoundErr = True
			   ErrMsg = ErrMsg & "●请设置作者类型\n"
			Else
			   AuthorType = CLng(AuthorType)
			   If AuthorType = 0 Then
			   ElseIf AuthorType = 1 Then
				  If AsString = "" Or AoString = "" Then
					 FoundErr = True
					 ErrMsg = ErrMsg & "●请将作者开始/结束标记填写完整！\n"
				  End If
			   ElseIf AuthorType = 2 Then
				  If AuthorStr = "" Then
					 FoundErr = True
					 ErrMsg = ErrMsg & "●请指定作者\n"
				  End If
			   Else
				  FoundErr = True
				  ErrMsg = ErrMsg & "●参数错误，请从有效链接进入\n"
			   End If
			End If
			
			If CopyFromType = "" Then
			   FoundErr = True
			   ErrMsg = ErrMsg & "●请设置来源类型\n"
			Else
			   CopyFromType = CLng(CopyFromType)
			   If CopyFromType = 0 Then
			   ElseIf CopyFromType = 1 Then
				  If FsString = "" Or FoString = "" Then
					 FoundErr = True
					 ErrMsg = ErrMsg & "●请将来源开始/结束标记填写完整！\n"
				  End If
			   ElseIf CopyFromType = 2 Then
				  If CopyFromStr = "" Then
					 FoundErr = True
					 ErrMsg = ErrMsg & "●请指定来源\n"
				  End If
			   Else
				  FoundErr = True
				  ErrMsg = ErrMsg & "●参数错误，请从有效链接进入\n"
			   End If
			End If
			
			If KeyType = "" Then
			   FoundErr = True
			   ErrMsg = ErrMsg & "●请设置关键字类型\n"
			Else
			   KeyType = CLng(KeyType)
			   If KeyType = 0 Then
			   ElseIf KeyType = 1 Then
				  If KsString = "" Or KoString = "" Then
					 FoundErr = True
					 ErrMsg = ErrMsg & "●关键字开始/结束标记不能为空\n"
				  End If
			   ElseIf KeyType = 2 Then
				  If KeyStr = "" Then
					 FoundErr = True
					 ErrMsg = ErrMsg & "●请指定关键字\n"
				  End If
			   Else
				  FoundErr = True
				  ErrMsg = ErrMsg & "●参数错误，请从有效链接进入\n"
			   End If
			End If
			
			If NewsPageType = "" Then
			   FoundErr = True
			   ErrMsg = ErrMsg & "●请设置新闻分页类型\n"
			Else
			   NewsPageType = CLng(NewsPageType)
			   If NewsPageType = 0 Then
			   ElseIf NewsPageType = 1 Then
				  If NPsString = "" Or NPoString = "" Then
					 FoundErr = True
					 ErrMsg = ErrMsg & "●分页代码开始/分页代码结束标记不能为空\n"
				  End If
				  If NewsPageStr = "" or NewsPageEnd="" Then
					 FoundErr = True
					 ErrMsg = ErrMsg & "●分页URL开始代码/分页URL结束代码不能为空\n"
				  End If
			   ElseIf NewsPageType = 2 Then
				  FoundErr = True
				  ErrMsg = ErrMsg & "●暂不支持手动设置分页类型\n"
			   Else
				  FoundErr = True
				  ErrMsg = ErrMsg & "●参数错误，请从有效链接进入\n"
			   End If
			End If
			
			If FoundErr <> True Then
			   SqlItem = "Select top 1 * From KS_CollectItem Where ItemID=" & ItemID
			   Set RsItem = Server.CreateObject("adodb.recordset")
			   RsItem.Open SqlItem, ConnItem, 2, 3
			   RsItem("TsString") = TsString
			   RsItem("ToString") = ToString
			   RsItem("CsString") = CsString
			   RsItem("CoString") = CoString
			
			   RsItem("DateType") = DateType
			   RsItem("UpDateType") =DateType
			   If DateType = 1 Then
				  RsItem("DsString") = DsString
				  RsItem("DoString") = DoString
			   End If
			
			   RsItem("AuthorType") = AuthorType
			   If AuthorType = 1 Then
				  RsItem("AsString") = AsString
				  RsItem("AoString") = AoString
			   ElseIf AuthorType = 2 Then
				  RsItem("AuthorStr") = AuthorStr
			   End If
			
			   RsItem("CopyFromType") = CopyFromType
			   If CopyFromType = 1 Then
				  RsItem("FsString") = FsString
				  RsItem("FoString") = FoString
			   ElseIf CopyFromType = 2 Then
				  RsItem("CopyFromStr") = CopyFromStr
			   End If
			
			   RsItem("KeyType") = KeyType
			   If KeyType = 1 Then
				  RsItem("KsString") = KsString
				  RsItem("KoString") = KoString
			   ElseIf KeyType = 2 Then
				  RsItem("KeyStr") = KeyStr
			   End If
			
			   RsItem("NewsPageType") = NewsPageType
			   If NewsPageType = 1 Then
				  RsItem("NPsString") = NPsString
				  RsItem("NPoString") = NPoString
				  RsItem("NewsPageStr") = NewsPageStr
				  RsItem("NewsPageEnd") = NewsPageEnd
			   ElseIf NewsPageType = 2 Then
			   End If
			     ' RsItem("Tp_Lists")=Tp_Lists
				 ' RsItem("Tp_Listo")=Tp_Listo
			    '  RsItem("Tp_Srcs")=Tp_Srcs
				 ' RsItem("Tp_Srco")=Tp_Srco
				'  RsItem("Tp_Is")=Tp_Is
				  'RsItem("Tp_Io")=Tp_Io
			   RsItem.Update
			     Dim RS,SQL,I,RSV
				 Set RS=Server.CreateObject("ADODB.RECORDSET")
				 RS.Open "Select FieldTitle,FieldName,ChannelID,FieldID,OrderID,ShowType From KS_FieldItem Where ShowType=0 and ChannelID=" &RsItem("ChannelID") & " order by orderid",ConnItem,1,1
				 Do While Not RS.Eof 
				   Set RSV=Server.CreateObject("ADODB.RECORDSET")
				   RSV.Open "Select * From KS_FieldRules Where ItemID=" & ItemID & " and channelid=" & RS(2) & " and fieldid=" & rs(3),connItem,1,3
				   If RSV.Eof And RSV.Bof Then
				    RSV.AddNew
				   End If
				   RSV("ItemID")=ItemID
				   RSV("ChannelID")=rs(2)
				   RSV("FieldID")=rs(3)
				   RSV("FieldName")=rs(1)
				   RSV("OrderID")=rs(4)
				   RSV("BeginStr")=Request.Form("begin"&rs(1))
				   RSV("EndStr")=Request.Form("end"&rs(1))
				   RSV("ShowType")=rs(5)
				   RSV.Update
				   RSV.close
				   RS.MoveNext
				 Loop
				 RS.Close
				 
				 Dim FieldNameArr,FieldNameList
				 If KS.C_S(RsItem("ChannelID"),6)=2 Then
				  FieldNameList="Tp_List,Tp_Src,Tp_I"
				 ElseIf KS.C_S(RsItem("ChannelID"),6)=5 Then
				  FieldNameList="Shop_BigPhoto,Shop_BigPhotoSrc,Shop_Unit,Shop_OriginPrice,Shop_Price,Shop_MarketPrice,Shop_MemberPrice,Shop_ProModel,Shop_ProSpecificat,Shop_ProducerName,Shop_TrademarkName"
				 End If
				 FieldNameArr=Split(FieldNameList,",")
				 For I=0 To Ubound(FieldNameArr)
					 RS.Open "Select top 1 * From KS_FieldRules Where ItemID=" & ItemID & " And ChannelID=" & RsItem("ChannelID") & " And FieldName='" & FieldNameArr(i) & "'",connItem,1,3
					 If RS.Eof Then
					  RS.AddNew
					  RS("ItemID")=ItemID
					  RS("ChannelID")=RsItem("ChannelID")
					  RS("FieldID")=0
					  RS("FieldName")=FieldNameArr(I)
					 End If
					 RS("BeginStr")=Request.Form(FieldNameArr(I)&"BeginStr")
					 RS("EndStr")=Request.Form(FieldNameArr(I)&"EndStr")
					 RS.Update
					 RS.Close
				 Next
				 
				 Set RS=Nothing
				 
				 
			   RsItem.Close
			   Set RsItem = Nothing
			   
			End If
			End Sub
			
			'==================================================
			'过程名：GetTest
			'作  用：采集测试
			'参  数：无
			'==================================================
			Sub GetTest()
			   SqlItem = "Select top 1 * From KS_CollectItem Where ItemID=" & ItemID
			   Set RsItem = Server.CreateObject("adodb.recordset")
			   RsItem.Open SqlItem, ConnItem, 1, 1
			   If RsItem.EOF And RsItem.BOF Then
					 FoundErr = True
				  ErrMsg = ErrMsg & "●参数错误，找不到该项目\n"
			   Else
				  LoginType = RsItem("LoginType")
				  LoginUrl = RsItem("LoginUrl")
				  LoginPostUrl = RsItem("LoginPostUrl")
				  LoginUser = RsItem("LoginUser")
				  LoginPass = RsItem("LoginPass")
				  LoginFalse = RsItem("LoginFalse")
				  
				  ListStr = RsItem("ListStr")
				  LsString = RsItem("LsString")
				  LoString = RsItem("LoString")
				  ListPageType = RsItem("ListPageType")
				  LPsString = RsItem("LPsString")
				  LPoString = RsItem("LPoString")
				  ListPageStr1 = RsItem("ListPageStr1")
				  ListPageStr2 = RsItem("ListPageStr2")
				  ListPageID1 = RsItem("ListPageID1")
				  ListPageID2 = RsItem("ListPageID2")
				  ListPageStr3 = RsItem("ListPageStr3")
				  
				  HsString = RsItem("HsString")
				  HoString = RsItem("HoString")
				  HttpUrlType = RsItem("HttpUrlType")
				  HttpUrlStr = RsItem("HttpUrlStr")
				  
				  TsString = RsItem("TsString")
				  ToString = RsItem("ToString")
				  CsString = RsItem("CsString")
				  CoString = RsItem("CoString")
				  
				  DateType = RsItem("DateType")
				  DsString = RsItem("DsString")
				  DoString = RsItem("DoString")
			
				  AuthorType = RsItem("AuthorType")
				  AsString = RsItem("AsString")
				  AoString = RsItem("AoString")
				  AuthorStr = RsItem("AuthorStr")
			
				  CopyFromType = RsItem("CopyFromType")
				  FsString = RsItem("FsString")
				  FoString = RsItem("FoString")
				  CopyFromStr = RsItem("CopyFromStr")
			
				  KeyType = RsItem("KeyType")
				  KsString = RsItem("KsString")
				  KoString = RsItem("KoString")
				  KeyStr = RsItem("KeyStr")
			
				  NewsPageType = RsItem("NewsPageType")
				  NPsString = RsItem("NPsString")
				  NPoString = RsItem("NPoString")
				  NewsPageStr = RsItem("NewsPageStr")
				  NewsPageEnd = RsItem("NewsPageEnd")
				  
				  CharsetCode = RsItem("CharsetCode")
				  UpDateType = RsItem("UpDateType")
			   End If
			   RsItem.Close
			   Set RsItem = Nothing
			
			   If LoginType = 1 Then
				  If LoginUrl = "" Or LoginPostUrl = "" Or LoginUser = "" Or LoginPass = "" Or LoginFalse = "" Then
					 FoundErr = True
					 ErrMsg = ErrMsg & "●您要采集的网站需要登录！请将登录信息填写完整\n"
				  End If
			   End If
			   If LsString = "" Then
				  FoundErr = True
				  ErrMsg = ErrMsg & "●列表开始标记不能为空！\n"
			   End If
			   If LoString = "" Then
				  FoundErr = True
				  ErrMsg = ErrMsg & "●列表结束标记不能为空！\n"
			   End If
			   If ListPageType = 0 Or ListPageType = 1 Then
				  If ListStr = "" Then
					 FoundErr = True
					 ErrMsg = ErrMsg & "●列表索引页不能为空！\n"
				  End If
				  If ListPageType = 1 Then
					 If LPsString = "" Or LPoString = "" Then
						FoundErr = True
						ErrMsg = ErrMsg & "●索引分页开始、结束标记不能为空！\n"
					 End If
				  End If
				  If ListPageStr1 <> "" And Len(ListPageStr1) < 15 Then
					 FoundErr = True
					 ErrMsg = ErrMsg & "●索引分页重定向设置不正确！\n"
						End If
			   ElseIf ListPageType = 2 Then
				  If ListPageStr2 = "" Then
					 FoundErr = True
					 ErrMsg = ErrMsg & "●批量生成原字符串不能为空！\n"
				  End If
				  If IsNumeric(ListPageID1) = False Or IsNumeric(ListPageID2) = False Then
					 FoundErr = True
					 ErrMsg = ErrMsg & "●批量生成的范围只能是数字！\n"
				  Else
					 ListPageID1 = CLng(ListPageID1)
					 ListPageID2 = CLng(ListPageID2)
					 If ListPageID1 = 0 And ListPageID2 = 0 Then
						FoundErr = True
						ErrMsg = ErrMsg & "●批量生成的范围不正确！\n"
					 End If
				  End If
			   ElseIf ListPageType = 3 Then
				  If ListPageStr3 = "" Then
					 FoundErr = True
					 ErrMsg = ErrMsg & "●索引分页不能为空！\n"
				  End If
			   Else
				  FoundErr = True
				  ErrMsg = ErrMsg & "●请选择索引分页类型\n"
			   End If
				 If HsString = "" Or HoString = "" Then
				  FoundErr = True
				  ErrMsg = ErrMsg & "●链接开始/结束标记不能为空！\n"
				  End If
			
			
			   If FoundErr <> True And Action <> "SaveEdit" Then
				  Select Case ListPageType
				  Case 0, 1
						ListUrl = ListStr
				  Case 2
					 ListUrl = Replace(ListPageStr2, "{$ID}", CStr(ListPageID1))
				  Case 3
					 If InStr(ListPageStr3, "|") > 0 Then
						ListUrl = Left(ListPageStr3, InStr(ListPageStr3, "|") - 1)
					 Else
						ListUrl = ListPageStr3
					 End If
				  End Select
			   End If
			   
				  If FoundErr <> True And Action <> "SaveEdit" And LoginType = 1 Then
				  LoginData = KMCObj.UrlEncoding(LoginUser & "&" & LoginPass)
				  LoginResult = KMCObj.PostHttpPage(LoginUrl, LoginPostUrl, LoginData)
				  If InStr(LoginResult, LoginFalse) > 0 Then
					 FoundErr = True
					 ErrMsg = ErrMsg & "●登录网站时发生错误，请确认登录信息的正确性！\n"
				  End If
				  End If
			   
			   If FoundErr <> True And Action <> "SaveEdit" Then
					 ListCode = KMCObj.GetHttpPage(ListUrl,CharsetCode)
					 If ListCode <> "Error" Then
						ListCode = KMCObj.GetBody(ListCode, LsString, LoString, False, False)
						If ListCode <> "Error" Then
						   NewsArrayCode = KMCObj.GetArray(ListCode, HsString, HoString, False, False)
						   If NewsArrayCode <> "Error" Then
							  If InStr(NewsArrayCode, "$Array$") > 0 Then
								 NewsArray = Split(NewsArrayCode, "$Array$")
								 If HttpUrlType = 1 Then
									NewsUrl = Trim(Replace(HttpUrlStr, "{$ID}", NewsArray(0)))
								 Else
									NewsUrl = Trim(KMCObj.DefiniteUrl(NewsArray(0), ListUrl))
								 End If
							  Else
								 FoundErr = True
								 ErrMsg = ErrMsg & "●只发现一个有效链接？：" & NewsArrayCode & "\n"
							 End If
						  Else
							 FoundErr = True
							 ErrMsg = ErrMsg & "●在获取链接列表时出错。\n"
						  End If
					   Else
						   FoundErr = True
						  ErrMsg = ErrMsg & "●在截取列表时发生错误。\n"
					   End If
					Else
						FoundErr = True
					   ErrMsg = ErrMsg & "●在获取:" & ListUrl & "网页源码时发生错误。\n"
					End If
				 End If
			
			If FoundErr <> True Then
			   NewsCode = KMCObj.GetHttpPage(NewsUrl,CharsetCode)
			   If NewsCode <> "Error" Then
				  Title = KMCObj.GetBody(NewsCode, TsString, ToString, False, False)
				  Content = KMCObj.GetBody(NewsCode, CsString, CoString, False, False)
				  If Title = "Error" Then
					 FoundErr = True
					 ErrMsg = ErrMsg & "●在截取标题的时候发生错误：" & NewsUrl & "\n"
				  ElseIf  Content = "Error" Then
					 FoundErr = True
					 ErrMsg = ErrMsg & "●在截取正文的时候发生错误：" & NewsUrl & "\n"
				  Else
					 Title = KMCObj.FpHtmlEnCode(Title)
					 Title = KMCObj.dvHTMLEncode(Title)
			
					 '新闻分页
					' If NewsPageType = 1 Then
					'	NewsPageNext = KMCObj.GetPage(NewsCode, NPsString, NPoString, False, False)
					'	Do While NewsPageNext <> "Error"
					'	   If NewsPageStr = "" Or IsNull(NewsPageStr) = True Then
					'		  NewsPageNext = KMCObj.DefiniteUrl(NewsPageNext, NewsUrl)
					'	   Else
					'		  NewsPageNext = Replace(NewsPageStr, "{$ID}", NewsPageNext)
					'	   End If
					'	   If NewsPageNext = "" Or NewsPageNext = "Error" Then Exit Do
					'	   NewsPageNextCode = KMCObj.GetHttpPage(NewsPageNext,CharsetCode)
					'	   ContentTemp = KMCObj.GetBody(NewsPageNextCode, CsString, CoString, False, False)
					'	   If ContentTemp = "Error" Then
					'		  Exit Do
					'	   Else
					'		  Content = Content & NewsPageEnd & ContentTemp
					'		  NewsPageNext = KMCObj.GetPage(NewsPageNextCode, NPsString, NPoString, False, False)
					'	   End If
					'	Loop
					' End If
					
					
					'源代码中获取分页URL   
					If NewsPageType = 1 Then
						InfoPageStr = KMCObj.GetBody(NewsCode, NPsString, NPoString, False, False)
						If InfoPageStr = "Error" Then  '正文没有分页
							
				        Else
							 InfoPageArrayCode = KMCObj.GetArray(InfoPageStr, NewsPageStr, NewsPageEnd, False, False)
							 If InfoPageArrayCode = "Error" Then
								' FoundErr = True '为防止部分文章没有分页出错,屏蔽掉这句
								 'ErrMsg = ErrMsg & "●在分析：新闻正文分页时发生错误，请检查分页链接的开始代码和结束代码！\n"
							  Else
								 InfoPageArray = Split(InfoPageArrayCode, "$Array$")
								 If IsArray(InfoPageArray) = True Then
									For Testi = 0 To UBound(InfoPageArray)
										  InfoPageArray(Testi) = KMCObj.DefiniteUrl(InfoPageArray(Testi), NewsUrl)
									Next
									'UrlTest = InfoPageArray(0)
									'NewsCode = KMCObj.GetHttpPage(UrlTest,CharsetCode)
								 Else
									FoundErr = True
									ErrMsg = ErrMsg & "●在分析：" & NewsUrl & "新闻列表时发生错误！\n"
								 End If
							  End If
						End if
			        End If
			     			
					 If UpDateType = 0 Then
						UpDateTime = Now()
					 ElseIf UpDateType = 1 Then
						If DateType = 0 Then
						   UpDateTime = Now()
						Else
						   UpDateTime = KMCObj.GetBody(NewsCode, DsString, DoString, False, False)
						   'UpDateTime = KMCObj.FpHtmlEnCode(UpDateTime)
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
					 End If
					 If Author = "Error" Or Trim(Author) = "" Then
						Author = "佚名"
					 Else
						Author = KMCObj.FpHtmlEnCode(Author)
					 End If
			
					 '来源
					 If CopyFromType = 1 Then
						CopyFrom = KMCObj.GetBody(NewsCode, FsString, FoString, False, False)
					 ElseIf CopyFromType = 2 Then
						CopyFrom = CopyFromStr
					 End If
					 If CopyFrom = "Error" Or Trim(CopyFrom) = "" Then
						CopyFrom = "不详"
					 Else
						CopyFrom = KMCObj.FpHtmlEnCode(KS.ScriptHtml(CopyFrom, "A", 3))
					 End If
					 If KeyType = 0 Then
						Key = Title
						Key = KMCObj.CreateKeyWord(Key, 2)
					 ElseIf KeyType = 1 Then
						Key = KMCObj.GetBody(NewsCode, KsString, KoString, False, False)
						Key = KMCObj.FpHtmlEnCode(Key)
						'Key = KMCObj.CreateKeyWord(Key, 2)
					 ElseIf KeyType = 2 Then
						Key = KMCObj.FpHtmlEnCode(Key)
					 End If
					 If Key = "Error" Or Trim(Key) = "" Then
						Key = ""
					 End If
				 End If
			   Else
				 FoundErr = True
				 ErrMsg = ErrMsg & "●在获取源码时发生错误：" & NewsUrl & "\n"
			   End If
			End If
			
			If FoundErr <> True Then
			   Call GetFilters
			   Call Filters
			   Content = KMCObj.ReplaceSaveRemoteFile(UploadFiles, Content, strInstallDir, strChannelDir, False, NewsUrl)
			End If
			
			End Sub
			
			
			'==================================================
			'过程名：GetFilters
			'作  用：提取过滤信息
			'参  数：无
			'==================================================
			Sub GetFilters()
			   SqlF = "Select * From KS_Filters Where Flag=True And (PublicTf=True Or ItemID=" & ItemID & ") order by FilterID ASC"
			   Set RsF = ConnItem.Execute(SqlF)
			   If RsF.EOF And RsF.BOF Then
				  Arr_Filters = ""
			   Else
				  Arr_Filters = RsF.GetRows()
			   End If
			   RsF.Close
			   Set RsF = Nothing
			End Sub
			
			
			'==================================================
			'过程名：Filters
			'作  用：过滤
			'==================================================
			Sub Filters()
			If IsArray(Arr_Filters) = False Then
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
End Class
%> 

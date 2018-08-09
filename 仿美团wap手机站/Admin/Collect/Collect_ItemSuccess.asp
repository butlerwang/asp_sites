<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.CollectCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Collect_ItemSuccess
KSCls.Kesion()
Set KSCls = Nothing

Class Collect_ItemSuccess
        Private KS
		Private KMCObj
		Private ConnItem
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
		Dim SqlItem, RsItem, FoundErr, ErrMsg
		Dim ItemID, ItemName, ChannelID, strChannelDir, ClassID, SpecialID, PaginationType, MaxCharPerPage, ReadLevel
		Dim Stars, ReadPoint, Hits, UpDateType, UpDateTime, PicNews, Rolls, Comment, Recommend, Popular, FnameType, TemplateID
		Dim Script_Iframe, Script_Object, Script_Script, Script_Div, Script_Class, Script_Span, Script_Img, Script_Font, Script_A, Script_Html, Script_Table, Script_Tr, Script_Td
		Dim CollecListNum, CollecNewsNum, RepeatInto,IntoBase, BeyondSavePic, CollecOrder, Verific, InputerType, Inputer, EditorType, Editor, ShowComment
		Dim tClass, tSpecial
		FoundErr = False
		ItemID = Request("ItemID")
		
		If ItemID = "" Then
		   FoundErr = True
		   ErrMsg = ErrMsg & "●参数错误，请从有效链接进入\n"
		Else
		   ItemID = CLng(ItemID)
		End If
		
		If FoundErr <> True Then
		   ItemName = Trim(Request.Form("ItemName"))
		   ChannelID = Trim(Request.Form("ChannelID"))
		   ClassID = Trim(Request.Form("ClassID"))
		   SpecialID = Trim(Request.Form("SpecialID"))
		   PaginationType = Trim(Request.Form("PaginationType"))
		   MaxCharPerPage = Trim(Request.Form("MaxCharPerPage"))
		   ReadLevel = Trim(Request.Form("ReadLevel"))
		   Stars = Trim(Request.Form("Stars"))
		   ReadPoint = Trim(Request.Form("ReadPoint"))
		   Hits = Trim(Request.Form("Hits"))
		   UpDateType = Trim(Request.Form("UpdateType"))
		   UpDateTime = Trim(Request.Form("UpDateTime"))
		   PicNews = Trim(Request.Form("PicNews"))
		   Rolls = Trim(Request.Form("Rolls"))
		   Comment = Trim(Request.Form("Comment"))
		   Recommend = Trim(Request.Form("Recommend"))
		   Popular = Trim(Request.Form("Popular"))
		   FnameType = Trim(Request.Form("FnameType"))
		   TemplateID = Trim(Request.Form("TemplateID"))
		   Script_Iframe = Trim(Request.Form("Script_Iframe"))
		   Script_Object = Trim(Request.Form("Script_Object"))
		   Script_Script = Trim(Request.Form("Script_Script"))
		   Script_Div = Trim(Request.Form("Script_Div"))
		   Script_Class = Trim(Request.Form("Script_Class"))
		   Script_Span = Trim(Request.Form("Script_Span"))
		   Script_Img = Trim(Request.Form("Script_Img"))
		   Script_Font = Trim(Request.Form("Script_Font"))
		   Script_A = Trim(Request.Form("Script_A"))
		   Script_Html = Trim(Request.Form("Script_Html"))
		   CollecListNum = Trim(Request.Form("CollecListNum"))
		   CollecNewsNum = Trim(Request.Form("CollecNewsNum"))
		   IntoBase = Trim(Request.Form("IntoBase"))
		   BeyondSavePic = Trim(Request.Form("BeyondSavePic"))
		   CollecOrder = Trim(Request.Form("CollecOrder"))
		   Verific = Trim(Request.Form("Verific"))
		   InputerType = Trim(Request.Form("InputerType"))
		   Inputer = Trim(Request.Form("Inputer"))
		   RepeatInto=KS.ChkClng(Request.Form("RepeatInto"))
		   EditorType = Trim(Request.Form("EditorType"))
		   Editor = Trim(Request.Form("Editor"))
		   ShowComment = Trim(Request.Form("ShowComment"))
		   Script_Table = Trim(Request.Form("Script_Table"))
		   Script_Tr = Trim(Request.Form("Script_Tr"))
		   Script_Td = Trim(Request.Form("Script_Td"))
		   If ItemName = "" Then
			  FoundErr = True
			  ErrMsg = ErrMsg & "●项目名称不能为空！\n"
		   End If
		   If ChannelID = "" Or ChannelID = 0 Then
			  FoundErr = True
			  ErrMsg = ErrMsg & "●请选择项目所属频道！\n"
		   Else
			  ChannelID = CLng(ChannelID)
		   End If
		   If ClassID = "" Then
			  FoundErr = True
			  ErrMsg = ErrMsg & "●请选择项目所属栏目！\n"
		   Else
			  Set tClass = conn.Execute("select * From KS_Class where ID='" & ClassID & "'")
			  strChannelDir = tClass("Folder")
			  Set tClass = Nothing
		   End If
		   If SpecialID = "" Then
			  SpecialID = 0
		   Else
			  SpecialID = SpecialID
			  If SpecialID <> 0 Then
				 Set tSpecial = conn.Execute("select ID From KS_Special Where FolderID='" & ClassID & "'")
				 If tSpecial.BOF And tSpecial.EOF Then
					FoundErr = True
					ErrMsg = ErrMsg & "●在本频道内找不到指定的专题\n"
				 End If
				 Set tSpecial = Nothing
			  End If
		   End If
		
		   If PaginationType = "" Then
			  FoundErr = True
			  ErrMsg = ErrMsg & "●请选择分页类型！\n"
		   Else
			  PaginationType = CLng(PaginationType)
		   End If
		   
		   If MaxCharPerPage = "" And PaginationType = 1 Then
			  FoundErr = True
			  ErrMsg = ErrMsg & "●请输入每页字符数！\n"
		   Else
			  MaxCharPerPage = CLng(MaxCharPerPage)
		   End If
		   
		   
		   
		   If Stars = "" Then
			  FoundErr = True
			  ErrMsg = ErrMsg & "●请选择文章评分等级！\n"
		   Else
			  Stars = Stars
		   End If
		   
			 
		   If ReadPoint = "" Then
			  ReadPoint = 0
		   Else
			  ReadPoint = CLng(ReadPoint)
		   End If
		  
		   If Hits = "" Then
			  Hits = 0
		   Else
			  Hits = CLng(Hits)
		   End If
		  
		   If UpDateType = "" Then
			  FoundErr = True
			  ErrMsg = ErrMsg & "●请选择文章录入时间类型！\n"
		   Else
			  UpDateType = CLng(UpDateType)
			  If UpDateType = 2 Then
				 If IsDate(UpDateTime) = False Then
					FoundErr = True
					ErrMsg = ErrMsg & "●文章录入时间格式不正确！\n"
				 Else
					UpDateTime = CDate(UpDateTime)
				 End If
			  End If
		   End If
		   If FnameType = "" Then
			  FoundErr = True
			  ErrMsg = ErrMsg & "●请选择生成的扩展名！\n"
		   Else
			  FnameType = FnameType
		   End If
		  ' If TemplateID = "" Then
		  '	  FoundErr = True
		  '  ErrMsg = ErrMsg & "●请选择绑定的模板！\n"
		  ' End If
		   If CollecListNum = "" Then
			  FoundErr = True
			  ErrMsg = ErrMsg & "●请填写新闻列表深度！\n"
		   Else
			  CollecListNum = CLng(CollecListNum)
		   End If
		   If CollecNewsNum = "" Then
			  FoundErr = True
			  ErrMsg = ErrMsg & "●请填写新闻采集数量！\n"
		   Else
			  CollecNewsNum = CLng(CollecNewsNum)
		   End If
		
		   If InputerType = "" Then
			  FoundErr = True
			  ErrMsg = ErrMsg & "●请选择新闻录入者类型！\n"
		   Else
			  InputerType = CLng(InputerType)
			  If InputerType = 1 And Inputer = "" Then
				 FoundErr = True
				 ErrMsg = ErrMsg & "●请填写新闻录入者！\n"
			  End If
		   End If
		
		   If EditorType = "" Then
			  FoundErr = True
			  ErrMsg = ErrMsg & "●请选择责任编辑类型！\n"
		   Else
			  EditorType = CLng(EditorType)
			  If EditorType = 1 And Editor = "" Then
				 FoundErr = True
				 ErrMsg = ErrMsg & "●请填写责任编辑！\n"
			  End If
		   End If
		End If
		If FoundErr <> True Then
		   SqlItem = "Select * From KS_CollectItem Where ItemID=" & ItemID
		   Set RsItem = Server.CreateObject("adodb.recordset")
		   RsItem.Open SqlItem, ConnItem, 2, 3
		   If RsItem.EOF And RsItem.BOF Then
			  FoundErr = True
			  ErrMsg = ErrMsg & "●参数错误，没有找该项目\n"
		   Else
			  RsItem("ItemName") = ItemName
			  RsItem("ChannelID") = ChannelID
			  RsItem("ChannelDir") = strChannelDir
			  RsItem("ClassID") = ClassID
			  RsItem("SpecialID") = SpecialID
			  RsItem("PaginationType") = PaginationType
			  RsItem("MaxCharPerPage") = MaxCharPerPage
			  RsItem("ReadLevel") = ReadLevel
			  RsItem("Stars") = Stars
			  RsItem("ReadPoint") = ReadPoint
			  RsItem("Hits") = Hits
			  RsItem("UpdateType") = UpDateType
			  RsItem("RepeatInto") = RepeatInto
			  If UpDateType = 2 Then
				 RsItem("UpDateTime") = UpDateTime
			  End If
			  If PicNews = "" Then
				 RsItem("PicNews") = 0
			  Else
				 RsItem("PicNews") = 1
			  End If
			  If Rolls = "" Then
				 RsItem("Rolls") = 0
			  Else
				 RsItem("Rolls") = 1
			  End If
			  If Comment = "" Then
				 RsItem("Comment") = 0
			  Else
				 RsItem("Comment") = 1
			  End If
			  If Recommend = "" Then
				 RsItem("Recommend") = 0
			  Else
				 RsItem("Recommend") = 1
			  End If
			  If Popular = "" Then
				 RsItem("Popular") = 0
			  Else
				 RsItem("Popular") = 1
			  End If
			  RsItem("FnameType") = FnameType
			  RsItem("TemplateID") = TemplateID
			  If Script_Iframe = "yes" Then
				 RsItem("Script_Iframe") = -1
			  Else
				 RsItem("Script_Iframe") = 0
			  End If
			  If Script_Object = "yes" Then
				 RsItem("Script_Object") = -1
			  Else
				 RsItem("Script_Object") = 0
			  End If
			  If Script_Script = "yes" Then
				 RsItem("Script_Script") = -1
			  Else
				 RsItem("Script_Script") = 0
			  End If
			  If Script_Div = "yes" Then
				 RsItem("Script_Div") = -1
			  Else
				 RsItem("Script_Div") = 0
			  End If
			  If Script_Class = "yes" Then
				 RsItem("Script_Class") = -1
			  Else
				 RsItem("Script_Class") = 0
			  End If
			  If Script_Span = "yes" Then
				 RsItem("Script_Span") = -1
			  Else
				 RsItem("Script_Span") = 0
			  End If
			  If Script_Img = "yes" Then
				 RsItem("Script_Img") = -1
			  Else
				 RsItem("Script_Img") = 0
			  End If
		
			  If Script_Font = "yes" Then
				 RsItem("Script_Font") = -1
			  Else
				 RsItem("Script_Font") = 0
			  End If
			  If Script_A = "yes" Then
				 RsItem("Script_A") = -1
			  Else
				 RsItem("Script_A") = 0
			  End If
		
			  If Script_Html = "yes" Then
				 RsItem("Script_Html") = -1
			  Else
				 RsItem("Script_Html") = 0
			  End If
			  RsItem("CollecListNum") = CollecListNum
			  RsItem("CollecNewsNum") = CollecNewsNum
			  RsItem("IntoBase") = IntoBase
			  If BeyondSavePic = "" Then
				 RsItem("BeyondSavePic") = 0
			  Else
				 RsItem("BeyondSavePic") = 1
			  End If
			  If CollecOrder = "yes" Then
				 RsItem("CollecOrder") = -1
			  Else
				 RsItem("CollecOrder") = 0
			  End If
			  'If Verific = "" Then
			'	 RsItem("Verific") = 0
			  'Else
				 RsItem("Verific") = 1
			 ' End If
			  RsItem("InputerType") = InputerType
			  If InputerType = 1 Then
				 RsItem("Inputer") = Inputer
			  End If
			  RsItem("EditorType") = EditorType
			  If EditorType = 1 Then
				 RsItem("Editor") = Editor
			  End If
			  RsItem("ShowComment") = ShowComment
			  RsItem("Flag") = True
			  If Script_Table = "yes" Then
				 RsItem("Script_Table") = -1
			  Else
				 RsItem("Script_Table") = 0
			  End If
			  If Script_Tr = "yes" Then
				 RsItem("Script_Tr") = -1
			  Else
				 RsItem("Script_Tr") = 0
			  End If
			  If Script_Td = "yes" Then
				 RsItem("Script_Td") = -1
			  Else
				 RsItem("Script_Td") = 0
			  End If
		
		   End If
		   ErrMsg = ItemName & "项目设置完成！"
		   RsItem.Update
		   RsItem.Close
		   Set RsItem = Nothing
		End If
		
		If FoundErr = True Then
		   Call KS.AlertHistory(ErrMsg,-1)
		Else
		   Call KS.Alert(ErrMsg,"Collect_Main.asp?channelid=" & channelid)
		End If
		
		End Sub
End Class
%> 

<%

'---------------------------------------
'定义刷新临时变量保存通用类
'---------------------------------------
Dim FCls
Set FCls=New GlobalFsoCls
Class GlobalFsoCls
        Private M_RefreshTemplateID,M_RefreshTempFileContent,M_RefreshCurrTid
		Private M_PageList,M_PerPageNum,M_PageParam
		private M_PageStyle,M_ItemUnit,M_TotalPage,M_FsoListNum,M_TotalPut,M_ItemTitle
		private M_ChannelID,M_CallFrom3G
		private M_RefreshType,M_RefreshChannelHomeFlag,M_RefreshFolderID,M_RefreshInfoID,M_CurrSpecialID,M_BrandName,M_RefreshParentID,M_LocationStr
		private M_FromAspPage
		Private Sub Class_Initialize()
		End Sub
        Private Sub Class_Terminate()
		 Set FCls=Nothing
		End Sub
		
		'初始化栏目相关信息
		Public Sub SetClassInfo(S_ChannelID,S_RefreshFolderID,S_ParentID)
		   	RefreshType="Folder"
			ChannelID=S_ChannelID
			RefreshFolderID=S_RefreshFolderID
			RefreshParentID=S_ParentID
			If RefreshParentID="0" Then
			 RefreshChannelHomeFlag=true
			Else
			 RefreshChannelHomeFlag=false
			End If
		End Sub
		
		'初始化内容页相关信息
		Public Sub SetContentInfo(S_ChannelID,S_RefreshFolderID,S_RefreshInfoID,S_ItemTitle)
		   	RefreshType    = "Content"
			ChannelID      = S_ChannelID
			RefreshFolderID= S_RefreshFolderID
			RefreshInfoID  = S_RefreshInfoID
			ItemTitle      = S_ItemTitle
		End Sub
		
		'初始化专题信息
		Public Sub SetSpecialInfo(S_RefreshFolderID,S_CurrSpecialID)
		    RefreshType="Special"
			RefreshFolderID=S_RefreshFolderID
			CurrSpecialID=S_CurrSpecialID
		End Sub
		
		'====================模型ID================
		Public Property Let ChannelID(ByVal strVar) 
		M_ChannelID = strVar 
		End Property 
		
		Public Property Get ChannelID
		ChannelID= M_ChannelID
		End Property 
		'================================================
		'====================待分页内容================
		Public Property Let PageList(ByVal strVar) 
		M_PageList = strVar 
		End Property 
		
		Public Property Get PageList
		PageList= M_PageList
		End Property 
		'================================================
		'====================待生成页数================
		Public Property Let FsoListNum(ByVal strVar) 
		M_FsoListNum = strVar 
		End Property 
		
		Public Property Get FsoListNum
		 If M_FsoListNum="" Then
		  FsoListNum=0
		 Else
 		  FsoListNum= M_FsoListNum
		 End If
		End Property 
		'================================================
		'====================总页数================
		Public Property Let TotalPage(ByVal strVar) 
		M_TotalPage = strVar 
		End Property 
		
		Public Property Get TotalPage
		TotalPage= M_TotalPage
		End Property 
		'================================================
		'====================总记录数================
		Public Property Let TotalPut(ByVal strVar) 
		M_TotalPut = strVar 
		End Property 
		
		Public Property Get TotalPut
		TotalPut= M_TotalPut
		End Property 
		'================================================
		
		
		'====================项目单位================
		Public Property Let ItemUnit(ByVal strVar) 
		M_ItemUnit = strVar 
		End Property 
		
		Public Property Get ItemUnit
		ItemUnit= M_ItemUnit
		End Property 
		'================================================
		
		'====================项目标题名称================
		Public Property Let ItemTitle(ByVal strVar) 
		M_ItemTitle = strVar 
		End Property 
		
		Public Property Get ItemTitle
		ItemTitle= M_ItemTitle
		End Property 
		'================================================
		
		
		'====================每页显示条数================
		Public Property Let PerPageNum(ByVal strVar) 
		M_PerPageNum = strVar 
		End Property 
		
		Public Property Get PerPageNum
		PerPageNum= M_PerPageNum
		End Property 
		'================================================
		
		'====================分页样式================
		Public Property Let PageStyle(ByVal strVal)
		 M_PageStyle=strVal
		End Property
		Public Property Get PageStyle
		PageStyle= M_PageStyle
		End Property 
		'================================================
		'====================分页参数================
		Public Property Let PageParam(ByVal strVal)
		 M_PageParam=strVal
		End Property
		Public Property Get PageParam
		PageParam= M_PageParam
		End Property 
		'================================================



 		'====================刷新类型================
		Public Property Let RefreshType(ByVal strVar) 
		M_RefreshType = strVar 
		End Property 
		
		Public Property Get RefreshType
		RefreshType= M_RefreshType
		End Property 
		'================================================
 		'====================刷新的当前栏目ID================
		Public Property Let RefreshFolderID(ByVal strVar) 
		M_RefreshFolderID = strVar 
		End Property 
		
		Public Property Get RefreshFolderID
		RefreshFolderID= M_RefreshFolderID
		End Property 
		'================================================
 		'====================刷新的当前栏目父ID================
		Public Property Let RefreshParentID(ByVal strVar) 
		M_RefreshParentID = strVar 
		End Property 
		
		Public Property Get RefreshParentID
		RefreshParentID= M_RefreshParentID
		End Property 
		'================================================
 		'====================刷新的当前内容ID================
		Public Property Let RefreshInfoID(ByVal strVar) 
		M_RefreshInfoID = strVar 
		End Property 
		
		Public Property Get RefreshInfoID
		RefreshInfoID= M_RefreshInfoID
		End Property 
		'================================================
 		'====================刷新的当前模板名称================
		Public Property Let RefreshTemplateID(ByVal strVar) 
		M_RefreshTemplateID = strVar 
		End Property 
		
		Public Property Get RefreshTemplateID
		RefreshTemplateID= M_RefreshTemplateID
		End Property 
		'================================================
 		'====================刷新的当前模板内容================
		Public Property Let RefreshTempFileContent(ByVal strVar) 
		M_RefreshTempFileContent = strVar 
		End Property 
		
		Public Property Get RefreshTempFileContent
		RefreshTempFileContent= M_RefreshTempFileContent
		End Property 
		'================================================
 		'====================刷新的当前栏目ID,标签模板是不是从缓存里读取================
		Public Property Let RefreshCurrTid(ByVal strVar) 
		M_RefreshCurrTid = strVar 
		End Property 
		
		Public Property Get RefreshCurrTid
		RefreshCurrTid= M_RefreshCurrTid
		End Property 
		'================================================
 		'====================刷新的当前专题ID================
		Public Property Let CurrSpecialID(ByVal strVar) 
		M_CurrSpecialID = strVar 
		End Property 
		
		Public Property Get CurrSpecialID
		CurrSpecialID= M_CurrSpecialID
		End Property 
		'================================================
 		'====================刷新频道首页标志================
		Public Property Let RefreshChannelHomeFlag(ByVal strVar) 
		M_RefreshChannelHomeFlag = strVar 
		End Property 
		
		Public Property Get RefreshChannelHomeFlag
		RefreshChannelHomeFlag= M_RefreshChannelHomeFlag
		End Property 
		'================================================
 		'====================标志专题是否生成================
		Public Property Let FromAspPage(ByVal strVar) 
		M_FromAspPage = strVar 
		End Property 
		
		Public Property Get FromAspPage
		FromAspPage= M_FromAspPage
		End Property 
		'================================================
 		'====================商城品牌名称================
		Public Property Let BrandName(ByVal strVar) 
		M_BrandName = strVar 
		End Property 
		
		Public Property Get BrandName
		BrandName= M_BrandName
		End Property 
		'================================================
 		'====================位置导航临时标记================
		Public Property Let LocationStr(ByVal strVar) 
		M_LocationStr = strVar 
		End Property 
		
		Public Property Get LocationStr
		LocationStr= M_LocationStr
		End Property 
		'================================================
 		'====================3G页面调用标记================
		Public Property Let CallFrom3G(ByVal strVar) 
		M_CallFrom3G = strVar 
		End Property 
		
		Public Property Get CallFrom3G
		CallFrom3G= M_CallFrom3G
		End Property 
		'================================================
		
 
End Class
%>
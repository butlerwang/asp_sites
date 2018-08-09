<%
'***********************************************************************************************************
'过程名：CreateXML
'作  用：生成自定义XML文档
'参  数：ID ----  标签ID号
'返回值：无
'*************************************************************************************************************
Sub CreateXML(id)
      Dim KS:Set KS=New PublicCls
	  Dim Param:Param=" Where LabelType=7"
	  If ID<>"" Then Param=Param & " And ID='" & ID &"'"
	  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
	  RS.Open "select * from KS_Label" & Param,conn,1,1
	  Dim CreatePath
	  If KS.Setting(127)<>"/" Then
	  CreatePath=KS.Setting(3) & KS.Setting(127)
	  Call KS.CreateListFolder(CreatePath)
	  Else
	  CreatePath=KS.Setting(3)
	  End If
	  Dim Content,PArr
	  Dim DCls:Set Dcls=New DIYCls
	  Do While Not RS.Eof 
	     CreatePath = CreatePath & RS("FileName")
		 PArr       = Split(RS("Description"),"@@@")
		 Content    = RS("LabelContent")
		 DCls.DataSourceType = KS.ChkClng(PArr(6))
		 DCls.DataSourceStr  = PArr(7)
		 if DCls.DataSourceType=1 Or DCls.DataSourceType=5 Or DCls.DataSourceType=6 then DCls.DataSourceStr=LFCls.GetAbsolutePath(DCls.DataSourceStr)
		 If DCls.DataSourceType<>0 Then
		  If DCls.OpenExtConn=false Then KS.AlertHintScript "标签[" & RS("LabelName") & "]的外部数据源连接出错,请检查！"
		 End If
		 Content    = DCls.ExecSQL(PArr(0),Content)
		 Call KS.WriteTOFile(CreatePath, Content)
	   RS.MoveNext
	  Loop
	  Set DCls=nothing
	  RS.Close:Set RS=Nothing
	  Set KS=Nothing
End Sub
	
'***********************************************************************************************************
'函数名：ReturnLabelFolderTree
'作  用：显示标签目录列表。
'参  数：SelectID ----  默认目录树ID号,ChannelID频道ID号,FolderType目录类型 0系统函数标签目录,1自由标签目录
'返回值：标签目录列表
'*************************************************************************************************************
Public Function ReturnLabelFolderTree(SelectID, FolderType)
		   Dim TempStr,ID,FolderName
		   SelectID = Trim(SelectID)
		   If FolderType = "" Then FolderType = 0
		   TempStr = "<select class='textbox' style='width:200;border-style: solid; border-width: 1' name='ParentID'>"
		   
		   TempStr = TempStr & "<option value='0' Selected>根目录</option>"
			Dim RS:Set RS=Conn.Execute("Select ID,FolderName from KS_LabelFolder Where FolderType=" & FolderType & " And ParentID='0' Order By AddDate desc")
			
			Do While Not RS.EOF
			   ID = Trim(RS(0))
			   FolderName = Trim(RS(1))
			   TempStr = TempStr & "<option  "
			   If SelectID = ID Then TempStr = TempStr & " Selected"
			   TempStr = TempStr & " value='" & ID & "'>" & FolderName & " </option>"
			   TempStr = TempStr & ReturnSubLabelFolderTree(ID, SelectID)
			RS.MoveNext
			Loop
			RS.Close:Set RS = Nothing
			TempStr = TempStr & "</select>"
			ReturnLabelFolderTree = TempStr
End Function
	
'************************************************************************************
'函数名：ReturnSubLabelFolderTree
'作  用：查找并返子树数据。
'参  数：ParentID ----父节点ID,   FolderID ----选择项ID
'返回值：标签目录子树列表
'************************************************************************************
Public Function ReturnSubLabelFolderTree(ParentID, FolderID)
	  Dim SubTypeList, SubRS, SpaceStr, k, Total, Num,FolderName, ID,TJ
	  
	  Set SubRS = Server.CreateObject("ADODB.RECORDSET")
	  SubRS.Open ("Select count(ID) AS total from KS_LabelFolder Where ParentID='" & ParentID & "'"), Conn, 1, 1
	  Total = SubRS("Total")
	  SubRS.Close
	  SubRS.Open ("Select ID,FolderName,TS from KS_LabelFolder Where ParentID='" & ParentID & "' Order BY AddDate Desc"), Conn, 1, 1
	  Num = 0
	  Do While Not SubRS.EOF
	   Num = Num + 1:SpaceStr = ""
		TJ = UBound(Split(SubRS(2), ","))
		For k = 1 To TJ - 1
		  If k = 1 And k <> TJ - 1 Then
		  SpaceStr = SpaceStr & "&nbsp;&nbsp;│"
		  ElseIf k = TJ - 1 Then
			If Num = Total Then
				 SpaceStr = SpaceStr & "&nbsp;&nbsp;└ "
			Else
				 SpaceStr = SpaceStr & "&nbsp;&nbsp;├ "
			End If
		  Else
		   SpaceStr = SpaceStr & "&nbsp;&nbsp;│"
		  End If
		Next
	  ID = Trim(SubRS(0))
	  FolderName = Trim(SubRS(1))
	  If FolderID = ID Then
	   SubTypeList = SubTypeList & "<option selected value='" & ID & "'>" & SpaceStr & FolderName & "</option>"
	  Else
	   SubTypeList = SubTypeList & "<option value='" & ID & "'>" & SpaceStr & FolderName & "</option>"
	  End If
	   SubTypeList = SubTypeList & ReturnSubLabelFolderTree(ID, FolderID)
	  SubRS.MoveNext
	 Loop
	  SubRS.Close:Set SubRS = Nothing:ReturnSubLabelFolderTree = SubTypeList
End Function
'***********************************************************************************************************
'函数名：ReturnLabelInfo
'参  数：LabelName ----  默认标签名称,FolderID---标签目录ID号,Descript---标签描述
'返回值：标签基本信息
'*************************************************************************************************************
Public Function ReturnLabelInfo(LabelName, FolderID, Descript)
	  ReturnLabelInfo = ReturnLabelInfo & ("        <table width=""98%"" border='0' align='center' cellpadding='2' cellspacing='1' class='border' style='margin-top:6px'>")
	  ReturnLabelInfo = ReturnLabelInfo & ("          <tr  height=""26"" class=""title""><td colspan=2 align=center><strong>")
	  If Request("labelid")="" Then
	  ReturnLabelInfo = ReturnLabelInfo & ("创 建 新 标 签")
	  Else
	  ReturnLabelInfo = ReturnLabelInfo & (" 修 改 标 签 属 性")
	  End If
	  ReturnLabelInfo = ReturnLabelInfo & ("</strong></td>")
	  ReturnLabelInfo = ReturnLabelInfo & ("          </tr>")
	  ReturnLabelInfo = ReturnLabelInfo & ("          <tr class=tdbg>")
	  ReturnLabelInfo = ReturnLabelInfo & ("      <td  colspan=2 height=""30"">标签名称")
	  ReturnLabelInfo = ReturnLabelInfo & ("        <input name=""LabelName"" size='35' class=""textbox"" type=""text"" id=""LabelName"" value=""" & LabelName & """>")
	  ReturnLabelInfo = ReturnLabelInfo & ("        <font color=""#FF0000""> * 调用格式""{LB_标签名称}""</font></td>")
	  ReturnLabelInfo = ReturnLabelInfo & ("    </tr>")
	  ReturnLabelInfo = ReturnLabelInfo & ("    <tr class=tdbg>")
	  ReturnLabelInfo = ReturnLabelInfo & ("      <td  colspan=2 height=""30"">标签目录 " & ReturnLabelFolderTree(FolderID, 0) & "<font color=""#FF0000""> 请选择标签归属目录，以便日后管理标签</font></td>")
	  ReturnLabelInfo = ReturnLabelInfo & ("    </tr>")
	  ReturnLabelInfo = ReturnLabelInfo & ("    <tr class=tdbg style='display:none'>")
	  ReturnLabelInfo = ReturnLabelInfo & ("      <td  colspan=2 height=""30"">标签描述")
	  ReturnLabelInfo = ReturnLabelInfo & ("        <input name=""Descript"" class=""textbox"" type=""text"" id=""Descript"" value=""" & Descript & """ size=""40"">")
	  ReturnLabelInfo = ReturnLabelInfo & ("        <font color=""#FF0000""> 请在此输入标签的说明,方便以后查找</font></td>")
	  ReturnLabelInfo = ReturnLabelInfo & ("    </tr>")
	 ' ReturnLabelInfo = ReturnLabelInfo & ("    </table>")
End Function
	'**************************************************
	'函数名：ReturnDateFormat
	'作  用：返回系统支持的日期格式
	'参  数：SelectDate 预定选中的日期格式
	'**************************************************
	Public Function ReturnDateFormat(SelectDate)
	         TempFormatDateStr=" <input name=""DateRule"" class=""textbox"" type=""text"" id=""DateRule"" style=""width:100px;"" value=""" & SelectDate & """ >"
			 
			 Dim TempFormatDateStr, Str
				TempFormatDateStr = TempFormatDateStr & (" <select onchange=""$('#DateRule').val(this.value)"" name='sdate' style='width:120px;'><option value='' Selected>=快速选择日期=</option><option value="""">-不显示日期-</option> ")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""YYYY-MM-DD"">2005-10-1</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""YYYY.MM.DD"">2005.10.1</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""YYYY/MM/DD"">2005/10/1</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""MM/DD/YYYY"">10/1/2005</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""YYYY年MM月"">2005年10月</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""YYYY年MM月DD日"">2005年10月1日</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""MM.DD.YYYY"">10.1.2005</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""MM-DD-YYYY"">10-1-2005</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""MM/DD"">10/1</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""MM.DD"">10.1</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""MM月DD日"">10月1日</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""DD日hh时"">1日12时</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""DD日hh点"">1日12点</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""hh时mm分>12时12分</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""hh:mm"">12:12</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""MM-DD"">10-1</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""MM/DD hh:mm"">10/1 12:00</option>")
			  
			  TempFormatDateStr = TempFormatDateStr & ("<optgroup  label=""---加括号格式---""></optgroup>")

			   TempFormatDateStr = TempFormatDateStr & ("<option value=""(YYYY-MM-DD)"" >(2005-10-1)</option>") 
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""(YYYY.MM.DD)"">(2005.10.1)</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""(YYYY/MM/DD)"">(2005/10/1)</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""(MM/DD/YYYY)"">(10/1/2005)</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""(YYYY年MM月)"">(2005年10月)</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""(YYYY年MM月DD日)"">(2005年10月1日)</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""(MM.DD.YYYY)"">(10.1.2005)</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""(MM-DD-YYYY)"">(10-1-2005)</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""(MM/DD)"">(10/1)</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""(MM.DD)"">(10.1)</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""(MM月DD日)"">(10月1日)</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""(DD日hh时)"">(1日12时)</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""(DD日hh点)"">(1日12点)</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""(hh时mm分)"">(12时12分)</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""(hh:mm)"">(12:12)</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""(MM-DD)"">(10-1)</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""(MM/DD hh:mm)"">(10/1 12:00)</option>")


			  TempFormatDateStr = TempFormatDateStr & ("<optgroup  label=""---加中括号格式---""></optgroup>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""[YYYY-MM-DD]"">[2005-10-1]</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""[YYYY.MM.DD]"">[2005.10.1]</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""[YYYY/MM/DD]"">[2005/10/1]</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""[MM/DD/YYYY]"">[10/1/2005]</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""[YYYY年MM月]"">[2005年10月]</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""[YYYY年MM月DD日]"">[2005年10月1日]</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""[MM.DD.YYYY]"">[10.1.2005]</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""[MM-DD-YYYY]"">[10-1-2005]</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""[MM/DD]"">[10/1]</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""[MM.DD]"">[10.1]</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""[MM月DD日]"">[10月1日]</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""[DD日hh时]"">[1日12时]</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""[DD日hh点]"">[1日12点]</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""[hh时mm分]"">[12时12分]</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""[hh:mm]"">[12:12]</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""[MM-DD]"">[10-1]</option>")
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""[MM/DD hh:mm]"">[10/1 12:00]</option></select><br/><font color=green>支持标签：YYYY:年(4位) YY:年(2位) MM:月&nbsp;DD:日 hh:时 mm:分 ss:秒</font>")
			ReturnDateFormat = TempFormatDateStr
	End Function
	
	'**************************************************
	'函数名：ReturnOpenTypeStr
	'作  用：返回系统支持的打开窗口方式(带可输入的下拉框)
	'参  数：SelectValue 预定选中的链接目标
	'**************************************************
	Public Function ReturnOpenTypeStr(SelectValue)
	  ReturnOpenTypeStr = "链接目标 <select onchange=""document.getElementById('OpenType').value=this.value;"" name='sOpenType'><option value=''>-没有设置-</option>"
	  ReturnOpenTypeStr = ReturnOpenTypeStr & "<option value=""_blank""> 新窗口(_blank) </option>"
	  ReturnOpenTypeStr = ReturnOpenTypeStr & "<option value=""_parent""> 父窗口(_parent) </option>"
	  ReturnOpenTypeStr = ReturnOpenTypeStr & "<option value=""_self""> 本窗口(_self) </option>"
	  ReturnOpenTypeStr = ReturnOpenTypeStr & "<option value=""_top""> 整页(_top) </option>"
	  ReturnOpenTypeStr = ReturnOpenTypeStr & "</select>=>"
	  ReturnOpenTypeStr = ReturnOpenTypeStr & "<input type='text' class='textbox' name='OpenType' id='OpenType' size='10' value='" & SelectValue &"'>"
	  Exit Function
	End Function
	
		'****************************************************************************************************************************
	'函数名：ReturnJSInfo
	'参  数：JSID--JSID号,JSName ----    默认JS名称,JSFileName----JS文件名,FolderID---标签目录ID号,FolderType---目录类型,Descript---标签描述
	'返回值：标签基本信息
'*******************************************************************************************************************************
	Public Function ReturnJSInfo(JSID, JSName, JSFileName, FolderID, FolderType, Descript)
		 ReturnJSInfo = "<table width=""96%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
		 ReturnJSInfo = ReturnJSInfo & ("    <tr>")
		 ReturnJSInfo = ReturnJSInfo & ("       <td>")
		 ReturnJSInfo = ReturnJSInfo & ("      <FIELDSET align=center><LEGEND align=left>JS基本信息</LEGEND>")
		 ReturnJSInfo = ReturnJSInfo & ("        <table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">")
		 ReturnJSInfo = ReturnJSInfo & ("            <tr>")
		 ReturnJSInfo = ReturnJSInfo & ("             <td height=""22"" style='text-align:left'>JS 名 称")
		 ReturnJSInfo = ReturnJSInfo & ("                &nbsp;<input name=""JSName"" type=""text"" class=""textbox"" id=""JSName"" value=""" & JSName & """>")
		 ReturnJSInfo = ReturnJSInfo & ("                <font color=""#FF0000""> *</font><font color=""#FF0000""> 例如JS名称：&quot;推荐文章列表&quot;，则在模板中调用：&quot;{JS_推荐文章列表}&quot;（注意英文大小写及全半角）。</font></td>")
		 ReturnJSInfo = ReturnJSInfo & ("            </tr>")
		 ReturnJSInfo = ReturnJSInfo & ("            <tr>")
		 ReturnJSInfo = ReturnJSInfo & ("              <td height=""22"" style='text-align:left'>JS 文件名")
		 
		   If JSID <> "" Then
			  ReturnJSInfo = ReturnJSInfo & ("                <input class=""textbox"" disabled=true name=""JSFileName"" type=""text"" id=""JSFileName"" title=""JS文件名：不能带\/：*？“ < > | 等特殊符号"" value=""" & JSFileName & """>")
		   Else
			  ReturnJSInfo = ReturnJSInfo & ("                <input class=""textbox"" name=""JSFileName"" type=""text"" id=""JSFileName"" title=""JS文件名：不能带\/：*？“ < > | 等特殊符号"" value=""" & JSFileName & """>")
		   End If
		 ReturnJSInfo = ReturnJSInfo & ("            <font color=""#FF0000""> * 例如 &quot;News.js&quot; 一定要以扩展名 &quot;.js&quot;结束</font></td>")
		 ReturnJSInfo = ReturnJSInfo & ("        </tr>")
		 ReturnJSInfo = ReturnJSInfo & ("        <tr>")
		 ReturnJSInfo = ReturnJSInfo & ("         <td height=""22"" style='text-align:left'>存放目录 " & ReturnLabelFolderTree(FolderID, FolderType) & " </td>")
		 ReturnJSInfo = ReturnJSInfo & ("       </tr>")
		 ReturnJSInfo = ReturnJSInfo & ("            <tr>")
		 ReturnJSInfo = ReturnJSInfo & ("              <td height=""22"" style='text-align:left'>JS 描 述")
		 ReturnJSInfo = ReturnJSInfo & ("                <textarea class=""textbox"" style=""height:60px"" name=""Descript"" cols=""60"" rows=""4"" id=""Descript"">" & Descript & "</textarea>")
		 ReturnJSInfo = ReturnJSInfo & ("           <font color=""#FF0000""> 请在此输入JS的说明,方便以后查找</font></td>")
		 ReturnJSInfo = ReturnJSInfo & ("            </tr>")
		 ReturnJSInfo = ReturnJSInfo & ("          </table>")
		 ReturnJSInfo = ReturnJSInfo & ("        </FIELDSET></td>")
		 ReturnJSInfo = ReturnJSInfo & ("      </tr>")
		 ReturnJSInfo = ReturnJSInfo & ("   </table>")
		 
		 '采集搜索参数
		 ReturnJSInfo = ReturnJSInfo & ("<input type=""hidden"" name=""KeyWord"" value=""" & Request.QueryString("KeyWord") & """>")
		 ReturnJSInfo = ReturnJSInfo & ("<input type=""hidden"" name=""SearchType"" value=""" & Request.QueryString("SearchType") & """>")
		 ReturnJSInfo = ReturnJSInfo & ("<input type=""hidden"" name=""StartDate"" value=""" & Request.QueryString("StartDate") & """>")
		 ReturnJSInfo = ReturnJSInfo & ("<input type=""hidden"" name=""EndDate"" value=""" & Request.QueryString("EndDate") & """>")
	End Function
	
	'分页样式
	Public Function ReturnPageStyle(PageStyle)
		ReturnPageStyle = "         分页样式"
		ReturnPageStyle = ReturnPageStyle & "         <select name=""PageStyle"" id=""PageStyle"" style=""width:70%;"" class=""textbox"">"
		ReturnPageStyle = ReturnPageStyle & "          <option value=1"
		If PageStyle="1" Then ReturnPageStyle = ReturnPageStyle & " Selected"
		ReturnPageStyle = ReturnPageStyle & ">①首页 上一页 下一页 尾页</option>"
		ReturnPageStyle = ReturnPageStyle & "          <option value=2"
		If PageStyle="2" Then ReturnPageStyle = ReturnPageStyle & " Selected"
		ReturnPageStyle = ReturnPageStyle & ">②共N页/N篇 [1] [2] [3]</option>"
		ReturnPageStyle = ReturnPageStyle & "          <option value=3"
		If PageStyle="3" Then ReturnPageStyle = ReturnPageStyle & " Selected"
		ReturnPageStyle = ReturnPageStyle & ">③<< <  > >></option>"
		ReturnPageStyle = ReturnPageStyle & "          <option value=4"
		If PageStyle="4" Then ReturnPageStyle = ReturnPageStyle & " Selected"
		ReturnPageStyle = ReturnPageStyle & " style='color:blue'>④数字导航样式(新增)</option>"
		ReturnPageStyle = ReturnPageStyle & "         </select>"
	End Function
	
	'专题显示样式
	Public Function ReturnSpecialStyle(Sel)
		ReturnSpecialStyle= "显示样式&nbsp;<select name=""ShowStyle"" id=""ShowStyle"" style=""width:70%"" class=""textbox"">"
		ReturnSpecialStyle=ReturnSpecialStyle &"<option value=""1"""
		If Sel="1" Then ReturnSpecialStyle=ReturnSpecialStyle &" selected"
		ReturnSpecialStyle=ReturnSpecialStyle &">①标题式</option>"
		ReturnSpecialStyle=ReturnSpecialStyle &"<option value=""2"""
		If Sel="2" Then ReturnSpecialStyle=ReturnSpecialStyle &" selected"
		ReturnSpecialStyle=ReturnSpecialStyle &">②仅显示图片</option>"
		ReturnSpecialStyle=ReturnSpecialStyle &"<option value=""3"""
		If Sel="3" Then ReturnSpecialStyle=ReturnSpecialStyle &" Selected"
		ReturnSpecialStyle=ReturnSpecialStyle &">③图片+标题:上下</option>"
		ReturnSpecialStyle=ReturnSpecialStyle &"<option value=""4"""
		If Sel="4" Then ReturnSpecialStyle=ReturnSpecialStyle &" Selected"
		ReturnSpecialStyle=ReturnSpecialStyle &">④图片+介绍:左右</option>"
		ReturnSpecialStyle=ReturnSpecialStyle &"<option value=""5"""
		If Sel="5" Then ReturnSpecialStyle=ReturnSpecialStyle &" selected"
		ReturnSpecialStyle=ReturnSpecialStyle &">⑤图片+(名称+介绍:上下):左右</option>"
		ReturnSpecialStyle=ReturnSpecialStyle &"</select>"
	End Function
%>
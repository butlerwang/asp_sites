<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%

Dim KSCls
Set KSCls = New Link
KSCls.Kesion()
Set KSCls = Nothing

Class Link
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		 Dim Template,KSR
		 FCls.RefreshType = "LinkIndex"   '设置当前位置为友情链接首页
		 Set KSR = New Refresh
		    Template = KSR.LoadTemplate(KS.Setting(3) & KS.Setting(90) & "common/friendlink.html")
			Template = ReplaceListContent(Template)     '替换友情链接页标签为内容
			Template = KSR.KSLabelReplaceAll(Template)
		   Set KSR = Nothing
		   
		   Response.Write Template   
	End Sub
	    '*********************************************************************************************************
		'函数名：ReplaceLinkContent
		'作  用：替换友情链接页标签为内容
		'参  数：Template待替换的内容
		'*********************************************************************************************************
		Function ReplaceListContent(Template)
		   '  on error resume next
			 Dim Domain, ClassLinkStr, DetailListStr, KeyWord
			 Dim RClassID, ClassID, LinkType, ViewKind
			 Dim ObjRS:Set ObjRS=Server.CreateObject("ADODB.Recordset")
			 
			   Domain = KS.GetDomain()
			   RClassID = KS.ChkClng(KS.S("ClassID"))
			   LinkType = KS.ChkClng(KS.S("LinkType"))
			   ViewKind = KS.ChkClng(KS.S("ViewKind"))
			   KeyWord = KS.S("KeyWord")
			   IF ViewKind=0 Then ViewKind=1
			   If LinkType = 0 Then LinkType = 2
			   If Not IsNumeric(RClassID) Then
				Call KS.Alert("非法参数!", "")
				Set KS = Nothing:Exit Function
			   End If
			   
			   If InStr(Template, "{$GetLinkCommonInfo}") <> 0 Then
				  Template = Replace(Template, "{$GetLinkCommonInfo}", "<a href=""" & Domain & "plus/link/"">常规查看</a> | <a href=""" & Domain & "plus/link/?ViewKind=1"">按点击数查看</a> | <a href=""" & Domain & "plus/link/?ViewKind=2"">按类别查看</a> | <a href=""" & Domain & "plus/link/?ViewKind=3"">所有推荐站点</a> | <a href=""" & Domain & "plus/link/reg"">申请友情链接</a> ")
			   End If
			   If InStr(Template, "{$GetClassLink}") <> 0 Then
				  
				  ClassLinkStr = "<table width=""100%"" border=""0""><form action=""?"" method=""get"" name=""SearchLink""><tr><td>分类显示：<select name='LinkType' id='LinkType' onchange=""if(this.options[this.selectedIndex].value!=''){location='?LinkType='+this.options[this.selectedIndex].value;}"">"
				  If LinkType = 2 Then
				  ClassLinkStr = ClassLinkStr & "<option value='2' selected>所有类型</option>"
				  Else
				  ClassLinkStr = ClassLinkStr & "<option value='2'>所有类型</option>"
				  End If
				  If LinkType = 1 Then
				  ClassLinkStr = ClassLinkStr & "<option value='1' selected>LOGO链接</option>"
				  Else
				  ClassLinkStr = ClassLinkStr & "<option value='1'>LOGO链接</option>"
				  End If
				  If LinkType = 0 Then
				  ClassLinkStr = ClassLinkStr & "<option value='0' selected>文字链接</option>"
				  Else
				  ClassLinkStr = ClassLinkStr & "<option value='0'>文字链接</option>"
				  End If
				  ClassLinkStr = ClassLinkStr & "</select>"
				  
				  ClassLinkStr = ClassLinkStr & "&nbsp;<select name='ViewClassID' id='ViewClassID' onchange=""if(this.options[this.selectedIndex].value!=''){location='?LinkType=" & LinkType & "&ClassID='+this.options[this.selectedIndex].value;}""><option value='0'>所有分类站点</option>"
				  ObjRS.Open "Select FolderID,FolderName From KS_LinkFolder Order BY OrderID,FolderID Desc", Conn, 1, 1
				  If Not ObjRS.EOF Then
				   Do While Not ObjRS.EOF
					ClassID = ObjRS(0)
					If CStr(RClassID) = CStr(ClassID) Then
					ClassLinkStr = ClassLinkStr & "<option value='" & ClassID & "' selected>" & ObjRS(1) & "</option>"
					Else
					ClassLinkStr = ClassLinkStr & "<option value='" & ClassID & "'>" & ObjRS(1) & "</option>"
					End If
					ObjRS.MoveNext
				   Loop
				  End If
				  
				  ObjRS.Close
				  ClassLinkStr = ClassLinkStr & "</select>&nbsp;&nbsp;关键字：<input class=""textbox"" type=""text"" size=""22"" name=""KeyWord""> &nbsp;<input class=""inputbutton"" type=""submit"" value="" 搜 索 ""></td></tr></form></table>"
				  Template = Replace(Template, "{$GetClassLink}", ClassLinkStr)
			   End If
			   
			   If InStr(Template, "{$GetLinkDetail}") <> 0 Then
					Dim totalPut, CurrentPage,Para
					
				  If ViewKind = 2 Then      '按分类查看
					 If RClassID = 0 Then
					   Dim CRS:Set CRS=Server.CreateObject("ADODB.Recordset")
					   CRS.Open "Select FolderID,FolderName From KS_LinkFolder Order BY AddDate Desc", Conn, 1, 1
					   If CRS.EOF And CRS.BOF Then
						 DetailListStr = "还没有任何友情链接站点分类!"
					   Else
						 DetailListStr = "<table width=""100%"" cellpadding=""0"" cellspacing=""0"" Class=""table_border""><tr><td>"
						 DetailListStr = DetailListStr & "<table width='100%' cellSpacing=2 cellPadding=1 border=0>"
						  Do While Not CRS.EOF
							DetailListStr = DetailListStr & "<tr><td Class=""link_table_title""><a href='Index.asp?ViewKind=2&ClassID=" & CRS(0) & "'><b>" & CRS(1) & "</b></a></td></tr>"
							DetailListStr = DetailListStr & GetClassSiteList(CRS(0))
							CRS.MoveNext
						  Loop
						  DetailListStr = DetailListStr & "</table></td></tr></table>"
					   End If
						CRS.Close
						Set CRS = Nothing
					 Else
						  DetailListStr = "<table width=""100%""  cellpadding=""0"" cellspacing=""0""  Class=""table_border""><tr><td>"
						  DetailListStr = DetailListStr & "<table width='100%' cellSpacing=2 cellPadding=1 border=0>"
						  DetailListStr = DetailListStr & "<tr><td  Class=""link_table_title""><b>"
						  
						  Dim ClassRS
						  Set ClassRS = Conn.Execute("Select FolderName From KS_LinkFolder Where FolderID=" & RClassID)
						  DetailListStr = DetailListStr & ClassRS(0)
						  ClassRS.Close
						  Set ClassRS = Nothing
						  
						  DetailListStr = DetailListStr & "</b></td></tr>"
						  DetailListStr = DetailListStr & GetClassSiteList(RClassID)
						  DetailListStr = DetailListStr & "</table></td></tr></table>"
					 End If
				  Else                      '按常规等方式查看
				  
					 Const MaxPerPage = 10   '每页显示数量
					If KS.S("page") <> "" Then
					   CurrentPage = KS.ChkClng(KS.S("page"))
					Else
					  CurrentPage = 1
					End If
					
					DetailListStr = "<TABLE WIDTH=""100%""  Cellpadding=""0"" Cellspacing=""0"" Class=""table_border""><tr><td>"
					
					  Para = " Where Verific=1 And Locked=0"
					If LinkType = 0 Or LinkType = 1 Then
					  Para = Para & " And LinkType=" & LinkType
					End If
					If RClassID <> 0 Then
					  Para = Para & " And FolderID=" & RClassID
					End If
					If KeyWord <> "" Then
					  Para = Para & " And SiteName like '%" & KeyWord & "%' Or Description like '%" & KeyWord & "%'"
					End If
					If ViewKind = 3 Then
					  Para = Para & " And Recommend=1 Order By Hits Desc"
					ElseIf ViewKind = 1 Then
					  Para = Para & " Order By Hits Desc"
					Else
					  Para = Para & " Order By AddDate Desc"
					End If
					ObjRS.Open "Select * From KS_Link" & Para, Conn, 1, 1
					If ObjRS.EOF And ObjRS.BOF Then
					   If RClassID = 0 Then
						  DetailListStr = DetailListStr & "还没有加入任何友情链接!"
					   Else
						  DetailListStr = DetailListStr & "没有该类别的友情链接站点!"
					   End If
					Else
					   totalPut = ObjRS.RecordCount
							If CurrentPage < 1 Then CurrentPage = 1
							
							If CurrentPage = 1 Then
								  DetailListStr = DetailListStr & GetDetailListStr(ObjRS, totalPut, MaxPerPage, CurrentPage, RClassID)
							Else
								If (CurrentPage - 1) * MaxPerPage < totalPut Then
									ObjRS.Move (CurrentPage - 1) * MaxPerPage
								   DetailListStr = DetailListStr & GetDetailListStr(ObjRS, totalPut, MaxPerPage, CurrentPage, RClassID)
								Else
									CurrentPage = 1
								   DetailListStr = DetailListStr & GetDetailListStr(ObjRS, totalPut, MaxPerPage, CurrentPage, RClassID)
								End If
							End If
				   End If
					ObjRS.Close
					Set ObjRS = Nothing
					DetailListStr = DetailListStr & "</td></tr></table>"
			   End If
				  Template = Replace(Template, "{$GetLinkDetail}", DetailListStr)
			   End If
			   ReplaceListContent = Template
		End Function
		'结合上面ReplaceListContent函数使用
		Function GetDetailListStr(ObjRS, totalPut, MaxPerPage, CurrentPage, RClassID)
			  Dim AddDate, I, RecommendStr,LinkID
				  Do While Not ObjRS.EOF
					   AddDate = ObjRS("AddDate")
					   LinkID = ObjRS("LinkID")
					   If ObjRS("Recommend") = 1 Then
						RecommendStr = " <font color=""red"">推荐</font>"
					   Else
						RecommendStr = ""
					   End If
					   GetDetailListStr = GetDetailListStr & "<TABLE cellSpacing=1 cellPadding=4 width=100% align=center bgColor=#ffffff border=0>"
					   GetDetailListStr = GetDetailListStr & "<TR Class=""link_table_title"" height=20>"
					   If ObjRS("LinkType") = 0 Then
					   GetDetailListStr = GetDetailListStr & "<TD width=""14%""><a href=""Index.asp?LinkType=0"" title=""按文字链接查看"">文字链接</a></TD>"
					   Else
					   GetDetailListStr = GetDetailListStr & "<TD width=""14%""><a href=""Index.asp?LinkType=1"" title=""按LOGO链接查看"">LOGO链接</a></TD>"
					   End If
					   GetDetailListStr = GetDetailListStr & "<TD width=""36%""><A href = ""to?" & LinkID & """ target=""_blank"" title=""网站名称""><B>" & ObjRS("SiteName") & "</B>  " & RecommendStr & "</A></TD>"
					   GetDetailListStr = GetDetailListStr & "<TD width=""15%"">"
					   
					   on error resume next
					   Dim ClassRS:Set ClassRS = Conn.Execute("Select FolderID,FolderName From KS_LinkFolder Where FolderID=" & ObjRS("FolderID"))
					   GetDetailListStr = GetDetailListStr & "<a href=""Index.asp?ViewKind=2&ClassID=" & ClassRS(0) & """  Title=""网站类别"">" & ClassRS(1) & "</a>"
					   ClassRS.Close:Set ClassRS = Nothing
					   
					   GetDetailListStr = GetDetailListStr & "</TD>"
					 
					   GetDetailListStr = GetDetailListStr & "<TD width=""15%"" nowrap>" & Year(AddDate) & "-" & Month(AddDate) & "-" & Day(AddDate) & "</TD>"
					   GetDetailListStr = GetDetailListStr & "<TD width=""15%"" nowrap>点击 <B>" & ObjRS("Hits") & "</B> 次</TD>"
					   GetDetailListStr = GetDetailListStr & "</TR>"
					   GetDetailListStr = GetDetailListStr & "<TR height=40>"
					   GetDetailListStr = GetDetailListStr & "<TD Style = ""BORDER-RIGHT: #efefef 1px dotted; BORDER-LEFT: #efefef 1px dotted; BORDER-BOTTOM: #efefef 1px dotted"" align=middle width=""14%""><table border=0><tr><td>"
					   
					   If ObjRS("LinkType") = 0 Then
						GetDetailListStr = GetDetailListStr & "<A href = ""to?" & LinkID & """ target=""_blank""><IMG height=31 src=""../../Images/Default/nologo.gif"" alt=" & ObjRS("SiteName") & " width=88 border=0></A></td></tr>"
					   Else
						GetDetailListStr = GetDetailListStr & "<A href = ""to?" & LinkID & """ target=""_blank""><IMG height=31 src=""" & ObjRS("Logo") & """ alt=" & ObjRS("SiteName") & " width=88 border=0></A></td></tr>"
					   End If
					   GetDetailListStr = GetDetailListStr & "<tr><td align=""center""><a href=""modify/?LinkID=" & LinkID & """>修改</a> <a href=""del/?id=" & LinkID & """>删除</a></td></tr></table></TD>"
					   GetDetailListStr = GetDetailListStr & "<TD style=""BORDER-RIGHT: #efefef 1px dotted; BORDER-BOTTOM: #efefef 1px dotted"" title=""网站简介"" colSpan=5>"
					   If Trim(ObjRS("Description")) = "" Then
						 GetDetailListStr = GetDetailListStr & "暂无简介"
					   Else
						 GetDetailListStr = GetDetailListStr & KS.HtmlCode(ObjRS("Description"))
					   End If
					   GetDetailListStr = GetDetailListStr & "</TD></TR><TR><TD colSpan=6 height=3></TD></TR>"
					   GetDetailListStr = GetDetailListStr & "</TABLE>"
					 ObjRS.MoveNext
					  I = I + 1
					  If I >= MaxPerPage Then Exit Do
					 Loop
					 GetDetailListStr = GetDetailListStr & "<table width=""100%"" aling=""center""><tr><td>" & KS.ShowPagePara(totalPut, MaxPerPage, "Index.asp", True, "个站点", CurrentPage, KS.QueryParam("page")) & "</td></tr></table>"
		End Function
		'结合上面ReplaceListContent函数使用
		Function GetClassSiteList(FolderID)
				Dim ObjRS:Set ObjRS=Server.CreateObject("ADODB.Recordset")
				Dim SiteName,I
				
				FolderID = KS.ChkClng(FolderID)
				GetClassSiteList = "<tr><td>"
				ObjRS.Open "Select LinkID,sitename From KS_Link Where FolderID=" & FolderID & " And Verific=1 And Locked=0", Conn, 1, 1
					If ObjRS.EOF And ObjRS.BOF Then
						GetClassSiteList = GetClassSiteList & "该类别下没有任何站点!"
					Else
						 GetClassSiteList = GetClassSiteList & "<table width=""100%"" border=""0"">"
						Do While Not ObjRS.EOF
							GetClassSiteList = GetClassSiteList & "<tr>"
							For I = 1 To 6
								SiteName = ObjRS(1)
								GetClassSiteList = GetClassSiteList & "<td><a href = ""to?" & ObjRS(0) & """ target='blank' title='" & SiteName & "'>" & SiteName & "</a></td>"
								ObjRS.MoveNext
								If ObjRS.EOF Then Exit For
							Next
							GetClassSiteList = GetClassSiteList & "</tr>"
						 Loop
						 GetClassSiteList = GetClassSiteList & "</table>"
				 End If
				 GetClassSiteList = GetClassSiteList & "</td></tr>"
				 ObjRS.Close:Set ObjRS = Nothing
		End Function
End Class
%>

 

<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New CommandStatus
KSCls.Kesion()
Set KSCls = Nothing

Class CommandStatus
        Private KS,ChannelID,ItemName,Url
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
			Dim ButtonSymbol, Opstr, ShowOpStr
			Dim FolderRS, FolderID, ParentID, TSArr, I
			Dim LabelFolderID,DisplayMode
			
			FolderID = Request.QueryString("FolderID")
			LabelFolderID = Request.QueryString("LabelFolderID")
			ButtonSymbol = Trim(Request("ButtonSymbol"))
			Opstr = Request.QueryString("OpStr")
			DisplayMode=KS.G("DisplayMode")
			ChannelID = KS.ChkClng(KS.G("ChannelID"))
			If ChannelID=0 Then ChannelID=1
			ShowOpStr = Opstr
			If FolderID = "" Then FolderID = "0"
			If LabelFolderID = "" Then LabelFolderID = "0"
			If FolderID <> "0" Then
					 ShowOpStr = GetFolderName(FolderID)
					If Opstr <> "" Then
					  ShowOpStr = ShowOpStr & " >> <Font Color=Red>" & Opstr & "</Font>"
					End If
				
			End If
			If LabelFolderID <> "0" Then
			   Set FolderRS = Conn.Execute("Select TS,FolderType From KS_LabelFolder Where ID='" & LabelFolderID & "'")
				If Not FolderRS.EOF Then
				   If FolderRS(1) = 1 Then
					  ShowOpStr = "标签管理 >> 自由标签"
				   ElseIf FolderRS(1) = 0 Then
					  ShowOpStr = "标签管理 >> 系统函数标签"
				   ElseIf FolderRS(1) = 2 Then
					  ShowOpStr = "JS 管理 >> 系统 JS"
				 ElseIf FolderRS(1) = 3 Then
					  ShowOpStr = "JS 管理 >> 自由 JS"
				   End If
				   TSArr = Split(FolderRS(0), ",")
				  For I = LBound(TSArr) To UBound(TSArr) - 1
						ShowOpStr = ShowOpStr & " >> " & GetLabelFolderName(TSArr(I))
				  Next
				  ShowOpStr = Right(ShowOpStr, 30)
				End If
				FolderRS.Close
				Set FolderRS = Nothing
				
			End If
			
			ItemName=KS.C_S(ChannelID,3)
			Select Case KS.C_S(ChannelID,6)
			 Case 1:Url="KS.iteminfo.asp"
			 Case 2:Url="KS.iteminfo.asp"
			 Case 3:Url="KS.iteminfo.asp"
			 Case 4:Url="KS.iteminfo.asp"
			 Case 5:Url="KS.iteminfo.asp"
			 Case 6:ItemName="歌曲"
			 Case 7:Url="KS.iteminfo.asp"
			 Case 8:Url="KS.iteminfo.asp"
			 Case else
			  ItemName="文章"
			 End Select
			 If KS.G("Go")="Class" Then Url="KS.Class.asp"
			With KS
			    .echo"<html>"
				.echo"<head>"
				.echo"<meta http-equiv=""Content-Language"" content=""zh-cn"">"
				.echo"<meta HTTP-EQUIV=""Content-Type"" CONTENT=""text/html; charset=utf-8"">"
				.echo"<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
				.echo"<script language=""JavaScript"" src=""Include/SetFocus.js""></script>"
				.echo"<script language=""JavaScript"" src=""../KS_Inc/jquery.js""></script>"
				.echo"<body>"
				.echo"  <ul id='split'>"
				.echo"      <div id=""daohang""><strong>&nbsp;操作导航：</strong>" & ShowOpStr & "</div>"
				.echo"      <div id='splitright'>"
                .echo"       <input type='button' value='添加信息' id=""Button1"" onclick=""ClickButton1();"">"
				.echo"       <input type='button' value='编辑信息' id=""Button2"" onclick=""ClickButton2();"">"
				.echo"       <input type='button' value='删除信息' id=""Button3"" onclick=""ClickButton3();"">"
				.echo"      </div>"
				.echo"   </ul>"
				.echo"</body>"
				.echo"</html>"
				.echo("<SCRIPT language=javascript>")
				.echo("   function ClickButton1(){ ")
				   Select Case (UCase(ButtonSymbol))
				            Case "GO","GOSAVE","SETPARAM","EDITORADDSAVE","EDITOREDIT","VOTEADDSAVE","VOTEEDIT","FILTERSADD", "FILTERSEDIT","LABELADD","SETPOWER","DIYFUNCTIONSTEP1"
							     .echo("$(parent.document).find('#MainFrame')[0].contentWindow.CheckForm();" & vbcrlf)
							Case "VIEWFOLDER", "ARTICLESEARCH"
								 .echo("$(parent.document).find('#MainFrame')[0].contentWindow.CreateNews();" & vbcrlf)
							Case "ADDINFO"
								 .echo("$(parent.document).find('#MainFrame')[0].contentWindow.SubmitFun();" & vbcrlf)
							Case "FREELABEL", "FUNCTIONLABEL","DIYFUNCTIONLABEL"
								 .echo("$(parent.document).find('#MainFrame')[0].contentWindow.AddLabel('Include/');" & vbcrlf)

							Case "LABELFOLDERADD"
								 .echo("$(parent.document).find('#MainFrame')[0].contentWindow.frames['CreateLabelFolderFrame'].CheckForm();")
							Case "SYSJSLIST", "FREEJSLIST"
								 .echo("$(parent.document).find('#MainFrame')[0].contentWindow.AddJS('Include/');")
							Case "JSADD", "JSEDIT"
								 .echo("$(parent.document).find('#MainFrame')[0].contentWindow.frames['JSFrame'].CheckForm();")
							Case "SPECIALSEARCH"
								 .echo("$(parent.document).find('#MainFrame')[0].contentWindow.Edit();")
							Case "VOTELIST"
								 .echo("$(parent.document).find('#MainFrame')[0].contentWindow.VoteAdd();")
							Case "INNERLINKLIST"
								  .echo("$(parent.document).find('#MainFrame')[0].contentWindow.InnerAdd();")
							Case "VIEWLINK"
								 .echo("$(parent.document).find('#MainFrame')[0].contentWindow.CreateLink();")
							Case "COLLECTHISTORY"
								 .echo("$(parent.document).find('#MainFrame')[0].contentWindow.DelRecords('Collect/');")
							Case "DATACOLLECT"
								 .echo("$(parent.document).find('#MainFrame')[0].contentWindow.Collect(1);")
					 End Select
				.echo("   }")
				.echo(" function ClickButton2()")
				.echo(" {")
				 Select Case (UCase(ButtonSymbol))
						 Case "VIEWFOLDER", "ARTICLESEARCH"
							.echo("$(parent.document).find('#MainFrame')[0].contentWindow.Edit();")
						 Case "PICTURESEARCH"
							.echo("$(parent.document).find('#MainFrame')[0].contentWindow.Edit();")
						 Case "DOWNSEARCH"
							.echo("$(parent.document).find('#MainFrame')[0].contentWindow.Edit();")
						 Case "FLASHSEARCH"
							.echo("$(parent.document).find('#MainFrame')[0].contentWindow.Edit();")
						 Case "PRODUCTSEARCH"
							.echo("$(parent.document).find('#MainFrame')[0].contentWindow.Edit();")
						 Case "VIEWSHOPFOLDER", "SHOPSEARCH"
							.echo("$(parent.document).find('#MainFrame')[0].contentWindow.Edit();")
						 Case "MOVIESEARCH"
							.echo("$(parent.document).find('#MainFrame')[0].contentWindow.Edit();")
						 Case "ADDINFO"
							.echo("$(parent.document).find('#MainFrame')[0].contentWindow.location.href='" & Url & "?ComeFrom=" & KS.S("ComeFrom") & "&ChannelID=" & ChannelID &"&ID=" & FolderID & "';")
							.echo("location.href='KS.Split.asp?ButtonSymBol=ViewFolder&FolderID=" & FolderID & "&ChannelID=" & ChannelID & "';")
						Case "ADDGQ"
							.echo("$(parent.document).find('#MainFrame')[0].contentWindow.location.href='KS.Supply.asp?ID=" & FolderID & "&DisplayMode=" & DisplayMode &"';")
							.echo("location.href='KS.Split.asp?ButtonSymBol=ViewGQFolder&FolderID=" & FolderID & "';")
						Case "FREELABEL","DIYFUNCTIONLABEL","FUNCTIONLABEL"
							.echo("$(parent.document).find('#MainFrame')[0].contentWindow.Edit('Include/');")
						Case "SYSJSLIST", "FREEJSLIST", "SYSLABELSEARCH", "FREELABELSEARCH", "SYSJSSEARCH", "FREEJSSEARCH"
							.echo("$(parent.document).find('#MainFrame')[0].contentWindow.Edit('Include/');")
						Case "SPECIALSEARCH"
							.echo("$(parent.document).find('#MainFrame')[0].contentWindow.Delete();")
						Case "VOTELIST"
							.echo("$(parent.document).find('#MainFrame')[0].contentWindow.VoteControl(1);")
						Case "INNERLINKLIST"
							.echo("$(parent.document).find('#MainFrame')[0].contentWindow.InnerControl(1);")
						Case "GO","GOSAVE","SETPOWER","EDITOREDIT","VOTEEDIT", "EDITORADDSAVE", "JSADD", "VOTEADDSAVE", "FILTERSADD", "FILTERSEDIT","DIYFUNCTIONSTEP1"
							.echo("$(parent.document).find('#MainFrame')[0].contentWindow.location.href='javascript:history.back()';")
							.echo("history.back(-1);")
						Case "VIEWLINK", "LINKSEARCH"
							.echo("$(parent.document).find('#MainFrame')[0].contentWindow.Edit();")
						Case "COLLECTHISTORY"
							.echo("$(parent.document).find('#MainFrame')[0].contentWindow.DelAllRecords('Collect/');")
				  End Select
				 .echo(" }")
				 .echo("   function ClickButton3()")
				  .echo(" {")
				  Select Case (UCase(ButtonSymbol))
						 Case "VIEWFOLDER", "ARTICLESEARCH"
							.echo("$(parent.document).find('#MainFrame')[0].contentWindow.Delete();")
						Case "PICTURESEARCH"
							.echo("$(parent.document).find('#MainFrame')[0].contentWindow.Delete();")
						Case "DOWNSEARCH"
							.echo("$(parent.document).find('#MainFrame')[0].contentWindow.Delete();")
						Case "FLASHSEARCH"
							.echo("$(parent.document).find('#MainFrame')[0].contentWindow.Delete();")
						Case "PRODUCTSEARCH"
							.echo("$(parent.document).find('#MainFrame')[0].contentWindow.Delete();")
						Case "VIEWSHOPFOLDER", "SHOPSEARCH"
							.echo("$(parent.document).find('#MainFrame')[0].contentWindow.Delete();")
						Case "MOVIESEARCH"
							.echo("$(parent.document).find('#MainFrame')[0].contentWindow.Delete();")
						Case "FREELABEL", "FUNCTIONLABEL","DIYFUNCTIONLABEL"
							.echo("$(parent.document).find('#MainFrame')[0].contentWindow.Delete('Include/');")
						Case "SYSJSLIST", "FREEJSLIST", "SYSLABELSEARCH", "FREELABELSEARCH", "SYSJSSEARCH", "FREEJSSEARCH","DIYFUNCTIONSEARCH","DIYFUNCTIONLABEL"
							.echo("$(parent.document).find('#MainFrame')[0].contentWindow.Delete('Include/');")
						Case "VOTELIST"
							.echo("$(parent.document).find('#MainFrame')[0].contentWindow.VoteControl(2);")	
						Case "INNERLINKLIST"
							.echo("$(parent.document).find('#MainFrame')[0].contentWindow.InnerControl(2);")
						Case "VIEWLINK", "LINKSEARCH"
							.echo("$(parent.document).find('#MainFrame')[0].contentWindow.Delete();")
				  End Select
				  .echo(" }")
				  .echo(" function ClickButton4()")
				  .echo(" {")

				  .echo(" }")
				  .echo(" $(document).ready(function(){")
				  Select Case (UCase(ButtonSymbol))
					Case "DISABLED"
					   .echo(" $('#Button1').attr('disabled',true);")
					   .echo(" $('#Button2').attr('disabled',true);")
					   .echo(" $('#Button3').attr('disabled',true);")
					Case "VIEWFOLDER"
					  .echo(" $('#Button1').val('添加" & ItemName &"');")
					  .echo(" $('#Button2').val('编辑" & ItemName &"');")
					  .echo(" $('#Button3').val('删除" & ItemName &"');")
				   Case "ADDINFO"
					  .echo(" $('#Button1').val('确定保存');")
					  .echo(" $('#Button2').val('取消返回');")
					  .echo(" $('#Button3').attr('disabled',true);")
				   Case "SEARCH"
					  .echo(" $('#Button1').attr('disabled',true);")
					  .echo(" $('#Button2').val('编辑" & ItemName &"');")
					  .echo(" $('#Button3').val('删除" & ItemName &"');")
				  Case "GO","EDITORADDSAVE", "AUTHORADDSAVE", "JSADD", "VOTEADDSAVE","FILTERSADD"
					 .echo(" $('#Button1').val('确定增加');")
					 .echo(" $('#Button2').val('取消返回');")
					 .echo(" $('#Button3').attr('disabled',true);")
					Case "GOSAVE","EDITOREDIT", "KEYWORDEDIT", "VOTEEDIT", "FILTERSEDIT"
					 .echo(" $('#Button1').val('确定修改');")
					 .echo(" $('#Button2').val('取消返回');")
					 .echo(" $('#Button3').attr('disabled',true);")
					 Case "FREELABEL", "FUNCTIONLABEL","DIYFUNCTIONLABEL"
					  .echo("$('#Button1').val('新建标签');")
					  .echo("$('#Button2').val('修改标签');")
					  .echo("$('#Button3').val('删除标签');")
					 Case "DIYFUNCTIONSTEP1"
					  .echo("$('#Button1').val('下 一 步');")
					  .echo("$('#Button2').val('取消返回');")
					  .echo("$('#Button3').attr('disabled',true);")
					 Case "SYSLABELSEARCH", "FREELABELSEARCH","DIYFUNCTIONSEARCH"
					  .echo("$('#Button1').attr('disabled',true);")
					  .echo("$('#Button2').val('修改标签');")
					  .echo("$('#Button3').val('删除标签');")
					 Case "LABELADD"
					  .echo("$('#Button1').val('保存标签');")
					  .echo("$('#Button2').attr('disabled',true);")
					  .echo("$('#Button3').attr('disabled',true);")
					Case "LABELFOLDERADD"
					  .echo("$('#Button1').val('保存目录');")
					  .echo("$('#Button2').attr('disabled',true);")
					  .echo("$('#Button3').attr('disabled',true);")
					Case "SETPARAM"
					  .echo("$('#Button1').val('保存设置');")
					  .echo("$('#Button2').attr('disabled',true);")
					  .echo("$('#Button3').attr('disabled',true);")
					Case "SETPOWER"
					  .echo("$('#Button1').val('保存设置');")
					  .echo("$('#Button2').val('取消返回');")
					  .echo("$('#Button3').attr('disabled',true);")
					Case "SYSJSLIST", "FREEJSLIST"
					  .echo("$('#Button1').val('新建 JS');")
					  .echo("$('#Button2').val('修改 JS');")
					  .echo("$('#Button3').val('删除 JS');")
					Case "SYSJSSEARCH", "FREEJSSEARCH"
					  .echo("$('#Button1').attr('disabled',true);")
					  .echo("$('#Button2').val('修改 JS');")
					  .echo("$('#Button3').val('删除 JS');")
					Case "JSEDIT"
					  .echo(" $('#Button1').val('确定修改');")
					  .echo(" $('#Button2').attr('disabled',true);")
					  .echo(" $('#Button3').attr('disabled',true);")
					Case "MANAGERSEARCH"
					  .echo("$('#Button1').attr('disabled',true);")
					  .echo("$('#Button2').val('改管理员');")
					  .echo("$('#Button3').val('删管理员');")
					Case "VOTELIST"
					  .echo("$('#Button1').val('添加主题');")
					  .echo("$('#Button2').val('编辑主题');")
					  .echo("$('#Button3').val('删除主题');")
				   Case "LINKSEARCH"
					  .echo(" $('#Button1').attr('disabled',true);")
					  .echo(" $('#Button2').val('编辑链接');")
					  .echo(" $('#Button3').val('删除链接');")
				   Case "DATACOLLECT"
					  .echo(" $('#Button1').val('开始采集');")
					  .echo(" $('#Button2').attr('disabled',true);")
					  .echo(" $('#Button3').attr('disabled',true);")
				 End Select
				.echo(" });  ")
				.echo("</SCRIPT>")
				End With
			End Sub
			Function GetLabelFolderName(FolderID)
				  Dim FolderRS:Set FolderRS = Conn.Execute("Select FolderName From KS_LabelFolder Where ID='" & FolderID & "'")
				  If Not FolderRS.EOF Then
					GetLabelFolderName = FolderRS(0)
				  Else
					GetLabelFolderName = ""
				  End If
				  FolderRS.Close:Set FolderRS = Nothing
			End Function
			Function GetFolderName(FolderID)
				  Dim FolderRS, I, TSArr, TempFolderName
				  Set FolderRS = Conn.Execute("Select TS,FolderName,ChannelID From KS_Class Where ID='" & FolderID & "'")
				  If Not FolderRS.EOF Then
					   TSArr = Split(FolderRS(0), ",")
					   ChannelID=FolderRS(2)
					  For I = LBound(TSArr) To UBound(TSArr) - 1
						If I = 0 Then
						 TempFolderName = KS.C_C(TSArr(I),1)
						Else
						 TempFolderName = TempFolderName & " >> " & KS.C_C(TSArr(I),1)
						End If
					  Next
					 GetFolderName = TempFolderName
				  Else
					GetFolderName = ""
				  End If
				  FolderRS.Close:Set FolderRS = Nothing
				End Function
End Class
%> 

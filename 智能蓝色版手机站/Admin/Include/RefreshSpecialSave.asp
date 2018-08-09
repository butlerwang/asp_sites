<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="Session.asp"-->
<%

Dim KSCls
Set KSCls = New RefreshSpecialSave
KSCls.Kesion()
Set KSCls = Nothing

Class RefreshSpecialSave
        Private KS,KSRObj
		Private RefreshFlag
		Private ReturnInfo
		Private StartRefreshTime
		Private ChannelID
		Private Types
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSRObj=New Refresh
		End Sub
		Function Kesion()
		Types = Request("Types")             'Index 生成专题首页操作 Special 生成专题页操作
		RefreshFlag = Request("RefreshFlag") '取得是按何种类型刷新,如Folder发布指定的专题 All发布所有专题
		'刷新时间
		StartRefreshTime = Request("StartRefreshTime")
		If StartRefreshTime = "" Then StartRefreshTime = Timer()
		  Select Case Types
			 Case "Special"          '刷新专题页
				 Call RefreshSpecial
			 Case "Index"            '刷新专题首页
				 Call RefreshSpecialIndex
			 Case "ChannelSpecial"   '刷新频道专题列表页
				 Call RefreshChannelSpecial
		End Select
		End Function
		Sub Main()
		 Response.Write ("<html>")
		 Response.Write ("<head>")
		 Response.Write ("<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">")
		 Response.Write ("<title>系统信息</title>")
		 Response.Write ("</head>")
		 Response.Write ("<link rel=""stylesheet"" href=""Admin_Style.css"">")
		 Response.Write ("<body oncontextmenu=""return false;"">")
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

		 Response.Write ("<table width=""80%"" height=""50%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">")
		 Response.Write (" <tr>")
		 Response.Write ("   <td height=""50"">")
		 Response.Write ("     <div align=""center""> ")
		 Response.Write (ReturnInfo)
		 Response.Write ("       </div></td>")
		 Response.Write ("   </tr>")
		 Response.Write ("</table>")
		 Response.Write ("</body>")
		 Response.Write ("</html>")
		End Sub
		
		'=============================================================================================
		'以下为本模块相应处理的函数
		'===============================================================================================
		
		'生成专题首页的处理过程
		Sub RefreshSpecialIndex()
		   Dim InstallDir, IndexFile, SaveFilePath
		   Dim SpecialDir, FileContent, Domain
		   FCls.RefreshType = "SpecialIndex"  '设置刷新类型，以便取得当前位置导航等
		   FCls.RefreshFolderID = "0"         '设置当前刷新目录ID 为"0" 以取得通用标签
		   FCls.CurrSpecialID="" '清除当前专题ID
		   
		   InstallDir = KS.Setting(3)
		   SpecialDir = KS.Setting(95)
		   If Left(SpecialDir, 1) = "/" Or Left(SpecialDir, 1) = "\" Then SpecialDir = Right(SpecialDir, Len(SpecialDir) - 1)
		   IndexFile = KS.Setting(5)
			SaveFilePath = InstallDir & SpecialDir
			FileContent = KSRObj.LoadTemplate(KS.Setting(111))
			If FileContent = "" Then
			  ReturnInfo = "数据库中找不到专题首页模板"
			  Call Main
			  Response.End
			Else
			  On Error Resume Next
			  FileContent = KSRObj.ReplaceLableFlag(KSRObj.ReplaceAllLabel(FileContent)) '替换函数标签
			  FileContent = KSRObj.ReplaceGeneralLabelContent(FileContent)  '替换通用标签 如{$GetWebmaster}
			  If Err Then
			   ReturnInfo = Err.Description
				 Err.Clear
				Call Main
				Response.End
			  End If
			  Call KS.CreateListFolder(SaveFilePath)
			  Call KSRObj.FSOSaveFile(FileContent, SaveFilePath & IndexFile)
			  If Err Then
				ReturnInfo = Err.Description
				 Err.Clear
				Call Main
				Response.End
			  End If
			  Domain = KS.GetDomain

			  ReturnInfo = "专题首页发布成功！总费时:<font color=red>" & Left((Timer() - StartRefreshTime), 4) & "</font>秒<br><br>"
			  ReturnInfo = ReturnInfo & "点击浏览: <a href=" & Domain & SpecialDir & IndexFile & " target=_blank>浏览专题首页</a><br><br>"
			  if request("from")="quick" then '一键发布
			     Response.Write ("<link rel=""stylesheet"" href=""Admin_Style.css"">")
			  	 response.Write "<br/><br/><br/><br/><br/><div style='text-align:center'>专题首页生成完成，<a href='" & Domain & SpecialDir & IndexFile & "' target='_blank'>点此浏览</a>!两秒后执行下一个发布任务...</div>"
				 ks.die "<script>setTimeout(""parent.myframeleft.location.href='refreshquick.asp?action=startleft';"",2000);</script>"

			  end if
			  
			  ReturnInfo = ReturnInfo & "<input name=""button1"" type=""button"" onclick=""javascript:location='RefreshSpecial.asp';"" class=""button"" value="" 返 回 "">"
			  Call Main
			    Response.Write "<script>" & vbCrlf
			  	Response.Write "img2.width=400;" & vbCrLf
				Response.Write "txt2.innerHTML=""生成专题首页结束！100"";" & vbCrLf
				Response.Write "document.getElementById('txt3').parentElement.style.display='none';" & vbCrLf
				Response.Write "</script>" & vbCrLf
			End If
		End Sub
		'生成专题分类的处理过程
		Sub RefreshChannelSpecial()
		 Dim FolderID, RefreshSql, RefreshTotalNum, RefreshRS, NewsTotalNum, NewsNo
		  RefreshSql = Trim(Request("RefreshSql"))
		  NewsNo = Request("NewsNo")
		 If NewsNo = "" Then NewsNo = 0
		 If RefreshSql = "" Then
		  Select Case RefreshFlag
			Case "Folder"
				FolderID = Replace(Request("FolderID")," ","")
				If FolderID <> "" Then
				  RefreshSql = "Select * from [KS_SpecialClass] where ClassID IN (" & FolderID & ") Order By ClassID"
				Else
				  RefreshSql = "Select * From [KS_SpecialClass] Where 1=0"
				End If
		   Case "All"
				RefreshSql = "Select * from [KS_SpecialClass] Order By ClassID"
		   Case Else
			RefreshSql = ""
			RefreshTotalNum = 0
		  End Select
		End If
		If RefreshSql <> "" Then
		    Call Main
			Set RefreshRS = Server.CreateObject("ADODB.RecordSet")
			RefreshRS.Open RefreshSql, Conn, 1, 1
			NewsTotalNum = RefreshRS.RecordCount
			If RefreshRS.EOF Then
				ReturnInfo = "没有要刷新的专题分类&nbsp;&nbsp;<br><input name=""button1"" type=""button"" onclick=""javascript:location='RefreshSpecial.asp';"" class=""button"" value="" 返 回 "">"
				Set RefreshRS = Nothing
			Else
				For NewsNo=1 To NewsTotalNum
				   Call KSRObj.RefreshSpecialClass(RefreshRS)  '调用频道专题刷新函数
				   Call InnerJs(NewsNo,NewsTotalNum,"个专题分类")
				   RefreshRS.MoveNext
				Next
			End If
				Response.Write "<script>"
				Response.Write "img2.width=400;" & vbCrLf
				Response.Write "txt2.innerHTML=""生成专题分类结束！100"";" & vbCrLf
				Response.Write "txt3.innerHTML=""总共生成了 <font color=red><b>" & NewsTotalNum & "</b></font> 个专题分类,总费时:<font color=red>" & Left((Timer() - StartRefreshTime), 4) & "</font> 秒"
				if request("from")="quick" then '一键发布
				 response.write "<br/>两秒后执行下一个发布任务..."
				Else
				Response.Write "<br><br><input name='button1' type='button' onclick=javascript:location='RefreshSpecial.asp'; class='button' value=' 返 回 '>"
				End IF
				Response.Write """;</script>" & vbCrLf
				
				if request("from")="quick" then '一键发布
				 ks.die "<script>setTimeout(""parent.myframeleft.location.href='refreshquick.asp?action=startleft';"",2000);</script>"
			     end if

			Set RefreshRS = Nothing
		End If
		End Sub
		'生成专题页的处理过程
		Sub RefreshSpecial()
		  Dim FolderID, RefreshSql, RefreshTotalNum, RefreshRS, NewsTotalNum, NewsNo
		  RefreshSql = Trim(Request("RefreshSql"))
		  NewsNo = Request("NewsNo")
		 If NewsNo = "" Then NewsNo = 0
		 If RefreshSql = "" Then
		  Select Case RefreshFlag
		  	Case "ID"
				RefreshSql = "Select * From KS_Special where specialid in(" & KS.G("ID") & ") Order By SpecialAddDate Desc"
			Case "New"
				Dim TotalNum
				TotalNum = Request.Form("TotalNum")
				If TotalNum = "" Then TotalNum = 20
				RefreshSql = "Select Top " & TotalNum & " * From KS_Special Order By SpecialAddDate Desc"
			Case "Folder"
				FolderID = KS.FilterIDs(Request("FolderID"))
				If FolderID <> "" Then
				RefreshSql = "Select * from [KS_Special] where  ClassID IN (" & FolderID & ") Order By SpecialAddDate Desc"
				Else
				RefreshSql = "Select * From [KS_Special] Where 1=0"
				End If
		   Case "All"
				'RefreshSql = "Select * from [KS_Special] a inner join ks_channel b on a.channelid=b.channelid where b.FsoHtmlTF=1 order by specialadddate desc"
				RefreshSql = "Select * from [KS_Special] where 1=0 order by specialadddate desc"
		   Case Else
			RefreshSql = ""
			RefreshTotalNum = 0
		  End Select
		End If
		If RefreshSql <> "" Then
			Call Main
			Set RefreshRS = Server.CreateObject("ADODB.RecordSet")
			RefreshRS.Open RefreshSql, Conn, 1, 1
			NewsTotalNum = RefreshRS.RecordCount
			If RefreshRS.EOF Then
				Response.Write "<script>img2.width=""0"";" & vbCrLf
				Response.Write "txt2.innerHTML=""对不起，没有可生成的专题！"
				if request("from")="quick" then '一键发布
				 response.write "两秒后执行下一个发布任务..."
				Else
				Response.Write "<br><br><input name='button1' type='button' onclick=javascript:location='RefreshSpecial.asp'; class='button' value=' 返 回 '>"
				End IF
				Response.Write """;" & vbCrLf
				Response.Write "txt3.innerHTML="""";" & vbCrLf
				Response.Write "txt4.innerHTML="""";" & vbCrLf
				Response.Write "document.all.BarShowArea.style.display='none';" & vbCrLf
				Response.Write "</script>" & vbCrLf
				if request("from")="quick" then '一键发布
				 ks.die "<script>setTimeout(""parent.myframeleft.location.href='refreshquick.asp?action=startleft';"",2000);</script>"
			     end if
				
				Response.Flush
				Set RefreshRS = Nothing
			Else
			   For NewsNo=1 To NewsTotalNum
				   Call KSRObj.RefreshSpecials(RefreshRS)  '调用专题刷新函数
                   Call InnerJS(NewsNo,NewsTotalNum,"个专题")
				   RefreshRS.MoveNext
			  Next 
				Response.Write "<script>"
				Response.Write "img2.width=400;" & vbCrLf
				Response.Write "txt2.innerHTML=""生成专题结束！100"";" & vbCrLf
				Response.Write "txt3.innerHTML=""总共生成了 <font color=red><b>" & NewsTotalNum & "</b></font> 条,总费时:<font color=red>" & Left((Timer() - StartRefreshTime), 4) & "</font> 秒"
				if request("from")="quick" then '一键发布
				 response.write "<br/>两秒后执行下一个发布任务..."
				Else
				Response.Write "<br><br><input name='button1' type='button' onclick=javascript:location='RefreshSpecial.asp'; class='button' value=' 返 回 '>"
				End If
				Response.Write """;</script>" & vbCrLf
				
				if request("from")="quick" then '一键发布
				 ks.die "<script>setTimeout(""parent.myframeleft.location.href='refreshquick.asp?action=startleft';"",2000);</script>"
			     end if
				
			End If
			Set RefreshRS = Nothing
		End If
		End Sub
        Sub InnerJS(NowNum,TotalNum,itemname)
		  With Response
				.Write "<script>"
				'.Write "fsohtml.innerHTML='" & FsoHtmlList & "';" & vbCrLf
				.Write "img2.width=" & Fix((NowNum / TotalNum) * 400) & ";" & vbCrLf
				.Write "txt2.innerHTML=""生成进度:" & FormatNumber(NowNum / TotalNum * 100, 2, -1) & """;" & vbCrLf
				.Write "txt3.innerHTML=""总共需要生成 <font color=red><b>" & TotalNum & "</b></font> " & itemname & ",<font color=red><b>在此过程中请勿刷新此页面！！！</b></font> 系统正在生成第 <font color=red><b>" & NowNum & "</b></font> " & itemname & """;" & vbCrLf
				.Write "img2.title=""(" & NowNum & ")"";</script>" & vbCrLf
				.Flush
		  End With
		End Sub
End Class
%> 

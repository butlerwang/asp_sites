<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="include/session.asp"-->
<%
Dim KSCls
Set KSCls= New AdminUploadFileCls
KSCls.Kesion()
Set KSCls=Nothing

Class AdminUploadFileCls
	Private KS
	Private ChannelDir, fullPath, FilePath, UploadDir, ThisDir
	Private Action, rsChannel
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		End Sub

	Public Sub Kesion()
		Action = LCase(Request("action"))
       If Not KS.ReturnPowerResult(0, "KMST10018") Then                  '权限检查
				Call KS.ReturnErr(1, "")   
				Response.End()
		End iF
				
		ChannelDir = KS.Setting(3)& KS.Setting(91)
		If Trim(Request("UploadDir")) <> "" Then
			UploadDir = Trim(Request("UploadDir")) & "/"
		End If
		If Trim(Request("ThisDir")) <> "" Then
			ThisDir = Trim(Request("ThisDir")) & "/"
		End If
		ThisDir = Replace(ThisDir, "\", "/")
		if instr(ThisDir,".")<>0 or instr(UploadDir,".")<>0 then
		  ks.die "非法路径!"
		end if

		if (left(UploadDir,1)="/") Then UploadDir=Right(UploadDir,len(UploadDir)-1)
		FilePath = Replace(ChannelDir & UploadDir, "\", "/")

		fullPath = Server.MapPath(FilePath)

		Select Case Trim(Action)
		Case "clear"
			Call ClearUploadFile
		Case "delete"
			Call DelUselessFile
		Case "del"
			Call DelFile
		Case "delalldirfile"
			Call DelAllDirFile
		Case "delthisallfile"
			Call DelThisAllFile
		Case "delemptyfolder"
			Call DelEmptyFolder
		Case Else
			Call ShowUploadMain
		End Select
	End Sub

	
	'=================================================
	'过程名：ShowChildFolder
	'作  用：显示子目录菜单
	'=================================================
	Private Sub ShowChildFolder()
		Dim fso, fsoFile, DirFolder
		Dim strFolderPath
		On Error Resume Next
		strFolderPath = ChannelDir & UploadDir
		strFolderPath = Server.MapPath(strFolderPath)
		Set fso = KS.InitialObject(KS.Setting(99))
		If fso.FolderExists(strFolderPath) Then
			Set fsoFile = fso.GetFolder(strFolderPath)
			For Each DirFolder In fsoFile.SubFolders
				Response.Write "<a href=""?UploadDir=" &Request("UploadDir") & "/"& DirFolder.Name& "&ThisDir=" & DirFolder.Name & """><img src=""images/folder/folderclosed.gif"" width=20 height=20 border=0 alt=""修改时间：" & DirFolder.DateLastModified & """ align=absMiddle> "
				If Replace(ThisDir, "/", "") = DirFolder.Name Then
					Response.Write "<font color=red>" & DirFolder.Name & "</font>"
				Else
					Response.Write DirFolder.Name
				End If
				Response.Write "</a> &nbsp;&nbsp;" & vbNewLine
			Next
		Else
			Response.Write "没有找到文件夹！"
		End If
		Set fsoFile = Nothing: Set fso = Nothing
	End Sub

	'=================================================
	'函数名：showpage
	'作  用：分页
	'=================================================
	Private Function showpage(ByVal CurrentPage, ByVal TotalNumber, ByVal maxperpage, ByVal TotleSize)
		Dim n
		Dim strTemp
		
		If (TotalNumber Mod maxperpage) = 0 Then
			n = TotalNumber \ maxperpage
		Else
			n = TotalNumber \ maxperpage + 1
		End If
		strTemp = "<table align='center'><form method='Post' action='?UploadDir=" & Request("UploadDir") & "&ThisDir=" & Request("ThisDir") & "'><tr><td>" & vbNewLine
		strTemp = strTemp & "共 <b>" & TotalNumber & "</b> 个文件，占用 <b>" & TotleSize & "</b>&nbsp;&nbsp;"
		'sfilename = JoinChar(sfilename)
		If CurrentPage < 2 Then
			strTemp = strTemp & "首页 上一页&nbsp;"
		Else
			strTemp = strTemp & "<a href='?page=1&UploadDir=" & Request("UploadDir") & "&ThisDir=" & Request("ThisDir") & "'>首页</a>&nbsp;"
			strTemp = strTemp & "<a href='?page=" & (CurrentPage - 1) & "&UploadDir=" & Request("UploadDir") & "&ThisDir=" & Request("ThisDir") & "'>上一页</a>&nbsp;"
		End If

		If n - CurrentPage < 1 Then
			strTemp = strTemp & "下一页 尾页"
		Else
			strTemp = strTemp & "<a href='?page=" & (CurrentPage + 1) & "&UploadDir=" & Request("UploadDir") & "&ThisDir=" & Request("ThisDir") & "'>下一页</a>&nbsp;"
			strTemp = strTemp & "<a href='?page=" & n & "&UploadDir=" & Request("UploadDir") & "&ThisDir=" & Request("ThisDir") & "'>尾页</a>"
		End If
		strTemp = strTemp & "&nbsp;页次：<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>页 "
		strTemp = strTemp & "&nbsp;转到："
		strTemp = strTemp & "<input name=page size=3 value='" & CurrentPage & "'> <input type=submit name=Submit value='转到' class=Button>"
		strTemp = strTemp & "</select>"
		strTemp = strTemp & "</td>"
		strTemp = strTemp & "<td></td>"
		strTemp = strTemp & "</tr></form></table>"
		showpage = strTemp
	End Function
	'=================================================
	'函数名：GetFilePic
	'作  用：获取文件图片
	'=================================================
	Private Function GetFilePic(sName)
		Dim FileName, Icon
		FileName = LCase(GetExtensionName(sName))
		Select Case FileName
			Case "gif", "jpg", "bmp", "png"
				Icon = sName
			Case "exe"
				Icon = "../editor/ksplus/FileIcon/file_exe.gif"
			Case "rar"
				Icon = "../editor/ksplus/FileIcon/file_rar.gif"
			Case "zip"
				Icon = "../editor/ksplus/FileIcon/file_zip.gif"
			Case "swf"
				Icon = "../editor/ksplus/FileIcon/file_flash.gif"
			Case "rm", "wma"
				Icon = "../editor/ksplus/FileIcon/file_rm.gif"
			Case "mid"
				Icon = "../editor/ksplus/FileIcon/file_media.gif"
			Case Else
				Icon = "../editor/ksplus/FileIcon/file_other.gif"
		End Select
		GetFilePic = Icon
	End Function
	'=================================================
	'函数名：GetExtensionName
	'作  用：获取文件扩展名
	'=================================================
	Private Function GetExtensionName(ByVal sName)
		Dim FileName
		FileName = Split(sName, ".")
		GetExtensionName = FileName(UBound(FileName))
	End Function
	'=================================================
	'函数名：GetFileSize
	'作  用：格式化文件的大小
	'=================================================
	Private Function GetFileSize(ByVal n)
		Dim FileSize
		FileSize = n / 1024
		FileSize = FormatNumber(FileSize, 2)
		If FileSize < 1024 And FileSize > 1 Then
			GetFileSize = "<font color=red>" & FileSize & "</font>&nbsp;KB"
		ElseIf FileSize > 1024 Then
			GetFileSize = "<font color=red>" & FormatNumber(FileSize / 1024, 2) & "</font>&nbsp;MB"
		Else
			GetFileSize = "<font color=red>" & n & "</font>&nbsp;Bytes"
		End If
	End Function
	'=================================================
	'过程名：DelFile
	'作  用：删除文件
	'=================================================
	Private Sub DelFile()
		Dim fso, i
		Dim strFileName, strFilePath
		Dim strFolderName, strFolderPath
		'---- 删除文件
		If Trim(Request("FileName")) <> "" Then
			strFileName = Split(Request("FileName"), ",")
			If UBound(strFileName) <> -1 Then '删除文件
				Set fso = KS.InitialObject(KS.Setting(99))
				For i = 0 To UBound(strFileName)
					strFilePath = Server.MapPath(FilePath & Trim(strFileName(i)))
					If fso.FileExists(strFilePath) Then
						fso.DeleteFile strFilePath, True
					End If
				Next
				Set fso = Nothing
			End If
		End If
		'---- 删除文件夹
		If Trim(Request("FolderName")) <> "" Then
			strFolderName = Split(Request("FolderName"), ",")
			If UBound(strFolderName) <> -1 Then '删除文件
				Set fso = KS.InitialObject(KS.Setting(99))
				For i = 0 To UBound(strFolderName)
					strFolderPath = Server.MapPath(FilePath & Trim(strFolderName(i)))
					If fso.FolderExists(strFolderPath) Then
						fso.DeleteFolder strFolderPath, True
					End If
				Next
				Set fso = Nothing
			End If
		End If
		Response.Redirect (Request.ServerVariables("HTTP_REFERER"))
	End Sub
	'=================================================
	'过程名：DelAllDirFile
	'作  用：删除所有文件和文件夹
	'=================================================
	Private Sub DelAllDirFile()
		Dim fso, oFolder
		Dim DirFile, DirFolder
		Dim tempPath
		
		Set fso = KS.InitialObject(KS.Setting(99))
		If fso.FolderExists(fullPath) Then
			Set oFolder = fso.GetFolder(fullPath)
			'---- 删除所有文件
			For Each DirFile In oFolder.Files
				tempPath = fullPath & "\" & DirFile.Name
				fso.DeleteFile tempPath, True
			Next
			'---- 删除所有子目录
			For Each DirFolder In oFolder.SubFolders
				tempPath = fullPath & "\" & DirFolder.Name
				fso.DeleteFolder tempPath, True
			Next
			Set oFolder = Nothing
		End If
		Set fso = Nothing
		Response.Redirect (Request.ServerVariables("HTTP_REFERER"))
	End Sub
	'=================================================
	'过程名：DelThisAllFile
	'作  用：删除当前目录所有文件
	'=================================================
	Private Sub DelThisAllFile()
		Dim fso, oFolder
		Dim DirFiles
		Dim tempPath
		
		Set fso = KS.InitialObject(KS.Setting(99))
		If fso.FolderExists(fullPath) Then
			Set oFolder = fso.GetFolder(fullPath)
			'---- 删除所有文件
			For Each DirFiles In oFolder.Files
				tempPath = fullPath & "\" & DirFiles.Name
				fso.DeleteFile tempPath, True
			Next
			Set oFolder = Nothing
		End If
		Set fso = Nothing
		Response.Redirect (Request.ServerVariables("HTTP_REFERER"))
	End Sub
	'=================================================
	'过程名：DelEmptyFolder
	'作  用：删除所有空文件夹
	'=================================================
	Private Sub DelEmptyFolder()
		Dim fso, oFolder
		Dim DirFolder, tempPath
		
		Set fso = KS.InitialObject(KS.Setting(99))
		If fso.FolderExists(fullPath) Then
			Set oFolder = fso.GetFolder(fullPath)
			'---- 删除所有空子目录
			For Each DirFolder In oFolder.SubFolders
				If DirFolder.Size = 0 Then
					tempPath = fullPath & "\" & DirFolder.Name
					fso.DeleteFolder tempPath, True
				End If
			Next
			Set oFolder = Nothing
		End If
		Set fso = Nothing
		Response.Redirect (Request.ServerVariables("HTTP_REFERER"))
	End Sub
	'=================================================
	'过程名：ShowUploadMain
	'作  用：显示上传文件主页面
	'=================================================
	Private Sub ShowUploadMain()
		Dim maxperpage, CurrentPage, TotalNumber, Pcount
		Dim fso, FileCount, TotleSize, totalPut
		
		maxperpage = 20 '###每页显示数
		
		If IsNumeric(Request("page")) And Trim(Request("page")) <> "" Then
			CurrentPage = CLng(Request("page"))
		Else
			CurrentPage = 1
		End If
		If CLng(CurrentPage) = 0 Then CurrentPage = 1
		On Error Resume Next
		If Not KS.IsObjInstalled(KS.Setting(99)) Then
			Response.Write "<b><font color=red>你的服务器不支持 fso(Scripting.FileSystemObject)! 不能使用本功能</font></b>"
		End If
		Response.Write "<html>"
		Response.Write "<head>"
		Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>"
		Response.Write "<title>Digg记录管理</title>"
		Response.Write "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
		Response.Write "<script>" &vbnewline
		Response.Write "function CheckAll(form) {  "&vbnewline
		Response.Write "	for (var i=0;i<form.elements.length;i++)  {  "&vbnewline
		Response.Write "		var e = form.elements[i];  "&vbnewline
		Response.Write "		if (e.name != 'chkall')  "&vbnewline
		Response.Write "		e.checked = true // form.chkall.checked;  "&vbnewline
		Response.Write "	}  "&vbnewline
		Response.Write "} "&vbnewline
		 
		Response.Write "function ContraSel(form) {"&vbnewline
		Response.Write "	for (var i=0;i<form.elements.length;i++){"&vbnewline
		Response.Write "		var e = form.elements[i];"&vbnewline
		Response.Write "		if (e.name != 'chkall')"&vbnewline
		Response.Write "		e.checked=!e.checked;"&vbnewline
		Response.Write "	}"&vbnewline
		Response.Write "}"&vbnewline
		Response.Write "</script>"&vbnewline
		Response.Write "</head>"
		
		Response.Write "<body topmargin='0' leftmargin='0'>"
		Response.Write "<ul id='mt'> <div style='font-weight:bold;margin-top:10px'><a href='?'>上传文件管理</a> | <a href='?action=clear'>清理无用文件</a></div></ul>"
		Response.Write "<table border=0 align=center cellpadding=3 style='margin-top:5px' cellspacing=1 width='99%' class='ctable'>"
		Response.Write "<tr>"
		Response.Write "        <td class=clefttitle colspan=""2"" style=""text-align:left"">"
		Call ShowChildFolder
		Response.Write "</td>"
		Response.Write "</tr>"
		Response.Write "<tr>"
		Response.Write "        <td width=""50%"" class=clefttitle style=""text-align:left"">当前目录：" & FilePath & "</td>"
		Response.Write "        <td width=""50%"" align=center class=clefttitle>"
		Response.Write "<!--<a href=""?action=clear&UploadDir=" & Request("UploadDir") & """>清理无用文件</a>--> &nbsp;&nbsp;"
		If Trim(Request("ThisDir")) <> "" Then

			Response.Write "<a href=""?UploadDir=" & Left(Request("UploadDir"),Len(Request("UploadDir"))-Len(Mid(Request("UploadDir"), InStrRev(Request("UploadDir"), "/")))) & "&ThisDir=" & Request("ThisDir") & """>↑返回上一层目录</a>"
		End If
		Response.Write "</td>"
		Response.Write "</tr>"
		Response.Write "</table><br>" & vbNewLine

		Set fso = KS.InitialObject(KS.Setting(99))
		If fso.FolderExists(fullPath) Then
			Dim fsoFile, fsoFileSize
			Dim DirFiles, DirFolder
			Set fsoFile = fso.GetFolder(fullPath)
			Dim c
			FileCount = fsoFile.Files.Count
			TotleSize = GetFileSize(fsoFile.Size)
			totalPut = fsoFile.Files.Count
			If CurrentPage < 1 Then
				CurrentPage = 1
			End If
			If (CurrentPage - 1) * maxperpage > totalPut Then
				If (totalPut Mod maxperpage) = 0 Then
					CurrentPage = totalPut \ maxperpage
				Else
					CurrentPage = totalPut \ maxperpage + 1
				End If
			End If
			FileCount = 0
			c = 0
			Response.Write "<table border=0 align=center cellpadding=3 cellspacing=1 class=ctable width='99%'>" & vbNewLine
			Response.Write "<tr><td colspan=4 class=clefttitle>" & vbNewLine
			Response.Write showpage(CurrentPage, totalPut, maxperpage, TotleSize)
			Response.Write "</td></tr>" & vbNewLine
			Response.Write "<form name=""myform"" method=""post"" action='KS.AdminFiles.asp'>" & vbCrLf
			Response.Write "<tr>" & vbNewLine
			Response.Write "<input type=hidden name=action value='del'>" & vbNewLine
			Response.Write "<input type=hidden name=UploadDir value='" & Request("UploadDir") & "'>" & vbNewLine
			Response.Write "<input type=hidden name=ThisDir value='" & Request("ThisDir") & "'>" & vbNewLine
			For Each DirFiles In fsoFile.Files
				c = c + 1
				If c > maxperpage * (CurrentPage - 1) Then
					Response.Write "<td class=clefttitle style='text-align:left'>"
					Response.Write "<div align=center><a href='" & FilePath & DirFiles.Name & "'target=_blank><img src='" & GetFilePic(FilePath & DirFiles.Name) & "' width=140 height=100 border=0 alt='点此图片查看原始文件！'></a></div>"
					Response.Write "文件名：<a href='" & FilePath & DirFiles.Name & "'target=_blank>" & DirFiles.Name & "</a><br>"
					Response.Write "文件大小：" & GetFileSize(DirFiles.Size) & "<br>"
					Response.Write "文件类型：" & DirFiles.Type & "<br>"
					Response.Write "修改时间：" & DirFiles.DateLastModified & "<br>"
					Response.Write "管理操作：<input type=checkbox name=FileName value='" & DirFiles.Name & "' checked> 选择&nbsp;&nbsp;"
					Response.Write "<a href='?action=del&UploadDir=" & Request("UploadDir") & "&ThisDir=" & Request("ThisDir") & "&FileName=" & DirFiles.Name & "' onclick=""return confirm('您确定要删除此文件吗!');"">×删除</a>"
					FileCount = FileCount + 1
					Response.Write "</td>" & vbNewLine
					If (FileCount Mod 4) = 0 And FileCount < maxperpage And c < totalPut Then
						Response.Write "</tr>" & vbNewLine & "<tr>" & vbNewLine
					End If
				End If
				If FileCount >= maxperpage Then Exit For
			Next
			Response.Write "</tr>" & vbNewLine
			Response.Write "<tr><td colspan=4 class=clefttitle>" & vbNewLine
			Response.Write "<input class=Button type=button name=chkall value='全选' onClick=""CheckAll(this.form)"">&nbsp;<input class=Button type=button name=chksel value='反选' onClick=""ContraSel(this.form)"">" & vbNewLine
			Response.Write "&nbsp;&nbsp;<input class=Button type=submit name=Submit2 value='删除选中的文件' onClick=""return confirm('确定要删除选中的文件吗？')"">" & vbNewLine
			Response.Write "&nbsp;&nbsp;<input class=Button type=submit name=Submit3 value='删除所有文件' onClick=""document.myform.action.value='DelThisAllFile';return confirm('确定要删除当前目录所有文件吗？')"">" & vbNewLine
			Response.Write "&nbsp;&nbsp;<input class=Button type=submit name=Submit4 value='删除所有文件和文件夹' onClick=""document.myform.action.value='DelAllDirFile';return confirm('确定要删除当前目录所文件和文件夹吗？')"">" & vbNewLine
			Response.Write "&nbsp;&nbsp;<input class=Button type=submit name=Submit5 value='删除所有空文件夹' onClick=""document.myform.action.value='DelEmptyFolder';return confirm('确定要删除当前目录所有空文件夹吗？')"">" & vbNewLine
			Response.Write "</tr></form>" & vbNewLine
			Response.Write "<tr><td colspan=4 class=clefttitle>" & vbNewLine
			Response.Write showpage(CurrentPage, totalPut, maxperpage, TotleSize)
			Response.Write "</td></tr>" & vbNewLine
			Response.Write "</table>"
			
			Response.Write " <div style='text-align:center;margin-top:27px'><input class=Button type=button name=Submit2 value=' 一键清理所有未关联的垃圾文件 ' onclick=""if (confirm('您确定要一键清除所有无用的文件吗？此操作不可逆，建议先备份UploadFiles目录后再执行！')){location.href='?action=delete';}""></div>"

		Else
			Response.Write "此目录没有任何文件！"
		End If
		Set fsoFile = Nothing: Set fso = Nothing
	End Sub
	'=================================================
	'过程名：ClearUploadFile
	'作  用：清理无用的上传文件
	'=================================================
	Private Sub ClearUploadFile()
		Response.Write "<html>"
		Response.Write "<head>"
		Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>"
		Response.Write "<title>管理</title>"
		Response.Write "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
		Response.Write "</head>"
		
		Response.Write "<body topmargin='0' leftmargin='0'>"
		Response.Write "<ul id='mt'> <div style='text-align:center;font-weight:bold;margin-top:10px'>"
		If LCase(Request("UploadDir")) = "DownUrl" Then
			Response.Write "清理无用的上传文件"
		Else
			Response.Write "清理无用的上传图片"
		End If
		Response.Write "</div></ul>"
		Response.Write "<table border=0 align=center cellpadding=3 style='margin-top:5px' cellspacing=1 width='99%' class='ctable'>"
		
		Response.Write "<form name=""myform"" method=""get"" action='KS.AdminFiles.asp'>" & vbCrLf
		Response.Write "<input type=hidden name=action value='delete'>" & vbNewLine
		Response.Write "<input type=hidden name=UploadDir value='" & Request("UploadDir") & "'>" & vbNewLine
		Response.Write "<tr><td class=clefttitle>" & vbNewLine
		Response.Write "<br>&nbsp;&nbsp;①、你的网站在使用一段时间后，就会产生大量无用垃圾文件。如果怕浪费您的空间，可以定期使用本功能进行清理；<br>"
		Response.Write "<br>&nbsp;&nbsp;②、新版（V6及以后的版本）上传目录统一命名为UploadFiles,此功能仅对UpLoadFiles目录执行清理功能；<br>"
		Response.Write "<br>&nbsp;&nbsp;③、如果上传文件很多，或者数据库的信息量较多，执行本操作需要耗费相当长的时间，请在访问量少时执行本操作。<br>"
		Response.Write "<br></td></tr>" & vbNewLine
		Response.Write "<tr align=center><td  class=clefttitle>请选择要清理的目录："
		Call ShowFolderPath
		Response.Write "<input class=Button type=submit name=Submit2 value=' 开始清理垃圾文件 ' onclick=""return confirm('您确定要清除所有无用的文件吗？');"">"
		Response.Write " <input class=Button type=button name=Submit2 value=' 一键清理所有垃圾文件 ' onclick=""if (confirm('您确定要一键清除所有无用的文件吗？此操作不可逆，建议先备份UploadFiles目录后再执行！')){location.href='?action=delete';}"">"
		Response.Write " 　　<a href='?'>返回上传管理</a>"
		Response.Write "</td></tr></form>" & vbNewLine
		Response.Write "</table>"
	End Sub
	'=================================================
	'过程名：ShowFolderPath
	'作  用：显示子目录菜单
	'=================================================
	Private Sub ShowFolderPath()
		Dim fso, fsoFile, DirFolder
		Dim strFolderPath
		On Error Resume Next
		strFolderPath = ChannelDir & UploadDir
		strFolderPath = Server.MapPath(strFolderPath)
		Set fso = KS.InitialObject(KS.Setting(99))
		If fso.FolderExists(strFolderPath) Then
			Set fsoFile = fso.GetFolder(strFolderPath)
			Response.Write "<select name=""path"">" & vbNewLine
			For Each DirFolder In fsoFile.SubFolders
			   If IsDate(DirFolder.Name) Then
				Response.Write "	<option value=""" & DirFolder.Name & """>" & DirFolder.Name & "</option>" & vbNewLine
			   End If
			Next
			'Response.Write "	<option value="""">上传根目录</option>" & vbNewLine
			Response.Write "</select>" & vbNewLine
			Set fsoFile = Nothing
		Else
			'Response.Write "没有找到文件夹！"
		End If
		Set fso = Nothing
	End Sub
	
	Sub DeleteFile(strFolderPath,i)
		Dim fso, fsoFile, DirFiles
		Dim strFileName,ParentPath
		Dim strFilePath
		Set fso = KS.InitialObject(KS.Setting(99))
		If fso.FolderExists(strFolderPath) Then
		    Set fsoFile = fso.GetFolder(strFolderPath)
			ParentPath=strFolderPath
			For Each DirFiles In fsoFile.SubFolders
			 Call DeleteFile(ParentPath & "\" & DirFiles.Name,i)
			Next
			
			For Each DirFiles In fsoFile.Files
			    
				strFileName = DirFiles.Name
				strFilePath = strFolderPath & "\" & DirFiles.Name
				If Not CheckFileExists(strFilePath) Then
					i = i + 1
					fso.DeleteFile(strFilePath)
				End If
			Next
			Set fsoFile = Nothing
		End If
		Set fso = Nothing
	End Sub
	'=================================================
	'过程名：DelUselessFile
	'作  用：删除所有无用的上传文件
	'=================================================
	Private Sub DelUselessFile()
		Dim SQL,i
		Dim fso, fsoFile, DirFiles
		Dim strFileName,strFolderPath
		Dim strFilePath,strDirName
		Server.ScriptTimeout = 9999999
		'On Error Resume Next
		i=0
		If Len(Request("path")) > 0 Then
			strDirName = Request("path") & "/"
		Else
			strDirName = vbNullString
		End If
		strFolderPath = ChannelDir & UploadDir & strDirName
		strFolderPath = Server.MapPath(strFolderPath)
		Call DeleteFile(strFolderPath,i)
	

		Response.Write "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
		Response.Write " <Br><br><br><table align='center' width=""95%"" height='200' class='ctable' cellpadding=""1"" cellspacing=""1"">"
		Response.Write "	  <tr class=""sort""> "
		Response.Write "		<td  height=""28"" colspan=2>系统操作提示信息</td>" & vbcrlf
		Response.Write "	  </tr>"
		Response.Write "    <tr class='tdbg'>"
		Response.Write "          <td align='center'><img src='images/succeed.gif'></td>"
		Response.Write "<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<li>文件清理完成！</li><li>一共清理了<font color=red><b>" & i & "</b></font>个垃圾文件！</b><br></td></tr>"
		Response.Write "	  <tr class=""sort""> "
		Response.Write "		<td  height=""28"" colspan=2><a href='#' onclick='javascript:history.back(-1);'>返回上一级</a> <a  href='?'>返回上传目录</a></td>" & vbcrlf
		Response.Write "	  </tr>"
		Response.Write "</table>"
	End Sub
	Public Function CheckFileExists(ByVal str)
	   str=Lcase(str)
	   str=replace(str,"\","/")
	   if instr(str,lcase(KS.Setting(91)))<0 then
	     CheckFileExists=false
		 exit function
	   end if
	   
	   str=Split(str,lcase(KS.Setting(91)))	   
	   If Ubound(Str)=0 Then
	     CheckFileExists=false
		 exit function
	   End If
	   Dim FileName
        
	   FileName=lcase(KS.Setting(91)) & str(1)
	  
		Dim Rs,SQL,Param
		IF INSTR(FileName,"[")<>0 and Instr(FileName,"]")<>0 then
		 FileName=Split(FileName,"[")(0) & "%" & Split(FileName,"]")(1)
		end if
		SQL = "SELECT TOP 1 ID FROM [KS_UploadFiles] WHERE FileName like '%" & FileName & "'"
		Set Rs = Conn.Execute(SQL)
		If Rs.BOF And Rs.EOF Then
			CheckFileExists = False
		Else
			CheckFileExists = True
		End If
		Set Rs = Nothing
	End Function
	
End Class
%>
<%
Const NoAllowExt = "asa|asax|ascs|ashx|asmx|asp|aspx|axd|cdx|cer|cfm|config|cs|csproj|idc|licx|rem|resources|resx|shtm|shtml|soap|stm|vb|vbproj|vsdisco|webinfo"    '不允许上传类型

Const NeedCheckFileMimeExt = "gif|jpg|jpeg|jpe|bmp|png|swf|mid|mp3|wmv|asf|avi|mpg|ram|rm|ra|rar|exe|doc|zip" '定义需要检查是否伪造的文件类型
Dim UpFileStream
Class UpFileClass
	Dim Form,File,Err ,KS
	Private Sub Class_Initialize
		Err = -1
		Set KS=New PublicCls
	End Sub
	Private Sub Class_Terminate  
		'清除变量及对像
		If Err < 0 Then
			Form.RemoveAll
			Set Form = Nothing
			File.RemoveAll
			Set File = Nothing
			UpFileStream.Close
			Set UpFileStream = Nothing
		End If
		Set KS=Nothing
	End Sub
	
	Public Property Get ErrNum()
		ErrNum = Err
	End Property
	
	
	Function UpSave(Path,FileSize,AllowExtStr,sFileName)
			Dim ErrStr,NoUpFileTF,FsoObj,FileName,FileExtName,FileContent,SameFileExistTF,FormName,tempstr,n
			NoUpFileTF = True
			ErrStr = "":N=0
			Set FsoObj = KS.InitialObject(KS.Setting(99))
			For Each FormName in File
				SameFileExistTF = False
				FileName = File(FormName).FileName
				If NoIllegalStr(FileName)=False Then ErrStr=ErrStr&"文件：上传被禁止！\n"
				FileExtName = File(FormName).FileExt
				FileContent = File(FormName).FileData
				Dim FileType:FileType=File(FormName).FileType
				'是否存在重名文件
				if File(FormName).FileSize > 1 then
					NoUpFileTF = False
					ErrStr = ""
					if File(FormName).FileSize > CLng(FileSize)*1024 then
						ErrStr = "errsize"
					end if
					' IF KS.ChkClng(KS.GetFolderSize(KSUser.GetUserFolder(ksuser.username))/1024+UpFileObj.File(FormName).FileSize/1024)>=KS.ChkClng(KSUser.GetUserInfo("SpaceSize")) Then
					'  Response.Write "<script>alert('上传失败1，您的可用空间不够！');history.back();<//script>"
					'  response.end
					'End If
					'if AutoRename = "0" then
					'	If FsoObj.FileExists(Path & FileName) = True  then
					'		ErrStr = ErrStr & FileName & "文件上传失败,存在同名文件\n"
					'	else
					'		SameFileExistTF = True
					'	end if
					'else
						SameFileExistTF = True
					'End If
					
					If Left(LCase(FileType), 5) = "text/" and KS.FoundInArr(NeedCheckFileMimeExt,FileExtName,"|")=true Then
					 KS.AlertHintScript FileName & "文件上传失败\n为了系统安全，不允许上传用文本文件伪造的图片文件！\n"
					 Response.End()
					End If
					If instr(FileName,";")>0 or instr(lcase(FileName),".asp")>0 or instr(lcase(FileName),".php")>0 or instr(lcase(FileName),".cdx")>0 or instr(lcase(FileName),".asa")>0 or instr(lcase(FileName),".cer")>0 or instr(lcase(FileName),".cfm")>0 or instr(lcase(FileName),".jsp")>0 then
					 KS.AlertHintScript FileName & "文件上传失败,文件名不合法\n"
					 Response.End()
					end if
					
					
					
					if CheckFileType(AllowExtStr,FileExtName) = False then
						ErrStr = "errext"
					end if
					if ErrStr <> "" then
						UpSave=ErrStr
						Exit Function
					end if

			        If n=0 Then
						Tempstr=SaveFile(Path,FormName,sFileName)
					Else
						Tempstr=TempStr &"|" & SaveFile(Path,FormName,sFileName&n)
					End If
                else
				    If n=0 Then
						Tempstr=""
					Else
						Tempstr=TempStr &"|"
					End If
				end if
				
                n=n+1
			Next
			
			Set FsoObj = Nothing
			if NoUpFileTF = True then
				UpSave = ""
			Else
			    UpSave = TempStr
			end if
End Function
Function NoIllegalStr(Byval FileNameStr)
			Dim Str_Len,Str_Pos
			Str_Len=Len(FileNameStr)
			Str_Pos=InStr(FileNameStr,Chr(0))
			If Str_Pos=0 or Str_Pos=Str_Len then
				NoIllegalStr=True
			Else
				NoIllegalStr=False
			End If
End function
Function DealExtName(Byval UpFileExt)
			If IsEmpty(UpFileExt) Then Exit Function
			DealExtName = Lcase(UpFileExt)
			DealExtName = Replace(DealExtName,Chr(0),"")
			DealExtName = Replace(DealExtName,".","")
			DealExtName = Replace(DealExtName,"'","")
			DealExtName = Replace(DealExtName,"asp","")
			DealExtName = Replace(DealExtName,"asa","")
			DealExtName = Replace(DealExtName,"aspx","")
			DealExtName = Replace(DealExtName,"cer","")
			DealExtName = Replace(DealExtName,"cdx","")
			DealExtName = Replace(DealExtName,"htr","")
			DealExtName = Replace(DealExtName,"php","")
End Function

Function CheckFileType(AllowExtStr,FileExtName)
			Dim i,AllowArray
			AllowArray = Split(AllowExtStr,"|")
			FileExtName = LCase(FileExtName)
			CheckFileType = False
			For i = LBound(AllowArray) to UBound(AllowArray)
				if LCase(AllowArray(i)) = LCase(FileExtName) then
					CheckFileType = True
				end if
			Next
			If KS.FoundInArr(LCase(NoAllowExt),FileExtName,"|")=true Then
				CheckFileType = False
			end if
End Function
		
Function SaveFile(FilePath,FormNameItem,sFileName)
			Dim FileName,FileExtName,FileContent,FormName,RandomFigure,n,RndStr
			Randomize 
			n=2* Rnd+10
			RndStr=KS.MakeRandom(n)
			RandomFigure = CStr(Int((99999 * Rnd) + 1))
			FileName = File(FormNameItem).FileName
			FileExtName = File(FormNameItem).FileExt
			FileExtName = DealExtName(FileExtName)
			FileContent = File(FormNameItem).FileData

			FileName= sFileName&"."&FileExtName

           ' response.write Server.MapPath(FilePath &FileName) & "<br/>"


			File(FormNameItem).SaveToFile Server.MapPath(FilePath &FileName)
 			'======================增加检查文件内容是否合法===================================
			call CheckFileContent(FilePath &FileName)
			'==================================================================================
			SaveFile=FilePath &FileName

End Function

		'检查文件内容的是否合法
		Function  CheckFileContent(byval path)
			        dim kk,NoAllowExtArr
					NoAllowExtArr=split(NoAllowExt,"|")
					for kk=0 to ubound(NoAllowExtArr)
					   if instr(lcase(path),"." & NoAllowExtArr(kk))<>0 then
					    call KS.DeleteFile(path)
					    ks.die  "<script>alert('文件上传失败,文件名不合法\n');</script>"
					   end if
					Next

		    on error resume next
		    Dim findcontent,regEx,foundtf
			findcontent=KS.ReadFromFile(Replace(path,KS.Setting(2),""))
			if err then exit function : err.clear
			foundtf=false
			Set regEx = New RegExp
			regEx.IgnoreCase = True
			regEx.Global = True
			regEx.Pattern = "@\s*LANGUAGE\s*=\s*[""]?\s*(vbscript|jscript|javascript).encode\b"
			If regEx.Test(findcontent) Then
				foundtf=true
			end if	
			
			regEx.Pattern = "execute\s*request"
			If regEx.Test(findcontent) Then
				foundtf=true
			end if
			
			regEx.Pattern = "executeglobal\s*request"
			If regEx.Test(findcontent) Then
				foundtf=true
			end if
			regEx.Pattern = "<script.*runat.*server(\n|.)*execute(\n|.)*<\/script>"
			If regEx.Test(findcontent) Then
				foundtf=true
			end if
			regEx.Pattern = "\<%(.|\n)*%\>"
			If regEx.Test(findcontent) Then
				foundtf=true
			end if
			If Instr(lcase(findcontent),"scripting.filesystemobject")<>0 or instr(lcase(findcontent),"adodb.stream")<>0 Then
			foundtf=true
			End If
			if foundtf then
			   KS.DeleteFile(path)
			   KS.Die "<script>alert('系统检查到您上传的文件可能存在危险代码，不允许上传！');history.back(-1);</script>"
			end if
			
	  End Function


	
Public Sub GetData ()
		'定义变量
		Dim RequestBinData,sSpace,bCrLf,sObj,iObjStart,iObjEnd,tStream,iStart,oFileObj
		Dim iFileSize,sFilePath,sFileType,sFormValue,sFileName
		Dim iFindStart,iFindEnd
		Dim iFormStart,iFormEnd,sFormName
		'代码开始
		If Request.TotalBytes < 1 Then  '如果没有数据上传
			Err = 1
			Exit Sub
		End If
		Set Form = Server.CreateObject ("Scripting.Dictionary")
		Form.CompareMode = 1
		Set File = Server.CreateObject ("Scripting.Dictionary")
		File.CompareMode = 1
		Set tStream = Server.CreateObject ("ADODB.Stream")
		Set UpFileStream = Server.CreateObject ("ADODB.Stream")
		UpFileStream.Type = 1
		UpFileStream.Mode = 3
		UpFileStream.Open
		UpFileStream.Write (Request.BinaryRead(Request.TotalBytes))
		UpFileStream.Position = 0
		RequestBinData=UpFileStream.Read 
		iFormEnd = UpFileStream.Size
		bCrLf = ChrB (13) & ChrB (10)
		'取得每个项目之间的分隔符
		sSpace=MidB (RequestBinData,1, InStrB (1,RequestBinData,bCrLf)-1)
		iStart=LenB (sSpace)
		iFormStart = iStart+2
		'分解项目
		Do
			iObjEnd=InStrB(iFormStart,RequestBinData,bCrLf & bCrLf)+3
			tStream.Type = 1
			tStream.Mode = 3
			tStream.Open
			UpFileStream.Position = iFormStart
			UpFileStream.CopyTo tStream,iObjEnd-iFormStart
			tStream.Position = 0
			tStream.Type = 2
			tStream.CharSet="utf-8"
			sObj = tStream.ReadText      
			'取得表单项目名称
			iFormStart = InStrB (iObjEnd,RequestBinData,sSpace)-1
			iFindStart = InStr (22,sObj,"name=""",1)+6
			iFindEnd = InStr (iFindStart,sObj,"""",1)
			sFormName = Mid  (sObj,iFindStart,iFindEnd-iFindStart)
			'如果是文件
			If InStr  (45,sObj,"filename=""",1) > 0 Then
				Set oFileObj = new FileObj_Class
				'取得文件属性
				iFindStart = InStr (iFindEnd,sObj,"filename=""",1)+10
				iFindEnd = InStr (iFindStart,sObj,"""",1)
				sFileName = Mid (sObj,iFindStart,iFindEnd-iFindStart)
				oFileObj.FileName = Mid (sFileName,InStrRev (sFileName, "\")+1)
				oFileObj.FilePath = Left (sFileName,InStrRev (sFileName, "\"))
				oFileObj.FileExt = Mid (sFileName,InStrRev (sFileName, ".")+1)
				iFindStart = InStr (iFindEnd,sObj,"Content-Type: ",1)+14
				iFindEnd = InStr (iFindStart,sObj,vbCr)
				oFileObj.FileType = Mid  (sObj,iFindStart,iFindEnd-iFindStart)
				oFileObj.FileStart = iObjEnd
				oFileObj.FileSize = iFormStart -iObjEnd -2
				oFileObj.FormName = sFormName
				File.add sFormName,oFileObj
			else
				'如果是表单项目
				tStream.Close
				tStream.Type = 1
				tStream.Mode = 3
				tStream.Open
				UpFileStream.Position = iObjEnd 
				UpFileStream.CopyTo tStream,iFormStart-iObjEnd-2
				tStream.Position = 0
				tStream.Type = 2
				tStream.CharSet="utf-8"
				sFormValue = tStream.ReadText
				If Form.Exists(sFormName)Then
					Form (sFormName) = Form (sFormName) & ", " & sFormValue
				else
					form.Add sFormName,sFormValue
				End If
			End If
			tStream.Close
			iFormStart = iFormStart+iStart+2
			'如果到文件尾了就退出
		Loop Until  (iFormStart+2) >= iFormEnd 
		RequestBinData = ""
		Set tStream = Nothing
	End Sub
End Class

'----------------------------------------------------------------------------------------------------
'文件属性类
Class FileObj_Class
	Dim FormName,FileName,FilePath,FileSize,FileType,FileStart,FileExt
	'保存文件方法
	Public Function SaveToFile (Path)
		On Error Resume Next
		Dim oFileStream
		Dim KS:Set KS=New PublicCls
		Set oFileStream = KS.InitialObject ("ADODB.Stream")
		oFileStream.Type = 1
		oFileStream.Mode = 3
		oFileStream.Open
		UpFileStream.Position = FileStart
		UpFileStream.CopyTo oFileStream,FileSize
		oFileStream.SaveToFile Path,2
		oFileStream.Close
		Set oFileStream = Nothing 
		Set KS=Nothing
	End Function
	'取得文件数据
	Public Function FileData
		UpFileStream.Position = FileStart
		FileData = UpFileStream.Read (FileSize)
	End Function
End Class
%>
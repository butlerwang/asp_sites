<%
'KesionCMS V8 无组件上类处理类及函数 修改于2010-6-28 by xiaolin
Const NoAllowExt = "asa|asax|ascs|ashx|asmx|asp|aspx|axd|cdx|cer|cfm|config|cs|csproj|idc|licx|rem|resources|resx|shtm|shtml|soap|stm|vb|vbproj|vsdisco|webinfo"    '不允许上传类型
Const NeedCheckFileMimeExt = "gif|jpg|jpeg|jpe|bmp|png|swf|mid|mp3|wmv|asf|avi|mpg|ram|rm|ra|rar|exe|doc|zip" '定义需要检查是否伪造的文件类型

Dim KS:Set KS=New PublicCls

'添加附件
Sub AddAnnexToDB(ChannelID,Username,TempFileStr,ByVal FileTitles,ClassID,ShowJS,EditorID)
	'写入KS_UploadFiles数据库
	Dim FileArr,n,FileIDS,MaxID,TitleArr
	FileArr=Split(TempFileStr,"|")
	TitleArr=Split(Replace(FileTitles,"'",""),"|")
	For N=0 To Ubound(FileArr)
	  If Not KS.IsNul(FileArr(n)) Then
	       'If Right(lcase(FileArr(n)),3)<>"gif" and Right(lcase(FileArr(n)),3)<>"bmp" and Right(lcase(FileArr(n)),3)<>"jpg" and Right(lcase(FileArr(n)),3)<>"png" and Right(lcase(FileArr(n)),4)<>"jpeg" Then
								 Conn.Execute("Insert Into [KS_UploadFiles](ChannelID,InfoID,Title,FileName,IsAnnex,UserName,Hits,AddDate,ClassID) values(" &ChannelID &",0,'" & TitleArr(n) & "','" & FileArr(n) & "',1,'" & UserName & "',0," & SQLNowString&"," & ClassID & ")")
								 MaxID=Conn.Execute("Select Max(ID) From  [KS_UploadFiles]")(0)
								 If FileIds="" Then
								   FileIds=MaxID
								 Else
								   FileIds=FileIds & "," & MaxID
								 End If
			'Else
			'   MaxID=0
			'End If

		 If ShowJS=False Then
		 KS.Echo  escape(FileArr(n)) & "|" & KS.GetFieSize(Server.MapPath(Replace(FileArr(n),KS.Setting(2),""))) & "|" & MaxID & "|" & escape(TitleArr(n)) & "|" & EditorID
		 Else
		 Response.Write("parent.InsertFileFromUp('" & FileArr(n) &"'," & KS.GetFieSize(Server.MapPath(Replace(FileArr(n),KS.Setting(2),""))) & "," & MaxID & ",'" & TitleArr(n) & "','" & EditorID &"');")
		 End If
	  End If
	Next
	If Session("UploadFileIDs")="" Then
	  Session("UploadFileIDs")=FileIds
	Else
	  Session("UploadFileIDs")=Session("UploadFileIDs") & "," & FileIds
	End If
End Sub

Function CheckUpFile(KSUser,MustCheckSpaceSize,UpFileObj,FormPath,Path,FileSize,AllowExtStr,ByRef U_FileSize,ByRef TempFileStr,ByRef FileTitles,ByRef CurrNum,CreateThumbsFlag,DefaultThumb,ThumbPathFileName)
			Dim ErrStr,NoUpFileTF,FsoObj,FileName,FileExtName,FileContent,SameFileExistTF,FormName,AutoReName,BasicType
			AutoReName = KS.ChkClng(UpFileObj.Form("AutoRename"))
			BasicType=KS.ChkClng(UpFileObj.Form("BasicType")) 
			NoUpFileTF = True
			ErrStr = ""
			Set FsoObj = KS.InitialObject(KS.Setting(99))
			For Each FormName in UpFileObj.File
				SameFileExistTF = False
				FileName = UpFileObj.File(FormName).FileName
				If NoIllegalStr(FileName)=False Then ErrStr=ErrStr&"文件：上传被禁止！\n"
				FileExtName = UpFileObj.File(FormName).FileExt
				If Instr(FileExtName,"?")<>0 Then FileExtName=right("00" & FileExtName,4) '防止swfupload的中文乱码处理
				If Instr(FileExtName,"?")<>0 Then FileExtName=right("00" & FileExtName,3)
				FileContent = UpFileObj.File(FormName).FileData
				U_FileSize=UpFileObj.File(FormName).FileSize
				Dim FileType:FileType=UpFileObj.File(FormName).FileType
				'是否存在重名文件
				if U_FileSize > 1 then
					NoUpFileTF = False
					ErrStr = ""
					if U_FileSize > CLng(FileSize)*1024 then
						ErrStr = ErrStr & FileName & "文件上传失败\n超过了限制，最大只能上传" & FileSize & "K的文件\n"
					end if
					
					If MustCheckSpaceSize=true Then
						If BasicType<>9994 Then
						 IF KS.ChkClng(KS.GetFolderSize(KSUser.GetUserFolder(ksuser.getuserinfo("userid")))/1024+UpFileObj.File(FormName).FileSize/1024)>=KS.ChkClng(KSUser.GetUserInfo("SpaceSize")) Then
						   CheckUpFile="上传失败1，您的可用空间不够！"
						   Exit Function
						 End If
						End If
					End If
					
					if AutoRename = "0" then
						If FsoObj.FileExists(Path & FileName) = True  then
							ErrStr = ErrStr & FileName & "文件上传失败,存在同名文件\n"
						else
							SameFileExistTF = True
						end if
					else
						SameFileExistTF = True
					End If
					if CheckFileType(AllowExtStr,FileExtName) = False then
						ErrStr = ErrStr & FileName & "文件上传失败,文件类型不允许\n允许的类型有" + AllowExtStr + "\n"
					end if
					If Left(LCase(FileType), 5) = "text/" and KS.FoundInArr(NeedCheckFileMimeExt,FileExtName,"|")=true Then
					 ErrStr = ErrStr & FileName & "文件上传失败\n为了系统安全，不允许上传用文本文件伪造的图片文件！\n"
					End If
					If instr(FileName,";")>0 or instr(lcase(FileName),".asp")>0 or instr(lcase(FileName),".php")>0 or instr(lcase(FileName),".cdx")>0 or instr(lcase(FileName),".asa")>0 or instr(lcase(FileName),".cer")>0 or instr(lcase(FileName),".cfm")>0 or instr(lcase(FileName),".jsp")>0 then
						ErrStr = ErrStr & FileName & "文件上传失败,文件名不合法\n"
					end if
					
					if ErrStr = "" then
						if SameFileExistTF = True then
							CheckUpFile = CheckUpFile & SaveFile(KSUser,UpFileObj,FormPath,Path,FormName,AutoReName,TempFileStr,FileTitles,CurrNum,CreateThumbsFlag,DefaultThumb,ThumbPathFileName)
						else
							CheckUpFile = CheckUpFile &SaveFile(KSUser,UpFileObj,FormPath,Path,FormName,"",TempFileStr,FileTitles,CurrNum,CreateThumbsFlag,DefaultThumb,ThumbPathFileName)

						end if
					else
						CheckUpFile = CheckUpFile & ErrStr
					end if
				end if
			Next
			Set FsoObj = Nothing
			if NoUpFileTF = True then
				CheckUpFile = "没有上传文件"
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

Function SaveFile(KSUser,UpFileObj,FormPath,FilePath,FormNameItem,AutoNameType,TempFileStr,FileTitles,CurrNum,CreateThumbsFlag,DefaultThumb,ThumbPathFileName)
			Dim FileName,FileExtName,FileContent,FormName,RandomFigure,n,RndStr,Title,BasicType,ChannelID,UpType,ThumbFileName,AddWaterFlag
		    BasicType=KS.ChkClng(UpFileObj.Form("BasicType"))      
		    ChannelID=KS.ChkClng(UpFileObj.Form("ChannelID")) 
			UpType=UpFileObj.Form("UpType")
			AddWaterFlag = UpFileObj.Form("AddWaterFlag")
			If AddWaterFlag <> "1" Then	'生成是否要添加水印标记
				AddWaterFlag = "0"
			End if

			Randomize 
			n=2* Rnd+10
			RndStr=KS.MakeRandom(n)
			RandomFigure = CStr(Int((99999 * Rnd) + 1))
			FileName = UpFileObj.File(FormNameItem).FileName
			FileExtName = UpFileObj.File(FormNameItem).FileExt
			FileExtName = DealExtName(FileExtName)
			FileContent = UpFileObj.File(FormNameItem).FileData
			
			If Instr(FileExtName,"?")<>0 Then FileExtName=right("00" & FileExtName,4)  '防止swfupload的中文乱码处理
			If Instr(FileExtName,"?")<>0 Then FileExtName=right("00" & FileExtName,3)

			Title=replace(FileName,"." & FileExtName,"") '原名称
			If BasicType=9999 Then   '头像
			   FileName=KSUser.GetUserInfo("UserID") & ".jpg"
			Else
				select case AutoNameType 
				  case "1"
					FileName= "副件" & FileName
				  case "2"
					FileName= RndStr&"."&FileExtName
				  Case "3"
					FileName= RndStr & FileName
				  case "4"
					FileName= Year(Now())&Right("0"&Month(Now()),2)&Right("0"&Day(Now()),2)&Right("0"&Hour(Now()),2)&Right("0"&Minute(Now()),2)&Right("0"&Second(Now()),2)&RandomFigure&"."&FileExtName
				  case else
					FileName=FileName
				End Select
			End If
		   UpFileObj.File(FormNameItem).SaveToFile FilePath  &FileName
		   
		   
		   
		   '======================增加检查文件内容是否合法===================================
		   Dim CheckContent:CheckContent=CheckFileContent(FormPath  &FileName,UpFileObj.File(FormNameItem).FileSize /1024)
		   If KS.IsNul(CheckContent) Then
			'==================================================================================
			
		   
		   TempFileStr=TempFileStr & FormPath & FileName & "|"
		   FileTitles=FileTitles & Title & "|"
		  Dim T:Set T=New Thumb
		  CurrNum=CurrNum+1
		  IF CreateThumbsFlag=true and  (cint(CurrNum)=cint(DefaultThumb) or BasicType=2 or (Channelid=5 and UpType="ProImage")) Then
		  	  If KS.TBSetting(0)=0 then
			   if ThumbPathFileName="" then
			   ThumbPathFileName=FormPath &FileName
			   Else
			   ThumbPathFileName=ThumbPathFileName & "|" & FormPath & FileName
			   End If
			  Else
				ThumbFileName=split(FileName,".")(0)&"_S."&FileExtName
				Dim CreateTF:CreateTF=T.CreateThumbs(FilePath & FileName,FilePath & ThumbFileName)
				if CreateTF=true Then
				 '取得缩略图地址
				 if ThumbPathFileName="" then
				 ThumbPathFileName=FormPath & ThumbFileName
				 else
				 ThumbPathFileName=ThumbPathFileName & "|" & FormPath & ThumbFileName
				end if
			   Else
				 '取得缩略图地址
				 if ThumbPathFileName="" then
				 ThumbPathFileName=FormPath & FileName
				 else
				 ThumbPathFileName=ThumbPathFileName & "|" & FormPath & FileName
				 end if
			   End If
			  End If
		  End if
		  If AddWaterFlag = "1" Then   '在保存好的图片上添加水印
				call T.AddWaterMark(FilePath  & FileName)
		  End if
		  Set T=Nothing
		  
		'======================增加检查文件内容是否合法===================================
	     Else
		  SaveFile=CheckContent
		 End If
		'==================================================================================
End Function

'检查文件内容的是否合法
Function  CheckFileContent(byval path,byval filesize)
		     dim kk,NoAllowExtArr
			 path=Replace(path,KS.Setting(2),"")
			 NoAllowExtArr=split(NoAllowExt,"|")
			 for kk=0 to ubound(NoAllowExtArr)
					   if instr(replace(lcase(path),lcase(KS.Setting(2)),""),"." & NoAllowExtArr(kk))<>0 then
					    call KS.DeleteFile(path)
					    CheckFileContent= "文件上传失败,文件名不合法"
						Exit Function
					   end if
			 Next

		    if filesize>50 then exit function  '超过1000K跳过检测
		    on error resume next
		    Dim findcontent,regEx,foundtf
			findcontent=KS.ReadFromFile(Replace(path,KS.Setting(2),""))
			if err then exit function:err.clear
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
			
			set regEx=nothing
			
			if foundtf then
			   KS.DeleteFile(path)
			   CheckFileContent="系统检查到您上传的文件可能存在危险代码，不允许上传！"
			end if
			
End Function


Dim UpFileStream
Class UpFileClass
	Dim Form,File,Err 
	Private Sub Class_Initialize
		Err = -1
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
	End Sub
	
	Public Property Get ErrNum()
		ErrNum = Err
	End Property
	
	Public Sub GetData ()
		'定义变量
		Dim RequestBinData,sSpace,bCrLf,sObj,iObjStart,iObjEnd,tStream,iStart,oFileObj
		Dim iFileSize,sFilePath,sFileType,sFormValue,sFileName
		Dim iFindStart,iFindEnd
		Dim iFormStart,iFormEnd,sFormName
		Dim KS:Set KS=New PublicCls
		'代码开始
		If Request.TotalBytes < 1 Then  '如果没有数据上传
			Err = 1
			Exit Sub
		End If
		Set Form = KS.InitialObject ("Scripting.Dictionary")
		Form.CompareMode = 1
		Set File = KS.InitialObject ("Scripting.Dictionary")
		File.CompareMode = 1
		Set tStream = KS.InitialObject ("ADODB.Stream")
		Set UpFileStream = KS.InitialObject ("ADODB.Stream")
		UpFileStream.Type = 1
		UpFileStream.Mode = 3
		UpFileStream.Open
		dim ReadedBytes,ChunkBytes
		ReadedBytes=0
		ChunkBytes=1024*100 '100K分块上传方案 
		Do   While   ReadedBytes   <   Request.TotalBytes   
		UpFileStream.Write   Request.BinaryRead(ChunkBytes)    
		ReadedBytes   =   ReadedBytes   +   ChunkBytes   
		If   ReadedBytes   >   Request.TotalBytes   Then   ReadedBytes   =   Request.TotalBytes   
		Loop
			
		'UpFileStream.Write (Request.BinaryRead(Request.TotalBytes))
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
		Set KS=Nothing
	End Sub
End Class

'----------------------------------------------------------------------------------------------------
'文件属性类
Class FileObj_Class
	Dim FormName,FileName,FilePath,FileSize,FileType,FileStart,FileExt
	'保存文件方法
	Public Function SaveToFile (Path)
		On Error Resume Next
		Dim KS:Set KS=New PublicCls
		Dim oFileStream
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
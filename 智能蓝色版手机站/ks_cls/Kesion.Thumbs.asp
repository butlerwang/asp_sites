<%

Class Thumb
		 Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		' Call CloseConn()
		 Set KS=Nothing
		End Sub
		'为图片添加水印
		Function AddWaterMark(FileName)
			Dim objFileSystem, strFileExtName, objImage
			If InStr(FileName, ":") = 0 Then                                            
				FileName = Server.MapPath(FileName)
			End If
			If FileName <> "" And Not IsNull(FileName) Then                           
				strFileExtName = ""
				If InStr(FileName, ".") <> 0 Then
					strFileExtName = LCase(Trim(Mid(FileName, InStrRev(FileName, ".") + 1)))
				End If
				If strFileExtName <> "jpg" And strFileExtName <> "gif" And strFileExtName <> "bmp" And strFileExtName <> "png" Then 
					Exit Function
				End If
				Set objFileSystem = KS.InitialObject(KS.Setting(99))
				If objFileSystem.FileExists(FileName) Then             
					If KS.TbSetting(5) <> "0" Then                   
						Select Case KS.TbSetting(5)
							Case "1"                                                           
								If KS.IsObjInstalled("Persits.Jpeg") Then                    
									If KS.IsExpired("Persits.Jpeg") Then
										Response.Write ("对不起，Persits.Jpeg组件已过期!")
										Response.End
									End If
									If KS.TbSetting(6) = "1" Then             
										AddWordMark 1, KS.TbSetting(8), KS.TbSetting(10), KS.TbSetting(11), KS.TbSetting(12), KS.TbSetting(9), KS.TbSetting(7), FileName
									Else                                               
										AddPhotoMark 1, KS.TbSetting(16), KS.TbSetting(17), KS.TbSetting(13), KS.TbSetting(14), KS.TbSetting(15), KS.TbSetting(7), FileName
									End If
								End If
							Case "2"                                                  
								If strFileExtName = "png" Then                        
									Exit Function
								End If
								If KS.IsObjInstalled("wsImage.Resize") Then             
									If KS.IsExpired("wsImage.Resize") Then
										Response.Write ("对不起，sImage.Resize组件已过期!")
										Response.End
									End If
									If KS.TbSetting(6) = "1" Then             
										AddWordMark 2, KS.TbSetting(8), KS.TbSetting(10), KS.TbSetting(11), KS.TbSetting(12), KS.TbSetting(9), KS.TbSetting(7), FileName
									Else                                               
										AddPhotoMark 2, KS.TbSetting(16), KS.TbSetting(17), KS.TbSetting(13), KS.TbSetting(14), KS.TbSetting(15), KS.TbSetting(7), FileName
									End If
								End If
							Case "3"                                                    
								If KS.IsObjInstalled("SoftArtisans.ImageGen") Then           
									If KS.IsExpired("SoftArtisans.ImageGen") Then
										Response.Write ("对不起，SoftArtisans.ImageGen组件已过期!")
										Response.End
									End If
									If KS.TbSetting(6) = "1" Then             
										AddWordMark 3, KS.TbSetting(8), KS.TbSetting(10), KS.TbSetting(11), KS.TbSetting(12), KS.TbSetting(9), KS.TbSetting(7), FileName
									Else                                               
										AddPhotoMark 3, KS.TbSetting(16), KS.TbSetting(17), KS.TbSetting(13), KS.TbSetting(14), KS.TbSetting(15), KS.TbSetting(7), FileName
									End If
								End If
						End Select
					End If
				End If
				Set objFileSystem = Nothing
			End If
		End Function
		'为图片添加文字水印
		Function AddWordMark(MarkComponentID, MarkText, MarkFontColor, MarkFontName, MarkFontBond, MarkFontSize, MarkPosition, FileName)
			Dim objImage, x, y, Text, TextWidth, FontColor, FontName, FondBond, FontSize, OriginalWidth, OriginalHeight
			If InStr(FileName, ":") = 0 Then                                                            
				FileName = Server.MapPath(FileName)
			End If
				
			Text = Trim(MarkText)
			If Text = "" Then
				Exit Function
			End If
			FontColor = Replace(MarkFontColor, "#", "&H")
			FontName = MarkFontName
			If MarkFontBond = "1" Then
				FondBond = True
			Else
				FondBond = False
			End If

			FontSize = CInt(MarkFontSize)
		
			Select Case MarkComponentID
				Case 1
				
					If Not KS.IsObjInstalled("Persits.Jpeg") Then
						Exit Function
					End If
					Set objImage = KS.InitialObject("Persits.Jpeg")
					objImage.Open FileName
					objImage.Canvas.Font.Color = FontColor
					objImage.Canvas.Font.Family = FontName
					objImage.Canvas.Font.Bold = FondBond

					objImage.Canvas.Font.size = FontSize
					on error resume next
					TextWidth = objImage.Canvas.GetTextExtent(Text)  
					if err then err.clear:TextWidth =200
                                  

					If objImage.OriginalWidth < TextWidth Or objImage.OriginalHeight < FontSize Then    
						Exit Function
					End If
					GetPostion CInt(MarkPosition), x, y, objImage.OriginalWidth, objImage.OriginalHeight, TextWidth, FontSize
					
					With objImage.Canvas
					  .Print x, y, Text
					End With
                    objImage.Quality=80
								   			
 
					objImage.Save FileName
		
				Case 2
					If Not KS.IsObjInstalled("wsImage.Resize") Then
						Exit Function
					End If
					Set objImage = KS.InitialObject("wsImage.Resize")
					objImage.LoadSoucePic CStr(FileName)
					objImage.TxtMarkFont = CStr(FontName)
					objImage.TxtMarkBond = FondBond
					objImage.TxtMarkHeight = FontSize
					FontColor = "&H" & Mid(FontColor, 7) & Mid(FontColor, 5, 2) & Mid(FontColor, 3, 2)  
					objImage.AddTxtMark CStr(FileName), CStr(Text), CLng(FontColor), 1, 1
				Case 3
					If Not KS.IsObjInstalled("SoftArtisans.ImageGen") Then
						Exit Function
					End If
					Set objImage = KS.InitialObject("SoftArtisans.ImageGen")
					objImage.LoadImage FileName
					objImage.Font.Height = FontSize
					objImage.Font.name = FontName
					FontColor = "&H" & Mid(FontColor, 7) & Mid(FontColor, 5, 2) & Mid(FontColor, 3, 2)  
					objImage.Font.Color = CLng(FontColor)
					objImage.Text = Text
					GetPostion CInt(MarkPosition), x, y, objImage.Width, objImage.Height, objImage.TextWidth, objImage.TextHeight 
					objImage.DrawTextOnImage x, y, objImage.TextWidth, objImage.TextHeight
					objImage.SaveImage 0, objImage.ImageFormat, FileName
			End Select
			Set objImage = Nothing
		End Function

		Function AddPhotoMark(MarkComponentID, MarkWidth, MarkHeight, MarkPicture, MarkOpacity, MarkTranspColor, MarkPosition, FileName)
			Dim objImage, objMark, x, y, OriginalWidth, OriginalHeight, Position
			If InStr(FileName, ":") = 0 Then                                                            
				FileName = Server.MapPath(FileName)
			End If
			If IsNull(MarkWidth) Or MarkWidth = "" Then
				MarkWidth = 0
			Else
				MarkWidth = CInt(MarkWidth)
			End If
			If IsNull(MarkHeight) Or MarkHeight = "" Then
				MarkHeight = 0
			Else
				MarkHeight = CInt(MarkHeight)
			End If
			If Trim(MarkPicture) = "" Or IsNull(MarkPicture) Then
				Exit Function
			End If
			If IsNull(MarkOpacity) Or MarkOpacity = "" Then
				MarkOpacity = 1
			Else
				MarkOpacity = CSng(MarkOpacity)
			End If
			If MarkTranspColor <> "" Then                                                              
				MarkTranspColor = Replace(MarkTranspColor, "#", "&H")
			Else
			End If
			Select Case MarkComponentID
				Case 1
					If Not KS.IsObjInstalled("Persits.Jpeg") Then
						Exit Function
					End If
					Set objImage = KS.InitialObject("Persits.Jpeg")
					Set objMark = KS.InitialObject("Persits.Jpeg")
					objImage.Open FileName
					If objImage.OriginalWidth < MarkWidth Or objImage.OriginalHeight < MarkHeight Then 
						Exit Function
					End If
					objMark.Open Server.MapPath(MarkPicture)
					
					'objImage.Canvas.DrawImage 0,objImage.OriginalHeight/2-33,objMark,0.6,&HFFFFFF
					
					GetPostion CInt(MarkPosition), x, y, objImage.OriginalWidth, objImage.OriginalHeight, MarkWidth, MarkHeight 
					If MarkTranspColor <> "" Then
						objImage.Canvas.DrawImage x, y, objMark, MarkOpacity, MarkTranspColor
						'objImage.Canvas.DrawImage x, y, objMark, MarkOpacity,&HFFFFFF
					Else
						objImage.DrawImage x, y, objMark, MarkOpacity
					End If
                    objImage.Quality=80  
					objImage.Save FileName
				Case 2
					If Not KS.IsObjInstalled("wsImage.Resize") Then
						Exit Function
					End If
					Set objImage = KS.InitialObject("wsImage.Resize")
					objImage.LoadSoucePic CStr(FileName)
					objImage.LoadImgMarkPic Server.MapPath(MarkPicture)
					objImage.GetSourceInfo OriginalWidth, OriginalHeight
					GetPostion CInt(MarkPosition), x, y, OriginalWidth, OriginalHeight, MarkWidth, MarkHeight
					If MarkTranspColor = "" Then
						MarkTranspColor = 0
					Else
						MarkTranspColor = "&H" & Mid(MarkTranspColor, 7) & Mid(MarkTranspColor, 5, 2) & Mid(MarkTranspColor, 3, 2)
					End If
					objImage.AddImgMark CStr(FileName), Int(x), Int(y), CLng(MarkTranspColor), Int(CSng(MarkOpacity) * 100)
				Case 3
					If Not KS.IsObjInstalled("SoftArtisans.ImageGen") Then
						Exit Function
					End If
					Set objImage = KS.InitialObject("SoftArtisans.ImageGen")
					objImage.LoadImage FileName
					Select Case CInt(MarkPosition)
						Case 1
							Position = 3
						Case 2
							Position = 5
						Case 3
							Position = 1
						Case 4
							Position = 6
						Case 5
							Position = 8
					End Select
					If MarkTranspColor <> "" Then
						MarkTranspColor = "&H" & Mid(MarkTranspColor, 7) & Mid(MarkTranspColor, 5, 2) & Mid(MarkTranspColor, 3, 2)
						objImage.AddWaterMark Server.MapPath(MarkPicture), Position, CSng(MarkOpacity), CLng(MarkTranspColor)
					Else
						objImage.AddWaterMark Server.MapPath(MarkPicture), Position, CSng(MarkOpacity)
					End If
					objImage.SaveImage 0, objImage.ImageFormat, FileName
			End Select
			Set objImage = Nothing
			Set objMark = Nothing
		End Function
		Function GetPostion(MarkPosition, x, y, ImageWidth, ImageHeight, MarkWidth, MarkHeight)
			Select Case CInt(MarkPosition)
				Case 1
					x = 1
					y = 1
				Case 2
					x = 1
					y = Int(ImageHeight - MarkHeight - 1)
				Case 3
					x = Int((ImageWidth - MarkWidth) / 2)
					y = Int((ImageHeight - MarkHeight) / 2)
				Case 4
					x = Int(ImageWidth - MarkWidth - 1)
					y = 1
				Case 5
					x = Int(ImageWidth - MarkWidth - 1)
					y = Int(ImageHeight - MarkHeight - 1)
			End Select
		End Function
		'由原图片根据数据里保存的设置生成缩略图
		Function CreateThumbs(ByVal FileName, ByVal ThumbFileName)
			CreateThumbs = False
			If KS.TbSetting(0) <> "0" And (Not IsNull(KS.TbSetting(0))) Then
				If KS.TbSetting(1) = "0" Then
				   Dim ThumbnailsConfig,Width,Height,GoldenPoint
				   ThumbnailsConfig= Session("ThumbnailsConfig")
				   If ThumbnailsConfig="" Then
				    GoldenPoint= Round(KS.TbSetting(18))
				    Width=CInt(KS.TbSetting(2))
					Height= CInt(KS.TbSetting(3))
				   Else
				    ThumbnailsConfig=Split(ThumbnailsConfig,"|")
					If Not IsNumeric(ThumbnailsConfig(0)) Then
					 GoldenPoint= 0
					Else
				     GoldenPoint= Round(ThumbnailsConfig(0))
					End If
					If Not IsNumeric(ThumbnailsConfig(1)) Then
					 Width=100
					Else
   				     Width=CInt(ThumbnailsConfig(1))
					End If
					If Not IsNumeric(ThumbnailsConfig(2)) Then
					 Height=80
					Else
					 Height= CInt(ThumbnailsConfig(2))
					End If
				   End If
					CreateThumbs = CreateThumb(FileName,Width ,Height,GoldenPoint, 0, ThumbFileName)
				Else
					CreateThumbs = CreateThumb(FileName, 0, 0, GoldenPoint,CSng(KS.TbSetting(4)), ThumbFileName)
				End If
			End If
		End Function
		'由原图片生成指定宽度和高度的缩略图
		Function CreateThumb(FileName, Width, Height,GoldenPoint, Rate, ThumbFileName)
		    'On Error Resume Next
			Dim strSql, RsSetting, objImage, iWidth, iHeight, strFileExtName
			CreateThumb = False
			If IsNull(FileName) Then                                    '如果原图片未指定直接退出
				Exit Function
			ElseIf FileName = "" Then
				Exit Function
			End If
			If InStr(FileName, ".") <> 0 Then
				strFileExtName = LCase(Trim(Mid(FileName, InStrRev(FileName, ".") + 1)))
			End If
			If strFileExtName <> "jpg" And strFileExtName <> "gif" And strFileExtName <> "bmp" And strFileExtName <> "png" Then '文件不是可用图片则退出
				Exit Function
			End If
			If IsNull(ThumbFileName) Then                          
				Exit Function
			ElseIf ThumbFileName = "" Then
				Exit Function
			End If
			If IsNull(Width) Then                                
				Width = 0
			ElseIf Width = "" Then
				Width = 0
			End If
			If IsNull(Rate) Then                                   
				Rate = 0
			ElseIf Rate = "" Then
				Rate = 0
			End If
			If IsNull(Height) Then                               
				Height = 0
			ElseIf Height = "" Then
				Height = 0
			End If
			If InStr(FileName, ":") = 0 Then      
				FileName = Server.MapPath(FileName)
			End If
			If InStr(ThumbFileName, ":") = 0 Then
				ThumbFileName = Server.MapPath(ThumbFileName)
			End If
			 '------检查原图是否存在---
			 Dim FsoObj:Set FsoObj = Server.CreateObject(KS.Setting(99))
			 If Not FsoObj.FileExists(FileName) Then Exit Function
			 SET FsoObj=Nothing
			
			Width = CInt(Width)
			Height = CInt(Height)
			Rate = CSng(Rate)
			
			Select Case CInt(KS.TbSetting(0))
				Case 0                                               
					Exit Function
				Case 1
					If Not KS.IsObjInstalled("Persits.Jpeg") Then
						Exit Function
					End If
					If KS.IsExpired("Persits.Jpeg") Then
						Response.Write ("对不起，Persits.Jpeg组件已过期！")
						Response.End
					End If
					Set objImage = KS.InitialObject("Persits.Jpeg")
					objImage.Open FileName
					If Rate = 0 And (Width <> 0 Or Height <> 0) Then
						If Width < objImage.OriginalWidth And Height < objImage.OriginalHeight And Height<>0 Then
							dim qjazhro_h,qjazhro_w,qjazhro_t,qjazhro_hj,qjazhro,mznvhai 
						    qjazhro=round((Width/Height),3)
						    mznvhai=round((objImage.OriginalWidth/objImage.OriginalHeight),3)
						    If qjazhro<mznvhai Then
							objImage.Height = Height
							objImage.Width = round((objImage.OriginalWidth / objImage.OriginalHeight * Height),3)
							qjazhro_w=round(((objImage.Width-Width)/2),3)
							qjazhro_t=Width+qjazhro_w
							objImage.crop qjazhro_w,0,qjazhro_t,Height
						   ElseIf qjazhro>mznvhai Then
							objImage.Width = Width
							objImage.Height = round((objImage.OriginalHeight / objImage.OriginalWidth * Width),3)
							qjazhro_h=objImage.Height-Height
							qjazhro_hj=qjazhro_h*GoldenPoint  'GoldenPoint为黄金分割点，你可以按自己的要求修改这个值
							qjazhro_t=Height+qjazhro_hj
							objImage.crop 0,qjazhro_hj,Width,qjazhro_t
						   ElseIf qjazhro=mznvhai Then
							objImage.Width = Width
							objImage.Height = Height
							End If
						End If
						
						If Height=0 Then      '当高度为0时,自适应高度
						 Height=Width * objImage.OriginalHeight / objImage.OriginalWidth
						 objImage.Height=Height
						 objImage.Width=Width
						End If
						
					ElseIf Rate <> 0 Then
						objImage.Width = objImage.OriginalWidth * Rate
						objImage.Height = objImage.OriginalHeight * Rate
					End If
					objImage.Interpolation=0
                    objImage.Quality=80  
					objImage.Save ThumbFileName
				Case 2
					If Not KS.IsObjInstalled("wsImage.Resize") Then  
						Exit Function
					End If
					If KS.IsExpired("wsImage.Resize") Then
						Response.Write ("对不起，wsImage.Resize组件已过期！")
						Response.End
					End If
					If strFileExtName = "png" Then   
						Exit Function
					End If
					Set objImage = KS.InitialObject("wsImage.Resize")
					objImage.LoadSoucePic CStr(FileName)
					If Rate = 0 And (Width <> 0 Or Height <> 0) Then
						objImage.GetSourceInfo iWidth, iHeight
						If Width < iWidth And Height < iHeight Then
							If Width = 0 And Height <> 0 Then
								objImage.OutputSpic CStr(ThumbFileName), 0, Height, 2
							ElseIf Width <> 0 And Height = 0 Then
								objImage.OutputSpic CStr(ThumbFileName), Width, 0, 1
							ElseIf Width <> 0 And Height <> 0 Then
								objImage.OutputSpic CStr(ThumbFileName), Width, Height, 0
							Else
								objImage.OutputSpic CStr(ThumbFileName), 1, 1, 3
							End If
						Else
							objImage.OutputSpic CStr(ThumbFileName), 1, 1, 3
						End If
					ElseIf Rate <> 0 Then
						objImage.OutputSpic CStr(ThumbFileName), Rate, Rate, 3
					Else
						objImage.OutputSpic CStr(ThumbFileName), 1, 1, 3
					End If
				Case 3
					If Not KS.IsObjInstalled("SoftArtisans.ImageGen") Then
						Exit Function
					End If
					If KS.IsExpired("SoftArtisans.ImageGen") Then
						Response.Write ("对不起，SoftArtisans.ImageGen组件已过期！")
						Response.End
					End If
					Set objImage = KS.InitialObject("SoftArtisans.ImageGen")
					objImage.LoadImage FileName
					If Rate = 0 And (Width <> 0 Or Height <> 0) Then
						If Width < objImage.Width And Height < objImage.Height Then
							If Width = 0 And Height <> 0 Then
								objImage.CreateThumb , CLng(Height), 0, True
							ElseIf Width <> 0 And Height = 0 Then
								objImage.CreateThumb CLng(Width), objImage.Height / objImage.Width * Width, 0, False
							ElseIf Width <> 0 And Height <> 0 Then
								objImage.CreateThumb CLng(Width), CLng(Height), 0, False
							End If
						End If
					ElseIf Rate <> 0 Then
						objImage.CreateThumb CLng(objImage.Width * Rate), CLng(objImage.Height * Rate), 0, False
					End If
					objImage.SaveImage 0, objImage.ImageFormat, ThumbFileName
				Case 4
					If Not KS.IsObjInstalled("CreatePreviewImage.cGvbox") Then       
						Exit Function
					End If
					Set objImage = KS.InitialObject("CreatePreviewImage.cGvbox")
					objImage.SetImageFile = FileName                           
					If Rate = 0 And (Width <> 0 Or Height <> 0) Then
						objImage.SetPreviewImageSize = Width                  
					ElseIf Rate <> 0 Then
						objImage.SetPreviewImageSize = objImage.SetPreviewImageSize * Rate            
					End If
					objImage.SetSavePreviewImagePath = ThumbFileName            
					If objImage.DoImageProcess = False Then                    
						Exit Function
					End If
			End Select
			CreateThumb = True
		End Function
End Class
%> 

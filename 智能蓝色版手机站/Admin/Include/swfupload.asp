<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.EscapeCls.asp"-->
<!--#include file="../../KS_Cls/UploadFunction.asp"-->
<!--#include file="../../plus/Session.asp"-->
<%
Server.ScriptTimeout=9999999
Response.CharSet="utf-8"

Dim KSCls
If Request("From")="Common" Then
 Set KSCls = New UpFileSaveByCommon
Else
 Set KSCls = New UpFileSaveBySwfUpload
End If
KSCls.Kesion()
Set KSCls = Nothing

Class UpFileSaveBySwfUpload
        Private KS,FileTitles,Title
		Dim FilePath,MaxFileSize,AllowFileExtStr,BasicType,ChannelID,UpType,EditorID
		Dim FormName,Path,TempFileStr,FormPath,ThumbFileName,ThumbPathFileName,UpGet
		Dim UpFileObj,CurrNum,CreateThumbsFlag,FieldName,U_FileSize
		Dim DefaultThumb    '设定第几张为缩略图
		Dim ReturnValue
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Function CheckIsLogin(UserName,Pass)
		     If UserName="" Or Pass="" Then Check=false: Exit Function
		     Dim ChkRS:Set ChkRS =Conn.Execute("Select top 1 * From KS_Admin Where UserName='" & KS.R(UserName) & "'")
			 If ChkRS.EOF And ChkRS.BOF Then
			   CheckIsLogin=false
			 Else
			   If ChkRS("PassWord")=Pass Then CheckIsLogin=true Else CheckIsLogin=false
			 End If
		     ChkRS.Close:Set ChkRS = Nothing
		End Function
		
		Sub Kesion()
			Set UpFileObj = New UpFileClass
			on error resume next
			UpFileObj.GetData
			If ERR.Number<>0 Then err.clear:KS.Die "error:" & escape("上传失败，可能您的上传的文件太大!")
			
			Dim KSLoginCls:Set KSLoginCls = New LoginCheckCls1
			If KSLoginCls.Check=false Then
			   If CheckIsLogin(UpFileObj.Form("AdminName"),UpFileObj.Form("AdminPass")) =false Then
				KS.Die "error:" & escape("对不起，没有登录!")
			   Else
			    Response.Cookies(KS.SiteSn)("AdminName")=UpFileObj.Form("AdminName")
			   End If
			End If
		
		FormPath=KS.GetUpFilesDir
		FilePath=Server.MapPath(FormPath) & "\"
		FormPath=FormPath & "/"
		If KS.Setting(97)=1 Then FormPath=KS.Setting(2) & FormPath
		EditorID =UpFileObj.Form("EditorID")
		BasicType=KS.ChkClng(UpFileObj.Form("BasicType"))   
		ChannelID=KS.ChkClng(UpFileObj.Form("ChannelID")) 
		UpType=UpFileObj.Form("UpType")
		UpGet=UpFileObj.Form("UpGet")
		CurrNum=0
		CreateThumbsFlag=false
		DefaultThumb=KS.ChkClng(UpFileObj.Form("DefaultUrl"))
		If UpType="Field" Then
		   Dim RS:Set RS=Conn.Execute("Select top 1 FieldName,AllowFileExt,MaxFileSize From KS_Field Where FieldID=" & KS.ChkClng(UpFileObj.Form("FieldID")))
		   If Not RS.Eof Then
		    FieldName=RS(0):MaxFileSize=RS(2):AllowFileExtStr=RS(1)
		   Else
		    Response.End()
		   End IF
		   RS.Close:Set RS=Nothing
        Elseif UpType="File" Then   '上传附件
			MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '设定文件上传最大字节数
			AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,0)
		ElseIf UpType="Pic" Then
			If DefaultThumb=1 Then CreateThumbsFlag=true Else CreateThumbsFlag=false
			MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '设定文件上传最大字节数
			AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,1)
		Else
			Select Case BasicType
			  Case 1,3,4,7,9    '下载,影片等
					MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '设定文件上传最大字节数
					If BasicType=4 Then
					 AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,2)
					Else
					 AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,0)
					End If
			  Case 2,5     '图片中心
					CreateThumbsFlag=true
					MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '设定文件上传最大字节数
					AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,1)
			End Select
		End If
		ReturnValue = CheckUpFile("",false,UpFileObj,FormPath,FilePath,MaxFileSize,AllowFileExtStr,U_FileSize,TempFileStr,FileTitles,CurrNum,CreateThumbsFlag,DefaultThumb,ThumbPathFileName)
        If Not KS.IsNul(UpFileObj.Form("fileNames")) Then FileTitles=unescape(UpFileObj.Form("fileNames")) '防止中文乱码
		If UpFileObj.Form("NoReName")="1" Then  '不更名
		        Dim PhysicalPath,FsoObj:Set FsoObj = KS.InitialObject(KS.Setting(99))
		        PhysicalPath = Server.MapPath(replace(TempFileStr,"|",""))
				TempFileStr= mid(TempFileStr,1, InStrRev(TempFileStr, "/")) &  FileTitles
				If FsoObj.FileExists(PhysicalPath)=true Then
				 FsoObj.MoveFile PhysicalPath,server.MapPath(TempFileStr)
			    End If
		End If
		
		if ReturnValue <> "" then
		     ReturnValue=replace(ReturnValue,"\n","。")
		     If Instr(ReturnValue,"上传失败")<>0 Then
		     KS.Die "error:" & escape("上传失败" & Replace(Split(ReturnValue,"上传失败")(1),"'","\'"))
			 Else
		     KS.Die "error:" & escape(Replace(ReturnValue,"'","\'"))
			 End If
		else 
			 TempFileStr=replace(TempFileStr,"'","\'")
			 If UpType="Field" Then
			 	KS.Die replace(TempFileStr,"|","")
             Elseif UpType="File" Then   '上传附件
				  Call AddAnnexToDB(ChannelID,KS.C("AdminName"),TempFileStr,FileTitles,0,false,EditorID)
			 ElseIf UpType="Pic" Then
			      if UpGet="min" then
				   KS.Echo replace(TempFileStr,"|","")
				  ElseIf BasicType=1 Or BasicType=5 Or BasicType=3  Or BasicType=8 Then
				   if ThumbPathFileName="" then ThumbPathFileName=replace(TempFileStr,"|","")
			       KS.Die ThumbPathFileName  &"@"& replace(TempFileStr,"|","") 
				  Else
				   if DefaultThumb=1 then
				     KS.Echo ThumbPathFileName
				     Call KS.DeleteFile(replace(TempFileStr,"|",""))  '删除原图片
				   Else
				     KS.Echo replace(TempFileStr,"|","")
				   End If
                   KS.Die ""
				  End If
			 Else
				 Select Case BasicType
				      Case 3 KS.Die escape(replace(TempFileStr,"|","")) & "|" & U_FileSize
					  Case 2,5  '图片
						  KS.Die replace(TempFileStr,"|","") &  "@" & ThumbPathFileName & "@" & escape(FileTitles)
					 Case Else KS.Die escape(replace(TempFileStr,"|",""))
				 End Select
			End If
		  End iF
		Set UpFileObj=Nothing
 End Sub
End Class





'普通上传处理类 
Class UpFileSaveByCommon
        Private KS,FileTitles
		Dim FilePath,MaxFileSize,AllowFileExtStr,BasicType,ChannelID,UpType
		Dim FormName,Path,TempFileStr,FormPath,ThumbPathFileName
		Dim UpFileObj,CurrNum,CreateThumbsFlag,FieldName,U_FileSize
		Dim DefaultThumb    '设定第几张为缩略图
		Dim ReturnValue
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		
		Function IsSelfRefer() 
			Dim sHttp_Referer, sServer_Name 
			sHttp_Referer = CStr(Request.ServerVariables("HTTP_REFERER")) 
			sServer_Name = CStr(Request.ServerVariables("SERVER_NAME")) 
			If Mid(sHttp_Referer, 8, Len(sServer_Name)) = sServer_Name Then 
			IsSelfRefer = True 
			Else 
			IsSelfRefer = False 
			End If 
		End Function 
		Sub Kesion()
		Response.Write("<style type='text/css'>" & vbcrlf)
		Response.Write("<!--" & vbcrlf)
		Response.Write("body {background:#f0f0f0;" & vbcrlf)
		Response.Write("	margin-left: 0px;" & vbcrlf)
		Response.Write("	margin-top: 0px;" & vbcrlf)
		Response.Write("}" & vbcrlf)
		Response.Write("-->" & vbcrlf)
		Response.Write("</style>" & vbcrlf)
		If KS.IsNul(KS.C("AdminName")) Or KS.IsNul(KS.C("AdminPass")) Or KS.IsNul(KS.C("PowerList"))="" Or KS.IsNUL(KS.C("UserName")) Then
			Response.Write "<script>alert('没有登录!');history.back();</script>"
			Response.end
		End If
		
		 If Trim(Request.ServerVariables("HTTP_REFERER"))="" Then
			Response.Write "<script>alert('非法上传1！');history.back();</script>"
			Response.end
		 End If
		 if instr(lcase(Request.ServerVariables("HTTP_REFERER")),"ks.upfileform.asp")<=0 and instr(lcase(Request.ServerVariables("HTTP_REFERER")),"upfileform.asp")<=0 then
			Response.Write "<script>alert('非法上传！');history.back();</script>"
			Response.end
		 end if
		 if IsSelfRefer=false Then
			Response.Write "<script>alert('请不要非法上传！');history.back();</script>"
			Response.end
		 End If
		 
		Set UpFileObj = New UpFileClass
		UpFileObj.GetData
		FormPath=Replace(UpFileObj.Form("Path"),".","") 
		IF Instr(FormPath,KS.Setting(3))=0 Then	FormPath=KS.Setting(3) & FormPath
		Call KS.CreateListFolder(FormPath)       '生成上传文件存放目录
		FilePath=Server.MapPath(FormPath) & "\"
		If KS.Setting(97)=1 Then
		FormPath=KS.Setting(2) & FormPath
		End if
		
		BasicType=KS.ChkClng(UpFileObj.Form("BasicType"))        ' 2-- 图片中心上传 3--下载中心缩略图/文件 
		ChannelID=KS.ChkClng(UpFileObj.Form("ChannelID")) 
		UpType=UpFileObj.Form("UpType")
		
		CurrNum=0
		CreateThumbsFlag=false
		DefaultThumb=UpFileObj.Form("DefaultUrl")
		if DefaultThumb="" then DefaultThumb=0
		
		If UpType="Pic" Then
			If DefaultThumb=1 Then CreateThumbsFlag=true Else CreateThumbsFlag=false
			MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '设定文件上传最大字节数
			AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,1)
		Else  '默认上传参数
			MaxFileSize = KS.ReturnChannelAllowUpFilesSize(0)  '设定文件上传最大字节数
			AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(0,0)
		End If	
		ReturnValue = CheckUpFile("",false,UpFileObj,FormPath,FilePath,MaxFileSize,AllowFileExtStr,U_FileSize,TempFileStr,FileTitles,CurrNum,CreateThumbsFlag,DefaultThumb,ThumbPathFileName)
		if ReturnValue <> "" then
		     ReturnValue = Replace(ReturnValue,"'","\'")
		     KS.AlertHintScript ReturnValue
			 Response.End()
		else 
			    TempFileStr=replace(TempFileStr,"'","\'")
				If UpType="Pic" Then 
				   If BasicType=1 Or BasicType=8 Then
						  Response.Write("<script language=""JavaScript"">")
						 If KS.CheckFile(Replace(ThumbPathFileName,KS.Setting(2),""))=true Then        '检查是否存在缩略图
							  Response.Write("parent.document.myform.PhotoUrl.value='" & ThumbPathFileName & "';")
						 Else
							  Response.Write("parent.document.myform.PhotoUrl.value='" & replace(TempFileStr,"|","") & "';")
							  response.write "parent.OpenImgCutWindow(0,'" & KS.Setting(3) & "','" &replace(TempFileStr,"|","") & "');" &vbcrlf
						 End If
					     If Split(Split(KS.C_S(ChannelID,5),"@@@")(0),"|")(9)=1 Then
							 Response.Write ("parent.insertHTMLToEditor('<img src=" & replace(TempFileStr,"|","") &" />');")
						 End If
						  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>图片上传成功！</font>');")
						  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=../KS.UpFileForm.asp?ChannelID=" & ChannelID & "&UpType=" & UpType & "\'>');")
						  Response.Write("</script>")
				   ElseIf BasicType=3 Or BasicType=5 Then
				           Response.Write("<script language=""JavaScript"">")
						  if DefaultThumb=0 then
						   Response.Write("parent.document.getElementById('PhotoUrl').value='" & replace(TempFileStr,"|","") & "';")
						   Response.Write("parent.document.getElementById('BigPhoto').value='" & replace(TempFileStr,"|","") & "';")
						   response.write "parent.OpenImgCutWindow(0,'" & KS.Setting(3) & "','" &replace(TempFileStr,"|","") & "');" &vbcrlf
						  else
						   Response.Write("parent.document.getElementById('PhotoUrl').value='" & ThumbPathFileName & "';")
						   Response.Write("parent.document.getElementById('BigPhoto').value='" & replace(TempFileStr,"|","") & "';")
						  end if
						  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>图片上传成功！</font>');")
						  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=../KS.UpFileForm.asp?ChannelID=" & ChannelID & "&UpType=" & UpType & "\'>');")
						  Response.Write("</script>")
				   Else
				      SuccessDefaultPhoto
				   End IF
				Else
						 if ReturnValue <> "" then
						  Response.Write("<script language=""JavaScript"">"&vbcrlf)
						  Response.Write("alert('" & ReturnValue & "');"&vbcrlf)
						  Response.Write("dialogArguments.location.reload();"&vbcrlf)
						  Response.Write("close();"&vbcrlf)
						  Response.Write("</script>"&vbcrlf)
						 else
						  Response.Write("<script language=""JavaScript"">"&vbcrlf)
						  Response.Write("dialogArguments.location.reload();"&vbcrlf)
						  Response.Write("close();"&vbcrlf)
						  Response.Write("</script>"&vbcrlf)
						 end if
			 End If
		  End iF
		Set UpFileObj=Nothing
		End Sub
		
		'上传默认图成功
		Sub SuccessDefaultPhoto()
	      Response.Write("<script language=""JavaScript"">")
		    if DefaultThumb=0 then
				 Response.Write("parent.myform.PhotoUrl.value='" & replace(TempFileStr,"|","") & "';")
				 Response.write "parent.OpenImgCutWindow(1,'" & KS.Setting(3) & "','" &replace(TempFileStr,"|","") & "');" &vbcrlf
		    else
				 Response.Write("parent.myform.PhotoUrl.value='" & ThumbPathFileName & "';")
				 Call KS.DeleteFile(replace(TempFileStr,"|",""))  '删除原图片
			end if
		   Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>图片上传成功！</font>');")
		   Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=../KS.UpFileForm.asp?ChannelID=7&upType=" & UpType & "\'>');")
		  Response.Write "</script>"
		End Sub
			
End Class
%> 

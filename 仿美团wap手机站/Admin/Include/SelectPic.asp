<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Session.asp"-->
<%

response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Dim KSCls
Set KSCls = New SelectPic
KSCls.Kesion()
Set KSCls = Nothing

Class SelectPic
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		With KS
		If KS.C("AdminName") = "" Then
		 .echo ("<script>alert('对不起，权限不足!');window.close();</script>")
		 Exit Sub
		End If
		Dim ChannelID, CurrPath, ShowVirtualPath
		Dim InstallDir
		Dim LimitUpFileFlag  '上传权限 值yes无上传权限
			InstallDir = KS.Setting(3)
			ChannelID = KS.G("ChannelID")
			CurrPath = KS.G("CurrPath")
			ShowVirtualPath = KS.G("ShowVirtualPath")
			If ChannelID = "" Or Not IsNumeric(ChannelID) Then ChannelID = 0
				 If KS.ReturnChannelAllowUpFilesTF(ChannelID) = False Then
				  LimitUpFileFlag = "yes"
				 End If
				 If KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10009") = False Then
				  LimitUpFileFlag = "yes"
				 End If
				 
			if instr(request("currpath"),".")<>0 then
			  ks.die "非法参数!"
			end if
           If InstallDir<>"/" then 
			if instr(CurrPath,InstallDir)=0 Then
			CurrPath = Replace(InstallDir & CurrPath,"//","/")
			End If
		  End iF
		  if instr(lcase(currpath),"uploadfiles/")=0 then currpath=KS.GetUpFilesDir
		  If KS.C("SuperTF")="1" Then CurrPath=KS.Setting(3) & left(ks.setting(91),len(ks.setting(91))-1)
		  
		  if currpath="/" then currpath =ks.setting(3) & left(ks.setting(91),len(ks.setting(91))-1)
		
		.echo "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">"
		.echo "<html>"
		.echo "<head>"
		.echo "<META HTTP-EQUIV=""pragma"" CONTENT=""no-cache"">" 
        .echo "<META HTTP-EQUIV=""Cache-Control"" CONTENT=""no-cache, must-revalidate"">"
        .echo "<META HTTP-EQUIV=""expires"" CONTENT=""Wed, 26 Feb 1997 08:21:57 GMT"">"
		.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		
		

		
		.echo "<title>选择文件</title>"
		.echo "<style type=""text/css"">" & vbCrLf
		.echo "<!--" & vbCrLf
		.echo ".PreviewStyle {" & vbCrLf
		.echo "    border: 2px outset #CCCCCC;"
		.echo "}"
		.echo ".ImgOver {"
		.echo "    cursor: default;"
		.echo "    border-top-width: 1px;"
		.echo "    border-right-width: 1px;"
		.echo "    border-bottom-width: 1px;"
		.echo "    border-left-width: 1px;"
		.echo "    border-top-style: solid;"
		.echo "    border-right-style: solid;"
		.echo "    border-bottom-style: solid;"
		.echo "    border-left-style: solid;"
		.echo "    border-top-color: #FFFFFF;"
		.echo "    border-right-color: #999999;"
		.echo "    border-bottom-color: #999999;"
		.echo "    border-left-color: #FFFFFF;"
		.echo "}"
		.echo " BODY   {border: 0; margin: 0; background: buttonface; cursor: default; font-family:宋体; font-size:9pt;}"
		.echo " BUTTON {width:5em}" & vbCrLf
		.echo " TABLE  {font-family:宋体; font-size:9pt}"
		.echo " P      {text-align:center}" & vbCrLf
		.echo "-->" & vbCrLf
		.echo "</style>"
		.echo "</head>"
		.echo "<script language=""JavaScript"" src=""../../KS_inc/Common.js""></script>"
		.echo "<body leftmargin=""0"">"
		.echo "<table width=""99%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">"
		.echo "  <tr>"
		 .echo "   <td colspan=""2""><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
		 .echo "       <tr>"
		 .echo "         <td width=""80"" align=""center"" nowrap>选择目录： </td>"
		 .echo "         <td width=""649""><select onChange=""ChangeFolder(this.value);"" id=""FolderSelectList"" style=""width:100%;"" name=""select"">"
		 .echo "             <option selected value=""" & CurrPath & """>"
		 .echo CurrPath
		 .echo "             </option>"
		 .echo "           </select> </td>"
		.echo "          <td width=""279"" height=""26"" valign=""middle"">"
		 .echo "           <table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		 .echo "             <tr align=""center"">"
		 .echo "               <td width=""25"">&nbsp;</td>"
		 .echo "                <td width=""25"" onMouseOver=""this.className='ImgOver'"" onMouseOut=""this.className=''""><img src=""../Images/Folder/R.gif"" align=""absmiddle"" onClick=""ChangeViewArea(this);"" id=""Img1"" title=""关闭预览区""></td>"
		 .echo "               <td width=""25"" onMouseOver=""this.className='ImgOver'"" onMouseOut=""this.className=''""><img src=""../Images/Folder/B.gif"" width=""21"" height=""22"" align=""absmiddle"" onClick=""frames['FolderList'].OpenParentFolder();"" title=""返回上一级目录""></td>"
		 .echo "               <td width=""25"" onMouseOver=""this.className='ImgOver'"" onMouseOut=""this.className=''""><img src=""../Images/Folder/AddFolder.gif"" width=""19"" height=""17"" align=""absmiddle"" onClick=""frames['FolderList'].AddFolderOperation();"" title=""添加新目录""></td>"
		.echo "                <td width=""25"" onMouseOver=""this.className='ImgOver'"" onMouseOut=""this.className=''""><img src=""../Images/Folder/Upfiles.gif"" width=""21"" height=""22"" align=""absmiddle"""
						 If LimitUpFileFlag <> "yes" Then
							 .echo ("onClick=UpFile();")
						   Else
							 .echo ("onclick='alert(""系统设定此模块不允许上传文件或你没有上传文件的权限"");'")
						   End If
		.echo "                  title=""上传新文件""></td>"
		.echo "                <td width=""25"">&nbsp;</td>"
		.echo "                <td width=""25"">&nbsp;</td>"
		.echo "                <td width=""25"">&nbsp;</td>"
		.echo "                <td width=""25"">&nbsp;</td>"
		.echo "                <td width=""25"">&nbsp;</td>"
		.echo "              </tr>"
		.echo "            </table>"
		.echo "          </td>"
		.echo "        </tr>"
		.echo "      </table>"
		.echo "    </td>"
		.echo "  </tr>"
		.echo "  <tr>"
		.echo "    <td width=""70%"" align=""center""> <iframe name=""FolderList"" id=""FolderList"" width=""100%"" height=""290"" frameborder=""1"" src=""FolderFileList.asp?ChannelID=" & ChannelID & "&CurrPath=" & CurrPath & "&ShowVirtualPath=" & ShowVirtualPath & """ scrolling=""yes""></iframe>"
		.echo "    </td>"
		.echo "    <td width=""30%""  align=""center"" valign=""middle"" ID=""ViewArea""> <iframe name=""PreviewArea"" id=""PreviewArea"" scrolling=""yes"" width=""100%"" height=""290"" frameborder=""1"" src=""Preview.asp""></iframe>"
		.echo "    </td>"
		.echo "  </tr>"
		.echo "  <tr>"
		.echo "    <td height=""35"" colspan=""2""> <table width=""100%"" height=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		.echo "        <tr>"
		.echo "          <td width=""80"" height=""40""> <div align=""center"">URL地址：</div></td>"
		.echo "          <td><input style=""width:65%"" type=""text"" name=""FileUrl"" id=""FileUrl""> <input type=""button"" onClick=""SetFileUrl();"" name=""Submit"" value="" 确 定 "">"
		.echo "            <input onClick=""window.close();"" type=""button"" name=""Submit3"" value="" 取 消 "">"
		.echo "          </td>"
		.echo "        </tr>"
		.echo "      </table></td>"
		.echo "  </tr>"
		.echo "</table>"
		.echo "</body>"
		.echo "</html>"
		.echo "<script language=""JavaScript"">"
		.echo "var ChannelID=" & ChannelID & ";"
		.echo "function ChangeFolder(FolderName)"
		.echo "{"
		.echo "    frames[""FolderList""].location='FolderFileList.asp?CurrPath='+FolderName;"
		.echo "}"
		.echo "function UpFile()"
		.echo "{"
		.echo "  OpenWindow('Frame.asp?ChannelID='+ChannelID+'&PageTitle='+escape('上传文件')+'&FileName=UpFileForm.asp&Path='+frames[""FolderList""].CurrPath,400,200,window);"
		.echo "    frames[""FolderList""].location='FolderFileList.asp?CurrPath='+frames[""FolderList""].CurrPath;"
		.echo "}"
		.echo "function SetFileUrl()"
		.echo "{"
		.echo "    if (document.getElementById('FileUrl').value=='') alert('请填写Url地址');"
		.echo "    else"
		if (request("from")="ckeditor") then
		.echo "{"
		.echo "window.opener.CKEDITOR.tools.callFunction('" & request("CKEditorFuncNum") &"',document.getElementById('FileUrl').value);"
		.echo " top.close();"
		.echo "}"
		else
		.echo "    {"
		.echo "       if (document.all){ window.returnValue=document.getElementById('FileUrl').value;}else{window.opener.setVal(document.getElementById('FileUrl').value)}"
		.echo "        top.close();"
		.echo "    }"
		end if
		.echo "}"
		.echo "window.onunload=CheckReturnValue;"
		.echo "function CheckReturnValue()"
		.echo "{"
		.echo "    if (typeof(window.returnValue)!='string') window.returnValue='';"
		.echo "}"
		.echo "var displayBar=true;"
		.echo "function ChangeViewArea(obj) {"
		.echo "  if (displayBar) {"
		.echo "  ViewArea.style.display='none';"
		.echo "    displayBar=false;"
		.echo "    obj.src='../Images/Folder/L.gif';"
		.echo "    obj.title='打开预览区';"
		.echo "  } else {"
		.echo "   ViewArea.style.display='';"
		.echo "    displayBar=true;"
		.echo "    obj.src='../Images/Folder/R.gif';"
		.echo "    obj.title='关闭预览区';"
		.echo "  }"
		.echo "}"
		
		.echo "</script>"
		End With
		End Sub
End Class
%> 

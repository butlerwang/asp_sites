<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Session.asp"-->
<%

Dim KSCls
Set KSCls = New Preview
KSCls.Kesion()
Set KSCls = Nothing

Class Preview
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		 On Error Resume Next
		Dim PreviewImagePath, FileExtName, FileIconDic, FileIcon, AvaiLabelShowTypeStr, PicPara
		PreviewImagePath = KS.G("FilePath")
		AvaiLabelShowTypeStr = "jpg,gif,bmp,pst,png,ico"
		Set FileIconDic = CreateObject("Scripting.Dictionary")
		FileIconDic.Add "txt", "../../editor/ksplus/FileIcon/txt.gif"
		FileIconDic.Add "gif", "../../editor/ksplus/FileIcon/gif.gif"
		FileIconDic.Add "exe", "../../editor/ksplus/FileIcon/exe.gif"
		FileIconDic.Add "asp", "../../editor/ksplus/FileIcon/asp.gif"
		FileIconDic.Add "html", "../../editor/ksplus/FileIcon/html.gif"
		FileIconDic.Add "htm", "../../editor/ksplus/FileIcon/html.gif"
		FileIconDic.Add "jpg", "../../editor/ksplus/FileIcon/jpg.gif"
		FileIconDic.Add "jpeg", "../../editor/ksplus/FileIcon/jpg.gif"
		FileIconDic.Add "pl", "../../editor/ksplus/FileIcon/perl.gif"
		FileIconDic.Add "perl", "../../editor/ksplus/FileIcon/perl.gif"
		FileIconDic.Add "zip", "../../editor/ksplus/FileIcon/zip.gif"
		FileIconDic.Add "rar", "../../editor/ksplus/FileIcon/zip.gif"
		FileIconDic.Add "gz", "../../editor/ksplus/FileIcon/zip.gif"
		FileIconDic.Add "doc", "../../editor/ksplus/FileIcon/doc.gif"
		FileIconDic.Add "xml", "../../editor/ksplus/FileIcon/xml.gif"
		FileIconDic.Add "xsl", "../../editor/ksplus/FileIcon/xml.gif"
		FileIconDic.Add "dtd", "../../editor/ksplus/FileIcon/xml.gif"
		FileIconDic.Add "vbs", "../../editor/ksplus/FileIcon/vbs.gif"
		FileIconDic.Add "js", "../../editor/ksplus/FileIcon/vbs.gif"
		FileIconDic.Add "wsh", "../../editor/ksplus/FileIcon/vbs.gif"
		FileIconDic.Add "sql", "../../editor/ksplus/FileIcon/script.gif"
		FileIconDic.Add "bat", "../../editor/ksplus/FileIcon/script.gif"
		FileIconDic.Add "tcl", "../../editor/ksplus/FileIcon/script.gif"
		FileIconDic.Add "eml", "../../editor/ksplus/FileIcon/mail.gif"
		FileIconDic.Add "swf", "../../editor/ksplus/FileIcon/flash.gif"
		If PreviewImagePath = "" Then
			PreviewImagePath = "../../editor/ksplus/FileIcon/DefaultPreview.gif"
		Else
			FileExtName = Right(PreviewImagePath, Len(PreviewImagePath) - InStrRev(PreviewImagePath, "."))
			If InStr(AvaiLabelShowTypeStr, lcase(FileExtName)) = 0 Then
				FileIcon = FileIconDic.Item(LCase(FileExtName))
				If FileIcon = "" Then
					FileIcon = "../../editor/ksplus/FileIcon/unknown.gif"
				End If
				PreviewImagePath = FileIcon
				PicPara = " width=""30"" height=""30"" "
			Else
				PicPara = ""
			End If
		End If
		Set FileIconDic = Nothing
		
		Response.Write "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">"
		Response.Write "<html>"
		Response.Write "<head>"
		Response.Write "<link href=""Admin_Style.CSS"" rel=""stylesheet"">"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		Response.Write "<title>预览</title>"
		Response.Write "</head>"
		Response.Write "<body topmargin=""0"" leftmargin=""0"">"
		Response.Write "<table width=""100%"" height=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		Response.Write "  <tr>"
		Response.Write "    <td align=""center"" valign=""middle""><img  " & PicPara & " src=""" & PreviewImagePath & """></td>"
		Response.Write "  </tr>"
		Response.Write "</table>"
		Response.Write "</body>"
		Response.Write "</html>"
		
		End Sub
End Class
%> 

<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="Session.asp"-->
<%

Dim KSCls
Set KSCls = New RefreshCommonPageSave
KSCls.Kesion()
Set KSCls = Nothing

Class RefreshCommonPageSave
        Private KS
		Private KSRObj
		Private ReturnInfo
		Private Sub Class_Initialize()
		  Set KSRObj=New Refresh
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KSRObj=Nothing
		 Set KS=Nothing
		End Sub
		Sub Kesion()
		 	'On Error Resume Next
		Dim RefreshFlag, RefreshSql, FolderID, NewsNo, RefreshRS, NewsTotalNum, StartRefreshTime
		'刷新时间
		StartRefreshTime = Request("StartRefreshTime")
		If StartRefreshTime = "" Then StartRefreshTime = Timer()
		
		RefreshFlag = Request.QueryString("RefreshFlag")
		  RefreshSql = Trim(Request("RefreshSql"))
		  NewsNo = Request("NewsNo")
		 If NewsNo = "" Then NewsNo = 0
		 If RefreshSql = "" Then
			If RefreshFlag = "Folder" Then
			  FolderID = Request("PageID")
			  If Right(FolderID,1)="," then  FolderID=Left(Folderid,Len(FolderID)-1)
			  RefreshSql = "Select * From KS_Template Where TemplateID IN (" & FolderID & ")"
			ElseIf RefreshFlag = "All" Then
			  RefreshSql = "Select * From KS_Template"
		   Else
			  RefreshSql = ""
		   End If
		End If
		If RefreshSql <> "" Then
			Set RefreshRS = Server.CreateObject("ADODB.RecordSet")
			RefreshRS.Open RefreshSql, Conn, 1, 1
			NewsTotalNum = RefreshRS.RecordCount
			If RefreshRS.EOF Then
				ReturnInfo = "没有要刷新的通用页面&nbsp;&nbsp;<br><input name=""button1"" type=""button"" onclick=""javascript:location='RefreshCommonPage.asp';""  class='button' value="" 返 回 "">"
				Set RefreshRS = Nothing
				Call Main
			Else
				RefreshRS.Move NewsNo
				If Not RefreshRS.EOF Then
				   Call KSRObj.RefreshCommonPage(refreshrs("TemplateFileName"),RefreshRS("fsofilename"))  '调用通用页面刷新函数
					NewsNo = NewsNo + 1
					Response.Write ("<meta http-equiv=""refresh"" content=""0;url='RefreshCommonPageSave.asp?StartRefreshTime=" & Server.URLEncode(StartRefreshTime) & "&NewsNo=" & NewsNo & "&RefreshSql=" & Server.URLEncode(RefreshSql) & "&RefreshFlag=" & RefreshFlag & "'"">")
					ReturnInfo = "总共需要刷新 <font color=red><b>" & NewsTotalNum & "</b></font> 个通用页面<br><br>正在刷新第 <font color=red><b>" & NewsNo - 1 & "</b></font> 个通用页面,请稍候... <font color=red><b>在此过程中请勿刷新此页面！！！</b></font><br>"
				Else
					ReturnInfo = "刷新通用页面结束！总费时:<font color=red>" & Left((Timer() - StartRefreshTime), 4) & "</font> 秒<br><br>总共刷新了 <font color=red><b>" & NewsTotalNum & "</b></font> 个通用页面 <br><br><input name=""button1"" type=""button"" onclick=""javascript:location='RefreshCommonPage.asp';""  class='button' value="" 返 回 "">"
				End If
				Set RefreshRS = Nothing
				Call Main
			End If
			Set RefreshRS = Nothing
		Else
			ReturnInfo = "对不起，您没有选择要发布的通用页面&nbsp;&nbsp;<font color=""red""><a href=""RefreshCommonPage.asp"">返回</a></font>"
			Call Main
		End If
		
		End Sub
		
		Sub Main()
		 Response.Write ("<html>")
		 Response.Write ("<head>")
		 Response.Write ("<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">")
		 Response.Write ("<title>系统信息</title>")
		 Response.Write ("</head>")
		 Response.Write ("<link rel=""stylesheet"" href=""Admin_Style.css"">")
		 Response.Write ("<body oncontextmenu=""return false;"">")
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
End Class
%> 

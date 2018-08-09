<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="Session.asp"-->
<%

Dim KSCls
Set KSCls = New RefreshIndex
KSCls.Kesion()
Set KSCls = Nothing

Class RefreshIndex
        Private KS,KSRObj
		Private SaveFilePath
		Private FileContent
        Private ReturnInfo
		Private ErrFlag
		Private Domain
		Private IndexFile
		Private StartRefreshTime
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSRObj=New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KSRObj=Nothing
		 Set KS=Nothing
		End Sub
		Sub Kesion()
		  With Response
		  If Not KS.ReturnPowerResult(0, "KMTL20000") Then          '检查是否有发布站点首页的权限
		   .Write ("<script>$(parent.document).find('#BottomFrame')[0].src='javascript:history.back();';</script>")
		   Call KS.ReturnErr(1, "")
		   Exit Sub
		   .End
		  End If
		 If Split(KS.Setting(5),".")(1)="asp" Then 
		   Head 
		   KS.Die ("<script>$.dialog.tips('KesionCMS系统提醒您：<br/>1、站点首页没有启用生成静态HTML功能<br/>2、请到系统设置->基本信息设置启用生成静态Html功能<br/><br/><input type=""button"" style=""background:url(../images/button_bg.gif) left repeat-x;cursor: pointer;height:23px;BORDER-RIGHT: #2C59AA 1px solid; PADDING-RIGHT: 2px; BORDER-TOP: #2C59AA 1px solid; PADDING-LEFT: 2px; FONT-SIZE: 12px; FILTER: progid:DXImageTransform.Microsoft.Gradient(GradientType=0, StartColorStr=#ffffff, EndColorStr=#C3DAF5); BORDER-LEFT: #2C59AA 1px solid;  color:#555; PADDING-TOP: 2px; padding-top:0px !important;BORDER-BOTTOM: #2C59AA 1px solid"" value="" 返 回 "" onclick=""history.back();""/>',5000,'error.gif',function(){history.back()});</script>"):Exit Sub
		 End If
		   StartRefreshTime = Timer()
		   FCls.RefreshType = "INDEX" '设置刷新类型，以便取得当前位置导航等
		   FCls.RefreshFolderID = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
		   Domain = KS.GetDomain
		   IndexFile = KS.Setting(5)
			SaveFilePath = KS.Setting(3) & IndexFile
			FileContent = KSRObj.LoadTemplate(KS.Setting(110))
			If FileContent = "" Then
			  ReturnInfo = "读不到首页模板"
			  ErrFlag = True
			  Call Main
			  .End
			Else
			 ' On Error Resume Next
			
			  FileContent = KSRObj.KSLabelReplaceAll(FileContent) '替换函数标签
			  If Err Then
			   ReturnInfo = Err.Description
			   ErrFlag = True
				 Err.Clear
				Call Main
				.End
			  End If
			  Call KSRObj.FSOSaveFile(FileContent, SaveFilePath)
			  If Err Then
				ReturnInfo = Err.Description
				ErrFlag = True
				 Err.Clear
				Call Main
				.End
			  End If
			  ReturnInfo = "首页更新成功"
			  ErrFlag = False
			  Call Main
			  if request("f")="task" then
			   KS.Echo "<script>setTimeout('window.close();',3000);</script>"
			  end if
			End If
		 End With
		End Sub
		Sub Head()
		 With Response	  
		        .Write 	"<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd""><html xmlns=""http://www.w3.org/1999/xhtml"">"
				.Write ("<head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'>")
				.Write ("<title>网站首页发布管理</title></head>")
				.Write ("<link rel='stylesheet' href='Admin_Style.CSS'>")
		        .Write "<script language='JavaScript' src='../../KS_Inc/jquery.js'></script>"
		        .Write "<script language='JavaScript' src='../../KS_Inc/lhgdialog.js'></script>"
         End With
		End Sub
		Sub Main()
			With Response	  
		        Head 
				If KS.G("RefreshFlag")<>"Info" Then
				.Write ("<body topmargin='0' leftmargin='0' oncontextmenu='return false;'>")
				Else
		        .Write ("<body oncontextmenu=""return false;"" scroll=no bgcolor='transparent'>")
				End If
				If KS.G("RefreshFlag")<>"Info" Then
				.Write ("<table width='100%' border='0' cellpadding='0' cellspacing='0'>")
				.Write ("  <tr>")
				.Write ("    <td height='25' class='sort'>")
				.Write ("      <div align='center'><strong>发布网站首页</strong></div></td>")
				.Write ("</tr>")
				.Write ("</table>")
				.Write ("<table width='100%' border='0' cellspacing='0' cellpadding='0'><tr><td>&nbsp;</td></tr>")
				.Write ("  <tr><td>&nbsp;</td></tr>")
				.Write ("  <tr>")
				.Write ("    <td height='50'><div align='center'><br>")
				Dim LeftTime:LeftTime=Left((Timer() - StartRefreshTime), 4)
				If LeftTime>0 And LeftTime<1 Then LeftTime="0" & LeftTime
				.Write ReturnInfo & "！总费时：<font color=red>" &LeftTime & "</font> 秒,文件大小：<font color=green>" & Round(KS.GetFieSize(Server.MapPath(SaveFilePath))/1024,2) & "</font> KB</div></td>"
				.Write ("</tr><tr><td><div align='center'>")
				.Write ("        <table width='100%' border='0' cellspacing='0' cellpadding='0'>")
				.Write ("          <tr><td width='50%' height='25'><div align='right'>操作结果:</div></td>")
				.Write ("            <td style='text-align:left' height='25'>")
						   
						   If ErrFlag = False Then
							.Write ("成功")
							Else
							.Write ("失败")
							End If
				.Write ("            </td></tr><tr><td height='25'> <div align='right'>当前时间:</div></td><td height='25' style='text-align:left'> " & Now & "</td></tr>")
						 
						 If ErrFlag = False Then
				.Write ("          <tr><td height='25'><div align='right'>点击浏览:</div></td>")
				.Write ("           <td height='25' style='text-align:left'><font color='#FF0000'> <a href='" & Domain & IndexFile & "' target='_blank'>浏览首页</a></font>")
				.Write ("            </td></tr>")
						  End If
				.Write ("        </table></div></td></tr></table>")
			 Else
				.Write ("<table width='67%' border='0' cellpadding='0' cellspacing='0'>")
				.Write ("  <tr>")
				.Write ("    <td height='25'>")
				.Write ("      <div ><li><strong>" & ReturnInfo & "</strong><font color='#FF0000'> <a href='" & Domain & IndexFile & "' target='_blank'>" & Domain & IndexFile &"</a></font></li></div></td>")
				.Write ("</tr>")
				.Write ("</table>")
			 End If
			 .Write    ("</body></html>")
		End With
		End Sub
End Class
%> 

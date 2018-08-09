<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.SpaceCls.asp" -->
<%

Dim KSCls
Set KSCls = New RSSCls
KSCls.Kesion()
Set KSCls = Nothing

Class RSSCls
        Private KS,KSBcls
		Private sRssBody,UserName
		Private sTitle, sDeScriptIon, sLogo
		Private ChannelID,RssBody
		Private RssTF,RssCode,RssTemplateTF,RssHomeNum,RssChannelNum,RssDescriptNum,CodeChar,CodeNum

		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSBcls=New BlogCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSBCls=Nothing
		End Sub
       Sub Kesion()
	    With Response		
		RSSTF          = KS.Setting(83)
		RssCode        = KS.Setting(84)
		RssTemplateTF  = KS.Setting(85)
		RssHomeNum     = KS.Setting(86)
		RssChannelNum  = KS.Setting(87)
		RssDescriptNum = KS.Setting(88)
		If Cint(RssCode)=1 Then
			CodeChar="UTF-8"
			CodeNum=65001
		Else
		  CodeChar="gb2312"
		  CodeNum=936
		End If
		UserName    = KS.R(KS.S("UserName"))
		WebUrl	    = KS.GetDomain
		sTitle		= KSBcls.GetUserBlogParam(UserName,"BlogName")
		sDeScriptIon= KSBcls.GetUserBlogParam(UserName,"Descript")
		sLogo		= Replace(KS.Setting(4),"{$GetInstallDir}",WebUrl)
	
		If RssTF=0 Then .Write "<br/><div align=center>对不起。本站点没有提供RSS订阅功能，请与网站管理员联系!</div>":.End
	  	.Expires=0
		.CodePage=CodeNum
		.ContentType="text/xml"
		.Charset=CodeChar
		RssBody     =GetRssBody
		.Write GetShowRssBody(RssTemplateTF)
	End With
End Sub

Function GetShowRssBody(RssTemplateTF)
	GetShowRssBody	=GetShowRssBody & "<?xml version=""1.0"" encoding=""" & CodeChar & """?>"&vbcrlf
	If RssTemplateTF=1 Then
	GetShowRssBody	=GetShowRssBody & "<?xml-stylesheet type=""text/xsl"" href=""rss.xsl"" version=""1.0""?>"&vbcrlf
	End If
	GetShowRssBody	=GetShowRssBody & "<rss version=""2.0"">"&vbcrlf
	GetShowRssBody	=GetShowRssBody & "<channel>"&vbcrlf
	GetShowRssBody	=GetShowRssBody & "<title>" & sTitle & "</title>"&vbcrlf
	GetShowRssBody	=GetShowRssBody & "<description>" & sDeScriptIon & "</description> "&vbcrlf
	GetShowRssBody	=GetShowRssBody & "<link>" & WebUrl & "</link>"&vbcrlf	GetShowRssBody	=GetShowRssBody & "<generator>Rss Generator By</generator>"&vbcrlf
	GetShowRssBody	=GetShowRssBody & "<language>zh-cn</language>"&vbcrlf
	GetShowRssBody	=GetShowRssBody & "<copyright>Copyright 2006-2010.All Rights Reserved</copyright>"&vbcrlf
	GetShowRssBody	=GetShowRssBody & "<webMaster>" & KS.Setting(10)  & "</webMaster>"&vbcrlf
	GetShowRssBody	=GetShowRssBody & "<email>" & KS.Setting(11) & "</email>"&vbcrlf
	GetShowRssBody	=GetShowRssBody & "<image>"&vbcrlf
	GetShowRssBody	=GetShowRssBody & "	<title>" & sTitle & "</title> "&vbcrlf
	GetShowRssBody	=GetShowRssBody & "	<url>" & sLogo & "</url> "&vbcrlf
	GetShowRssBody	=GetShowRssBody & "	<link>" & WebUrl & "</link> "&vbcrlf
	GetShowRssBody	=GetShowRssBody & "	<description>" & sDeScriptIon & "</description> "&vbcrlf
	GetShowRssBody	=GetShowRssBody & "</image>"&vbcrlf
	GetShowRssBody	=GetShowRssBody & RssBody
	GetShowRssBody	=GetShowRssBody & "</channel>"&vbcrlf
	GetShowRssBody	=GetShowRssBody & "</rss>"&vbcrlf
End Function

Function GetRssBody()
		  sTitle = sTitle & "-最新日志"
		 Dim SqlStr,SQL,Rs,i
		   SqlStr="Select Top " &RssChannelNum & " ID,Title,Content,AddDate,userid From KS_BlogInfo Where UserName='" & UserName & "'  and status=0 Order By ID Desc"
		Set Rs=Conn.Execute(SqlStr)
		if Rs.Bof and Rs.Eof then
			GetRssBody = GetRssBody & "<item></item>"
			Rs.Close : Set Rs = Nothing
		Else
			Do While Not RS.Eof 
				GetRssBody = GetRssBody & "<item>"&vbcrlf
				GetRssBody = GetRssBody & "<title> " & RS(1) & "</title>"&vbcrlf
				GetRssBody = GetRssBody & "<link><![CDATA[" & KSBCls.GetCurrLogUrl(RS("UserID"),RS(0)) & "]]></link>"&vbcrlf
				If RssDescriptNum<>0 Then
				GetRssBody = GetRssBody & "<description><![CDATA[" & KS.GotTopic(Replace(Replace(Replace(KS.LoseHtml(RS(2)), vbCrLf, ""), "[NextPage]", ""), "&nbsp;", ""),RssDescriptNum) & ".....]]>.</description>"&vbcrlf
				End IF
				GetRssBody = GetRssBody & "<author>" & UserName & "</author>"&vbcrlf
				GetRssBody = GetRssBody & "<pubDate>" & RS(3) & "</pubDate>"&vbcrlf
				GetRssBody = GetRssBody & "</item>"&vbcrlf
              RS.MoveNext
		   Loop
		   Rs.Close : Set Rs = Nothing
		End if
End Function
End Class
%> 

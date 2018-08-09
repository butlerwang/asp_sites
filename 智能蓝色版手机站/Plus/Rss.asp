<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp" -->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%

Dim KSCls
Set KSCls = New RSSCls
KSCls.Kesion()
Set KSCls = Nothing

Class RSSCls
        Private KS,KSR
		Private sRssBody,maps
		Private sTitle, sDeScriptIon, sLogo
		Private ChannelID, sClassID,sElite,sHot,RssBody
		Private RssTF,RssCode,RssTemplateTF,RssHomeNum,RssChannelNum,RssDescriptNum,CodeChar,CodeNum

		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSR = New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSR=Nothing
		End Sub
       Sub Kesion()
	    With Response
		ChannelID	= KS.ChkClng(KS.S("ChannelID"))
		if channelid=0 then call showrss():exit sub
		
		RSSTF          = KS.ChkClng(KS.Setting(83))
		RssCode        = KS.ChkClng(KS.Setting(84))
		RssTemplateTF  = KS.ChkClng(KS.Setting(85))
		RssHomeNum     = KS.ChKclng(KS.Setting(86))
		RssChannelNum  = KS.ChkClng(KS.Setting(87))
		RssDescriptNum = KS.ChkClng(KS.Setting(88))
		'response.write RssTemplateTF 
		'response.end

		If KS.ChkClng(RssCode)=1 Then
			CodeChar="UTF-8"
			CodeNum=65001
		Else
		  CodeChar="gb2312"
		  CodeNum=936
		End If
		WebUrl	    = KS.GetDomain
		'sClassID	= KS.ChkClng(KS.S("ClassID"))
		sClassID	= KS.S("ClassID")
		sElite      = KS.ChkClng(KS.S("Elite"))
		sHot        = KS.ChkClng(KS.S("Hot"))
		sTitle		= KS.Setting(1)
		sDeScriptIon= KS.Setting(1)
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

Sub ShowRSS()
	 Dim FileContent
	 Dim RssTemplatePath:RssTemplatePath=KS.Setting(3) & KS.Setting(90) & "Common/rss.html"  '模板地址
	 FileContent = KSR.LoadTemplate(RssTemplatePath)    
	 FCls.RefreshType = "rss" '设置刷新类型，以便取得当前位置导航等
	 FCls.RefreshFolderID = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
	 Call RssList()
	 FileContent=Replace(FileContent,"{$ShowRss}",maps)
	 FileContent=KSR.KSLabelReplaceAll(FileContent)
	 response.write FileContent
End Sub
		
Sub RssList()
				Dim RS,FolderName,TreeStr,ID,SqlStr,Tj,SpaceStr,K
				Set  RS=Server.CreateObject("ADODB.Recordset")
				SQLstr = "select a.ID,a.FolderName,a.FolderOrder,a.ClassType,a.ChannelID,a.tj,a.tn,a.adminpurview from KS_Class a inner join ks_channel b on a.channelid=b.channelid where b.channelstatus=1 Order BY root,folderorder"
				RS.Open SQLstr, Conn, 1, 1
				If Not RS.Eof Then Set ClassXml=KS.RsToXml(RS,"row","")
				RS.Close
				Set RS=Nothing
				If IsOBject(ClassXml) Then
				  For Each Node In ClassXML.DocumentElement.SelectNodes("row")
				      TJ=Node.SelectSingleNode("@tj").text
					  If tJ=1 Then
				        TreeStr = TreeStr  & "<li class='classname'>" & KS.GetClassNP(Node.SelectSingleNode("@id").text)& "<span class=""r""><a href=""rss.asp?classid=" & Node.SelectSingleNode("@id").text & "&channelid=" & Node.SelectSingleNode("@channelid").text & """ target=""_blank""><img src=""images/rss_xml.gif"" align=""absmiddle"" border=""0""></a></span></li><br />"
					  Else
						SpaceStr=""
						For k = 1 To TJ - 1
						  SpaceStr = SpaceStr & ""
						Next
	                    TreeStr = TreeStr & "<li class=""rss_list"">" & SpaceStr & KS.GetClassNP(Node.SelectSingleNode("@id").text) & "<span class=""r""><a href=""rss.asp?classid=" & Node.SelectSingleNode("@id").text & "&channelid=" & Node.SelectSingleNode("@channelid").text & """ target=""_blank""><img src=""images/rss_xml.gif""  align=""absmiddle""  border=""0""></a></span></li>" & vbcrlf
					  End If
				  Next
				End If
				
			 Maps=TreeStr
End Sub


Function GetShowRssBody(RssTemplateTF)
	GetShowRssBody	=GetShowRssBody & "<?xml version=""1.0"" encoding=""" & CodeChar & """?>"&vbcrlf
	If KS.ChkClng(RssTemplateTF)=1 Then
	GetShowRssBody	=GetShowRssBody & "<?xml-stylesheet type=""text/xsl"" href=""images/rss.xsl"" version=""1.0""?>"&vbcrlf
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
	IF ChannelID<>0 Then
	    IF sElite<>"0" Then
		  sTitle = sTitle & "-最新推荐" & KS.C_S(ChannelID,3)
		ElseIF sHot<>"0" Then
		  sTitle = sTitle & "-最新热门" & KS.C_S(ChannelID,3)
		Else
		  sTitle = sTitle & "-" & KS.C_S(ChannelID,1)
		End If
		GetRssBody	= GetChannelNewInfo(ChannelID,sClassID,RssChannelNum,RssDescriptNum)
	Else
		sTitle		= sTitle
		Dim RS:Set RS=Conn.Execute("Select ChannelID From KS_Channel Where ChannelStatus=1 And ChannelID<>6 And ChannelID<>9")
		Do While Not RS.Eof
		GetRssBody	= GetRssBody & GetChannelNewInfo(RS(0),sClassID,RssHomeNum,RssDescriptNum) 
		RS.MoveNext
		Loop
		RS.Close:Set RS=Nothing
	End If
End Function

       '分别取得各个模块的最新更新信息
	   '参数：	InfoNum-设定每个模块取得的最新信息数量, DescriptNum 设定每条信息介绍文字字数
       Function GetChannelNewInfo(ChannelID,sClassID,InfoNum,DescriptNum)
	     If ChannelID="" Then GetChannelNewInfo = GetChannelNewInfo & "<item></item>":Exit Function
		 Dim SqlStr,SQL,Rs,i,Param
		  Param=" Where 1=1 "
		 If SclassID<>"0" and SclassID<>"" Then 
		  Param= Param & " And Tid In(" & KS.GetFolderTid(sClassID) & ")"
		 End If
		 IF sElite<>"0" Then
		  Param= Param & " And Recommend=1"
		 End IF
		 IF sHot<>"0" Then
		  Param= Param & " And Popular=1"
		 End IF
		 Select Case KS.C_S(ChannelID,6)
		  Case 1
		   SqlStr="Select Top " &InfoNum & " a.ID,Title,Tid,Fname,AddDate,Intro,Author,FolderName From " & KS.C_S(ChannelID,2) & " a inner join ks_class b on a.tid=b.id " & Param &" And a.DelTF=0 And Verific=1 Order By a.ID Desc"    
		  Case 2
		   SqlStr="Select Top " &InfoNum & " a.ID,Title,Tid,Fname,AddDate,PictureContent,Author,FolderName From " & KS.C_S(ChannelID,2) & " a inner join ks_class b on a.tid=b.id " & Param &" And a.DelTF=0 And Verific=1 Order By a.ID Desc"    
		  Case 3
		   SqlStr="Select Top " &InfoNum & " a.ID,Title,Tid,Fname,AddDate,DownContent,Author,FolderName From " & KS.C_S(ChannelID,2) & " a inner join ks_class b on a.tid=b.id " & Param &" And a.DelTF=0 And Verific=1 Order By a.ID Desc"    
		  Case 4
		   SqlStr="Select Top " &InfoNum & " a.ID,Title,Tid,Fname,AddDate,FlashContent,Author,FolderName From " & KS.C_S(ChannelID,2) & " a inner join ks_class b on a.tid=b.id " & Param &" And a.DelTF=0 And Verific=1 Order By a.ID Desc"    
		  Case 5
		   SqlStr="Select Top " &InfoNum & " a.ID,Title,Tid,Fname,AddDate,ProIntro,ProducerName,FolderName From " & KS.C_S(ChannelID,2) & " a inner join ks_class b on a.tid=b.id " & Param &" And a.DelTF=0 And Verific=1 Order By a.ID Desc"    
		  Case 7
		   SqlStr="Select Top " &InfoNum & " a.ID,Title,Tid,Fname,AddDate,MovieContent,MovieAct,FolderName From " & KS.C_S(ChannelID,2) & " a inner join ks_class b on a.tid=b.id " & Param &" And a.DelTF=0 And Verific=1 Order By a.ID Desc"    
		  Case 8
		   SqlStr="Select Top " &InfoNum & " a.ID,Title,Tid,Fname,AddDate,GQContent,Inputer,FolderName From " & KS.C_S(ChannelID,2) & " a inner join ks_class b on a.tid=b.id " & Param &" And a.DelTF=0 And Verific=1 Order By a.ID Desc"    
		 End Select
		 
		 Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open SqlStr,Conn,1,1
		if Rs.Bof and Rs.Eof then
			'GetChannelNewInfo = GetChannelNewInfo & "<item></item>"
			Rs.Close : Set Rs = Nothing
		Else
			SQL = Rs.GetRows(-1)
			Rs.Close : Set Rs = Nothing
			For i = 0 to UBound(SQL,2)
				GetChannelNewInfo = GetChannelNewInfo & "<item>"
				GetChannelNewInfo = GetChannelNewInfo & "<title><![CDATA[[" & SQL(7,i) & "] " & SQL(1,i) & "]]></title>"
				GetChannelNewInfo = GetChannelNewInfo & "<link><![CDATA[" & KS.GetItemURL(ChannelID,SQL(2,I),SQL(0,I),SQL(3,I)) & "]]></link>"
				If RssDescriptNum<>0 Then
				GetChannelNewInfo = GetChannelNewInfo & "<description><blockquote><![CDATA[" & KS.GotTopic(Replace(KS.LoseHtml(SQL(5,I)), "[NextPage]", ""),DescriptNum) & "......]]></blockquote></description>"
				End IF
				GetChannelNewInfo = GetChannelNewInfo & "<author><![CDATA[" & SQL(6,i) & "]]></author>"
				GetChannelNewInfo = GetChannelNewInfo & "<pubDate><![CDATA[" & SQL(4,i) & "]]></pubDate>"
				GetChannelNewInfo = GetChannelNewInfo & "</item>"
			Next
		End if
	
	   End Function
End Class
%> 

<!--#include file="../conn.asp" -->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.BaseFunCls.asp"-->
<%
Response.Expires = 0
If Not IsNumeric(Request("id")) And Request("id")<>"" then
	Response.write"错误的系统参数!ID必须是数字"
	Response.End
End If
Dim rsOnline,strUsername,statuserid,remoteaddr,onlinemany
Dim Rs,SQL,strReferer,onlinemember,BrowserType,CurrentStation,KS
Set KS=New PublicCls
Application.Lock
remoteaddr = KS.GetIP()
strReferer = CheckInSQL(KS.URLDecode(Request.ServerVariables("HTTP_REFERER")))
If strReferer = Empty Then
	strReferer = "★直接输入或书签导入★"
Else
	strReferer = Left(strReferer,255)
End If
CurrentStation = CheckInSQL(Left(Request.ServerVariables("HTTP_REFERER"),220))

If KS.C("UserName") = "" Then
	strUsername = "匿名用户"
Else
	strUsername = KS.DelSQL(KS.C("UserName"))
End If

Set BrowserType=new SystemInfo_Cls
Call UserActiveOnline
Set BrowserType=Nothing
Application.UnLock
' 删除不活动的用户
Conn.Execute("DELETE FROM KS_Online WHERE DateDIff(" & DataPart_S &",lastTime," & SQLNowString & ") > "& CLng(KS.Setting(8)) &" * 60")
onlinemany = Conn.Execute("Select Count(*) from KS_Online")(0)
onlinemember = Conn.Execute("Select Count(*) from KS_Online where username <> '匿名用户'")(0)


'最高在线
Dim GXml,Doc
Set GXml=LFCls.GetXMLFromFile("guestbook")
If IsObject(GXml) Then
	If onlinemany>KS.ChkClng(GXml.documentElement.attributes.getNamedItem("maxonline").text) Then
         set Doc = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		 Doc.async = false
		 Doc.setProperty "ServerHTTPRequest", true 
		 Doc.load(Server.MapPath(KS.Setting(3)&"Config/guestbook.xml"))
		 Doc.documentElement.attributes.getNamedItem("maxonline").text=onlinemany
		 Doc.documentElement.attributes.getNamedItem("maxonlinedate").text=now
		 Doc.save(Server.MapPath(KS.Setting(3)&"Config/guestbook.xml"))
		 Set Doc=Nothing
		 Application(KS.SiteSN&"_Configguestbook")=empty
	End If
End If
Set GXml=Nothing


If CInt(Request.Querystring("id")) = "1" And Trim(Request.Querystring("id")) <> "" Then
	Response.Write "document.writeln(" & chr(34) & ""& onlinemany &""& chr(34) & ");"
ElseIf CInt(Request.Querystring("id")) = "2" And Trim(Request.Querystring("id")) <> "" Then
	Response.Write "document.writeln(" & Chr(34) & ""& onlinemember &""& chr(34) & ");"
Else
	Response.Write "document.writeln('总在线：<font color=red><b>" & onlinemany & "</b></font> 人 用户：<font color=blue><b>" & onlinemember & "</b></font> 人 游客：<b>" & onlinemany-onlinemember&"</b> 人');"
End If

Sub UserActiveOnline()
	Dim UserSessionID,OnlineSQL
	UserSessionID = Session.sessionid
	SQL = "SELECT top 1 * FROM [KS_Online] WHERE ip='" & remoteaddr & "' And username='" & strUsername & "' Or id=" & UserSessionID
	Set rsOnline = Server.CreateObject("ADODB.Recordset")
	rsOnline.Open SQL,Conn,1,1
	If rsOnline.BOF And rsOnline.EOF Then
		OnlineSQL = "INSERT INTO KS_Online(id,username,station,ip,browser,startTime,lastTime,strReferer) VALUES (" & UserSessionID & ",'" & strUsername & "','" & CurrentStation & "','" & remoteaddr & "','" & BrowserType.platform&"|"&BrowserType.Browser&BrowserType.version & "|"&BrowserType.AlexaToolbar&"'," & SqlNowString & "," & SqlNowString & ",'" & strReferer & "')"
		Call AddCountData
	Else
		OnlineSQL = "UPDATE KS_Online SET ID=" & UserSessionID & ",username='" & strUsername & "',station='" & CurrentStation & "',lastTime=" & SqlNowString & " WHERE ID = " & UserSessionID
		Call UpdateCountData
	End If
	Conn.Execute(OnlineSQL)
	rsOnline.close
	Set rsOnline = Nothing
End Sub
CloseConn
Class SystemInfo_Cls
	Public Browser, version, platform, IsSearch, AlexaToolbar
	Private Sub Class_Initialize()
	    on error resume next
		Dim Agent, Tmpstr
		IsSearch = False
		If Not IsEmpty(Session("SystemInfo_Cls")) Then
			Tmpstr = Split(Session("SystemInfo_Cls"), "|||")
			Browser = Tmpstr(0)
			version = Tmpstr(1)
			platform = Tmpstr(2)
			AlexaToolbar = Tmpstr(4)
			If Tmpstr(3) = "1" Then
				IsSearch = True
			End If
			Exit Sub
		End If
		Browser = "unknown"
		version = "unknown"
		platform = "unknown"
		Agent = CheckInSQL(Request.ServerVariables("HTTP_USER_AGENT"))
		If InStr(Agent, "Alexa Toolbar") > 0 Then
			AlexaToolbar = "YES"
		Else
			AlexaToolbar = "NO"
		End If
		If Left(Agent, 7) = "Mozilla" Then '有此标识为浏览器
			Agent = Split(Agent, ";")
			If InStr(Agent(1), "MSIE") > 0 Then
				Browser = "Internet Explorer "
				version = Trim(Left(Replace(Agent(1), "MSIE", ""), 6))
			ElseIf InStr(Agent(4), "Netscape") > 0 Then
				Browser = "Netscape "
				Tmpstr = Split(Agent(4), "/")
				version = Tmpstr(UBound(Tmpstr))
			ElseIf InStr(Agent(4), "rv:") > 0 Then
				Browser = "Mozilla "
				Tmpstr = Split(Agent(4), ":")
				version = Tmpstr(UBound(Tmpstr))
				If InStr(version, ")") > 0 Then
					Tmpstr = Split(version, ")")
					version = Tmpstr(0)
				End If
			End If
			If InStr(Agent(2), "NT 5.2") > 0 Then
				platform = "Windows 2003"
			ElseIf InStr(Agent(2), "Windows CE") > 0 Then
				platform = "Windows CE"
			ElseIf InStr(Agent(2), "NT 5.1") > 0 Then
				platform = "Windows XP"
			ElseIf InStr(Agent(2), "NT 4.0") > 0 Then
				platform = "Windows NT"
			ElseIf InStr(Agent(2), "NT 5.0") > 0 Then
				platform = "Windows 2000"
			ElseIf InStr(Agent(2), "NT") > 0 Then
				platform = "Windows NT"
			ElseIf InStr(Agent(2), "9x") > 0 Then
				platform = "Windows ME"
			ElseIf InStr(Agent(2), "98") > 0 Then
				platform = "Windows 98"
			ElseIf InStr(Agent(2), "95") > 0 Then
				platform = "Windows 95"
			ElseIf InStr(Agent(2), "Win32") > 0 Then
				platform = "Win32"
			ElseIf InStr(Agent(2), "Linux") > 0 Then
				platform = "Linux"
			ElseIf InStr(Agent(2), "SunOS") > 0 Then
				platform = "SunOS"
			ElseIf InStr(Agent(2), "Mac") > 0 Then
				platform = "Mac"
			ElseIf UBound(Agent) > 2 Then
				If InStr(Agent(3), "NT 5.1") > 0 Then
					platform = "Windows XP"
				End If
				If InStr(Agent(3), "Linux") > 0 Then
					platform = "Linux"
				End If
			End If
			If InStr(Agent(2), "Windows") > 0 And platform = "unknown" Then
				platform = "Windows"
			End If
		ElseIf Left(Agent, 5) = "Opera" Then '有此标识为浏览器
			Agent = Split(Agent, "/")
			Browser = "Mozilla "
			Tmpstr = Split(Agent(1), " ")
			version = Tmpstr(0)
			If InStr(Agent(1), "NT 5.2") > 0 Then
				platform = "Windows 2003"
			ElseIf InStr(Agent(1), "Windows CE") > 0 Then
				platform = "Windows CE"
			ElseIf InStr(Agent(1), "NT 5.1") > 0 Then
				platform = "Windows XP"
			ElseIf InStr(Agent(1), "NT 4.0") > 0 Then
				platform = "Windows NT"
			ElseIf InStr(Agent(1), "NT 5.0") > 0 Then
				platform = "Windows 2000"
			ElseIf InStr(Agent(1), "NT") > 0 Then
				platform = "Windows NT"
			ElseIf InStr(Agent(1), "9x") > 0 Then
				platform = "Windows ME"
			ElseIf InStr(Agent(1), "98") > 0 Then
				platform = "Windows 98"
			ElseIf InStr(Agent(1), "95") > 0 Then
				platform = "Windows 95"
			ElseIf InStr(Agent(1), "Win32") > 0 Then
				platform = "Win32"
			ElseIf InStr(Agent(1), "Linux") > 0 Then
				platform = "Linux"
			ElseIf InStr(Agent(1), "SunOS") > 0 Then
				platform = "SunOS"
			ElseIf InStr(Agent(1), "Mac") > 0 Then
				platform = "Mac"
			ElseIf UBound(Agent) > 2 Then
				If InStr(Agent(3), "NT 5.1") > 0 Then
					platform = "Windows XP"
				End If
				If InStr(Agent(3), "Linux") > 0 Then
					platform = "Linux"
				End If
			End If
		Else
			'识别搜索引擎
			Dim botlist, i
			botlist = "Google,Isaac,Webdup,SurveyBot,Baiduspider,ia_archiver,P.Arthur,FAST-WebCrawler,Java,Microsoft-ATL-Native,TurnitinBot,WebGather,Sleipnir"
			botlist = Split(botlist, ",")
			For i = 0 To UBound(botlist)
				If InStr(Agent, botlist(i)) > 0 Then
					platform = botlist(i) & "搜索器"
					IsSearch = True
					Exit For
				End If
			Next
		End If
		If IsSearch Then
			Browser = ""
			version = ""
			Session("SystemInfo_Cls") = Browser & "|||" & version & "|||" & platform & "|||1|||" & AlexaToolbar
		Else
			Session("SystemInfo_Cls") = Browser & "|||" & version & "|||" & platform & "|||0|||" & AlexaToolbar
		End If
	End Sub
End Class

Sub AddCountData()
	Dim strSQL,oRs
	Dim rowname,cid,strAgent
	rowname = GetSearcher(strReferer)
	If rowname = "3721" Then rowname = "C3721"
	strAgent = CheckInSQL(Request.ServerVariables("HTTP_USER_AGENT"))
	strSQL = "SELECT id FROM [KS_SiteCount] WHERE Datediff(" & DataPart_D & ",CountDate," & SqlNowString &")=0"
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open strSQL,Conn,1,1
	If oRs.BOF And oRs.EOF Then
		If InStr(strAgent, "Alexa Toolbar") > 0 Then
			strSQL = "INSERT INTO KS_SiteCount(UniqueIP,Pageview,CountDate," & rowname & ",AlexaToolbar) VALUES (1,1," & SqlNowString & ",1,1)"
		Else
			strSQL = "INSERT INTO KS_SiteCount(UniqueIP,Pageview,CountDate," & rowname & ") VALUES (1,1," & SqlNowString & ",1)"
		End If
	Else
		If InStr(strAgent, "Alexa Toolbar") > 0 Then
			strSQL = "UPDATE KS_SiteCount SET AlexaToolbar=AlexaToolbar+1 WHERE id=" & oRs("id")
			Conn.Execute(strSQL)
		End If
		strSQL = "UPDATE KS_SiteCount SET UniqueIP=UniqueIP+1,Pageview=Pageview+1," & rowname & "=" & rowname & "+1 WHERE id=" & oRs("id")
	End If
	oRs.Close:Set oRs = Nothing
	Conn.Execute(strSQL)
	strSQL = Empty
End Sub

Sub UpdateCountData()
	Dim strSQL,oRs,rowname,cid,strAgent
	rowname = GetSearcher(strReferer)
	If rowname = "3721" Then rowname = "C3721"
	strAgent = CheckInSQL(Request.ServerVariables("HTTP_USER_AGENT"))
	
	strSQL = "SELECT id FROM [KS_SiteCount] WHERE Datediff(" & DataPart_D & ",CountDate," & SQLNowString & ")=0"
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open strSQL,Conn,1,1
	If oRs.BOF And oRs.EOF Then
		If InStr(strAgent, "Alexa Toolbar") > 0 Then
			strSQL = "INSERT INTO KS_SiteCount(UniqueIP,Pageview,CountDate," & rowname & ",AlexaToolbar) VALUES (1,1," & SqlNowString & ",1,1)"
		Else
			strSQL = "INSERT INTO KS_SiteCount(UniqueIP,Pageview,CountDate," & rowname & ") VALUES (1,1," & SqlNowString & ",1)"
		End If
	Else
		strSQL = "UPDATE KS_SiteCount SET Pageview=Pageview+1 WHERE id=" & oRs("id")
	End If
	oRs.Close:Set oRs = Nothing
	Conn.Execute(strSQL)
	strSQL = Empty
End Sub

Function GetSearcher(ByVal strUrl)
	On Error Resume Next
	If Len(strUrl) < 5 Then
		GetSearcher = "DirectInput"
		Exit Function
	End If
	If strUrl = "★直接输入或书签导入★" Or InStr(strUrl, ":") = 0 Then
		GetSearcher = "DirectInput"
		Exit Function
	End If
	Dim Searchlist,i,SearchName
	
	strUrl = Left(strUrl, InStr(10, strUrl, "/") - 1)
	strUrl = LCase(strUrl)
	Searchlist = "google,baidu,yahoo,3721,zhongsou,sogou"
	
	Searchlist = Split(Searchlist, ",")
	For i = 0 To UBound(Searchlist)
		If InStr(strUrl, Searchlist(i)) > 0 Then
			SearchName = Searchlist(i)
			Exit For
		Else
			SearchName = "other"
		End If
	Next
	GetSearcher = SearchName
End Function

Function CheckInSQL(str)
	If IsNull(str) Then Exit Function
	On Error Resume Next
	Dim s,Badstring,i
	Badstring = " and | mid |exec|insert|select|delete|update|count|master|truncate|char|declare"
	str = Replace(str, Chr(0), ""):		str = Replace(str, Chr(9), " ")
	str = Replace(str, Chr(255), " "):	str = Replace(str, "　", " ")
	str = Replace(str, "'", "''"):		str = Replace(str, "--", "－－")
	str = Replace(str, "@", "＠"):		str = Replace(str, "*", "＊")
	str = Replace(str, "%", "％"):		str = Replace(str, "^", "＾")
	Badstring = Split(Badstring, "|")
	s = LCase(str)
	s = Replace(s, Chr(10), ""):s = Replace(s, Chr(13), "")
	For i = 0 To UBound(Badstring)
		If InStr(s, Badstring(i))>0 Then
			CheckInSQL = ""
			Exit Function
		End If
	Next
	CheckInSQL = str
End Function
%> 

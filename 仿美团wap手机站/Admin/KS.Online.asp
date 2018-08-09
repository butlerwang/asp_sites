<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Admin_Online
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Online
        Private KS
		Private Action,i,strClass,sFileName,RS,SQL,maxperpage,CurrentPage,totalPut,TotalPageNum
        Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub

		Public Sub Kesion()
		With Response
			If Not KS.ReturnPowerResult(0, "KSMS20005") Then
			  Call KS.ReturnErr(1, "")
			  exit sub
			End If
		.Write "<script src='../KS_Inc/common.js'></script>"
		.Write "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
		.Write "<ul id='menu_top'>"
		.Write "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
		.Write "  <tr>"
		.Write "    <td height=""23"" align=""left"" valign=""top"">"
		.Write "	<td ><b>菜单导航：</b><a href='KS.Online.asp'>管理首页</a> | "
		.Write "	<a href='KS.Online.asp?action=zone'>详细地址</a> | "
		.Write "	<a href='KS.Online.asp?action=refer'>访问来源</a> | "
		.Write "	<a href='KS.Online.asp?action=online'>流量统计</a> | "
		.Write "	<a href='KS.Online.asp?action=delall' onclick=""{if(confirm('您确定要删除所有在线人数吗?')){return true;}return false;}""><font color=blue>删除所有在线人数</font></a></td>"
		.Write "    </td>"
		.Write "  </tr>"
		.Write "</table>"
		.Write "</ul>"
		End With	
		
		
		maxperpage = 30 '###每页显示数

		If KS.G("page")<> "" Then
			CurrentPage = KS.ChkClng(KS.G("page"))
		Else
			CurrentPage = 1
		End If
		If CurrentPage = 0 Then CurrentPage = 1
		totalPut = Conn.Execute("Select Count(ID) from KS_Online")(0)
		TotalPageNum = CInt(totalPut / maxperpage)  '得到总页数
		If TotalPageNum < totalPut / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
		
		Action = LCase(Request("action"))
		Select Case Trim(Action)
			Case "refer"
				Call OnlineReferer
			Case "zone"
				Call OnlineZone
			Case "del"
				Call DelOnline
			Case "delall"
				Call DelAllOnline
			Case "online"
				Call OnlineCount
			Case "remove"
				Call DelCount
			Case "removeall"
				Call DelAllCount
			Case Else
				Call showmain
		End Select
End Sub

Private Sub showmain()
%>
<table width="100%" border="0" align="center" cellspacing="1" cellpadding="1">
<tr height="25" align="center" class='sort'>
	<td width='5%' nowrap>选择</th>
	<td nowrap>用户名</th>
	<td nowrap>访问时间</th>
	<td nowrap>活动时间</th>
	<td nowrap>用户IP地址</th>
	<td nowrap>操作系统</th>
	<td nowrap>浏览器</th>
</tr>
<%
	sFileName = "KS.Online.asp?"
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from [KS_Online] order by startTime desc"
	Rs.Open SQL, Conn, 1, 1
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr><td height=""25"" align=center bgcolor=""#ffffff"" colspan=7>当前无人在线！</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<form name=selform method=post action=?action=del>
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
		
%>
<tr height="22" align=center class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
	<td><input type=checkbox name=OnlineID value='<%=Rs("id")%>'></td>
	<td><%=Rs("username")%></td>
	<td><%=Rs("startTime")%></td>
	<td><%=Rs("lastTime")%></td>
	<td><%=Rs("ip")%></td>
	<td><%=usersysinfo(Rs("browser"), 0)%></td>
	<td><%=usersysinfo(Rs("browser"), 1)%></td>
</tr>
<tr><td colspan=7 background='images/line.gif'></td></tr>
<%
		Rs.movenext
			i = i + 1
			If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
<tr>
	<td  class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'" height='25' colspan=7>
	&nbsp;&nbsp;<input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">全选
	<input class=Button type="submit" name="Submit2" value=" 删 除 " onclick="{if(confirm('您确定要删除此在线人员吗?')){this.document.selform.submit();return true;}return false;}">
	<input class=Button type="button" name="Submit3" value="清空所有在线人数" onclick="{if(confirm('您确定要清空所有在线人数吗?')){location.href='KS.Online.asp?action=delall';return true;}return false;}"></td>
</tr>
</form>
<tr>
	<td  class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'" colspan=7 align=right>
	<%
	 Call KS.ShowPageParamter(totalPut, MaxPerPage, "KS.Online.asp", True, "条", CurrentPage, "Action=" & Action)
	%></td>
</tr>
</table>

<%
End Sub

Private Sub OnlineReferer()
%>
<table width="100%" border="0" align="center"cellspacing="1" cellpadding="1">
<tr height="25" align="center" class='sort'>
	<td width='5%' nowrap>选择</th>
	<td width='15%' nowrap>来访时间/IP</th>
	<td>访问来源</th>
	<td>当前位置</th>
	<td width='5%' nowrap>Alexa</th>
</tr>
<%
	sFileName = "KS.Online.asp?action=refer&"
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from [KS_Online] order by startTime desc"
	Rs.Open SQL, Conn, 1
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr><td align=center colspan=4 class=TableRow2>当前无人在线！</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<form name=selform method=post action=?action=del>
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
%>
<tr height="25" align=center  class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
	<td align=center><input type=checkbox name=OnlineID value='<%=Rs("id")%>'></td>
	<td align=center nowrap><%=Rs("startTime")%><br><%=Rs("ip")%></td>
	<td><a href='<%=Rs("strReferer")%>' target=_blank><%=Rs("strReferer")%></a></td>
	<td><a href='<%=Rs("station")%>' target=_blank><%=Rs("station")%></a></td>
	<td align=center><a href="http://www.alexa.com/data/details/traffic_details?q=&url=<%=Replace(Replace(KS.GetDomain,"http://",""),"/","")%>" target="_blank"><%=usersysinfo(Rs("browser"), 2)%></a></td>
</tr>
<tr><td colspan=6 background='images/line.gif'></td></tr>
<%
		Rs.movenext
			i = i + 1
			If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
<tr>
	<td  class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'" height='25' colspan=7>
	&nbsp;&nbsp;<input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">全选
	<input class=Button type="submit" name="Submit2" value=" 删 除 " onclick="{if(confirm('您确定要删除此在线人员吗?')){this.document.selform.submit();return true;}return false;}">
	<input class=Button type="button" name="Submit3" value="清空所有在线人数" onclick="{if(confirm('您确定要清空所有在线人数吗?')){location.href='KS.Online.asp?action=delall';return true;}return false;}"></td>
</tr>
</form>
<tr>
	<td colspan=5 class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'"><%
	 Call KS.ShowPageParamter(totalPut, MaxPerPage, "KS.Online.asp", True, "条", CurrentPage, "Action=" & Action)
	%></td>
</tr>
</table>

<%
End Sub

Private Sub OnlineZone()
%>
<table width="100%" border="0" align="center"cellspacing="1" cellpadding="1">
<tr height="22" align="center" class='sort'>
	<td width='5%' nowrap>选择</th>
	<td nowrap>用户名</th>
	<td nowrap>IP地址</th>
	<td nowrap>详细地址</th>
	<td nowrap>操作系统</th>
	<td nowrap>浏览器</th>
</tr>
<%
	sFileName = "KS.Online.asp?action=zone&"
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from [KS_Online] order by startTime desc"
	Rs.Open SQL, Conn, 1
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr><td align=center colspan=7 class=TableRow2>当前无人在线！</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>

<form name=selform method=post action=?action=del>
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
%>
<tr height="25" align=center  class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
	<td><input type=checkbox name=OnlineID value='<%=Rs("id")%>'></td>
	<td><%=Rs("username")%></td>
	<td><%=Rs("ip")%></td>
	<td><%=GetAddress(Rs("ip"))%></td>
	<td><%=usersysinfo(Rs("browser"), 0)%></td>
	<td><%=usersysinfo(Rs("browser"), 1)%></td>
</tr>
<tr><td colspan=6 background='images/line.gif'></td></tr>
<%
		Rs.movenext
			i = i + 1
			If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
<tr>
	<td  class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'" height='25' colspan=7>
	&nbsp;&nbsp;<input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">全选
	<input class=Button type="submit" name="Submit2" value=" 删 除 " onclick="{if(confirm('您确定要删除此在线人员吗?')){this.document.selform.submit();return true;}return false;}">
	<input class=Button type="button" name="Submit3" value="清空所有在线人数" onclick="{if(confirm('您确定要清空所有在线人数吗?')){location.href='KS.Online.asp?action=delall';return true;}return false;}"></td>
</tr>
</form>
<tr>
	<td colspan=7 class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'"><%
	 Call KS.ShowPageParamter(totalPut, MaxPerPage, "KS.Online.asp", True, "条", CurrentPage, "Action=" & Action)
	%></td>
</tr>
</table>

<%
End Sub

Private Sub OnlineCount()
'Conn.Execute ("UPDATE [KS_SiteCount] SET AlexaToolbar=0")
		If KS.G("page")<> "" Then
			CurrentPage = KS.ChkClng(KS.G("page"))
		Else
			CurrentPage = 1
		End If
%>
<table width="100%" border="0" align="center"cellspacing="1" cellpadding="1">
<tr height="22" align="center" class='sort'>
	<td width='5%' nowrap>选择</th>
	<td nowrap>统计日期</th>
	<td nowrap>唯一IP</th>
	<td nowrap>访问数</th>
	<td nowrap>Google</th>
	<td nowrap>百度</th>
	<td nowrap>雅虎索</th>
	<td nowrap>3721索</th>
	<td nowrap>中搜</th>
	<td nowrap>搜狗</th>
	<td nowrap>其它站点</th>
	<td nowrap>直接访问</th>
	<td nowrap>Alexa</th>
</tr>
<%
	totalPut = Conn.Execute("SELECT COUNT(id) FROM KS_SiteCount")(0)
	TotalPageNum = CInt(totalPut / maxperpage)  '得到总页数
	If TotalPageNum < totalPut / maxperpage Then TotalPageNum = TotalPageNum + 1
	If CurrentPage < 1 Then CurrentPage = 1
	If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum

	sFileName = "KS.Online.asp?action=online&"
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM [KS_SiteCount] ORDER BY CountDate DESC,id DESC"
	Rs.Open SQL, Conn, 1
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr><td align=center colspan=13 class=TableRow2>没有流量统计！</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>

<form name=selform method=post action=?action=remove>
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
%>
<tr height="22" align=center  class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
	<td><input type=checkbox name=id value='<%=Rs("id")%>'></td>
	<td nowrap><%=FormatDateTime(Rs("CountDate"),1)%></td>
	<td><%=Rs("UniqueIP")%></td>
	<td><%=Rs("Pageview")%></td>
	<td><%=Rs("google")%></td>
	<td><%=Rs("baidu")%></td>
	<td><%=Rs("yahoo")%></td>
	<td><%=Rs("C3721")%></td>
	<td><%=Rs("zhongsou")%></td>
	<td><%=Rs("sogou")%></td>
	<td><%=Rs("other")%></td>
	<td><%=Rs("DirectInput")%></td>
	<td><%=Rs("AlexaToolbar")%></td>
</tr>
<tr><td colspan=16 background='images/line.gif'></td></tr>
<%
		Rs.movenext
			i = i + 1
			If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
<tr>
	<td  class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'" height='25' colspan=13>
	&nbsp;&nbsp;<input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">全选
	<input class=Button type="submit" name="Submit2" value=" 删 除 " onclick="{if(confirm('您确定要删除此统计吗?')){this.document.selform.submit();return true;}return false;}">
	<input class=Button type="button" name="Submit3" value="清空所有浏览统计" onclick="{if(confirm('您确定要清空所有浏览统计吗?')){location.href='KS.Online.asp?action=removeall';return true;}return false;}"></td>
</tr>
</form>
<tr>
	<td  class='list' align='right' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'" colspan=17><%
	 Call KS.ShowPageParamter(totalPut, MaxPerPage, "KS.Online.asp", True, "条", CurrentPage, "Action=" & Action)
	%></td>
</tr>
</table>
<%
End Sub

Private Sub DelAllOnline()
	Conn.Execute("DELETE FROM KS_Online")
	Call KS.Alert("在线人数全部清除完成！","KS.Online.asp")
End Sub

Private Sub DelAllCount()
	Conn.Execute("DELETE FROM KS_SiteCount")
	Call KS.Alert ("流量统计全部清除完成！","KS.Online.asp")
End Sub

Private Sub DelCount()
	Dim cid
	If Request("id") <> "" Then
		cid = Request("id")
		Conn.Execute("DELETE FROM KS_SiteCount WHERE id in (" & cid & ")")
		Call KS.AlertHintScript("流量统计删除成功！")
	Else
		Call KS.AlertHistory("请选择正确的系统参数！",-1)
	End If
End Sub

Private Sub DelOnline()
	Dim OnlineID
	If Request("OnlineID") <> "" Then
		OnlineID = Request("OnlineID")
		Conn.Execute("DELETE FROM KS_Online WHERE ID in (" & OnlineID & ")")
		Call KS.AlertHintScript("在线人数删除成功！")
	Else
		Call KS.AlertHistory("请选择正确的系统参数！",-1)
	End If
End Sub

Private Function usersysinfo(info, getinfo)
	Dim usersys
	usersys = Split(info, "|")
	usersysinfo = usersys(getinfo)
End Function

Public Function GetAddress(sip)
	If Len(sip) < 5 Then
		GetAddress = "未知"
		Exit Function
	End If
	On Error Resume Next
	Dim Wry,IPType
	Set Wry = New TQQWry
	If Not Wry.IsIp(sip) Then
		GetAddress = " 未知"
		Exit Function
	End If
	IPType = Wry.QQWry(sip)
	GetAddress = Wry.Country & " " & Wry.LocalStr
End Function
End Class
Class TQQWry
	' ============================================
	' 变量声名
	' ============================================
	Dim Country, LocalStr, Buf, OffSet
	Private StartIP, EndIP, CountryFlag
	Public QQWryFile
	Public FirstStartIP, LastStartIP, RecordCount
	Private Stream, EndIPOff
	' ============================================
	' 类模块初始化
	' ============================================
	Private Sub Class_Initialize
		On Error Resume Next
		Country 		= ""
		LocalStr 		= ""
		StartIP 		= 0
		EndIP 			= 0
		CountryFlag 	= 0 
		FirstStartIP 	= 0 
		LastStartIP 	= 0 
		EndIPOff 		= 0 
		QQWryFile = Server.MapPath("../KS_Data/IPAddress.Dat") 'QQ IP库路径，要转换成物理路径
	End Sub
	' ============================================
	' IP地址转换成整数
	' ============================================
	Function IPToInt(IP)
		Dim IPArray, i
		IPArray = Split(IP, ".", -1)
		FOr i = 0 to 3
			If Not IsNumeric(IPArray(i)) Then IPArray(i) = 0
			If CInt(IPArray(i)) < 0 Then IPArray(i) = Abs(CInt(IPArray(i)))
			If CInt(IPArray(i)) > 255 Then IPArray(i) = 255
		Next
		IPToInt = (CInt(IPArray(0))*256*256*256) + (CInt(IPArray(1))*256*256) + (CInt(IPArray(2))*256) + CInt(IPArray(3))
	End Function
	' ============================================
	' 整数逆转IP地址
	' ============================================
	Function IntToIP(IntValue)
		p4 = IntValue - Fix(IntValue/256)*256
		IntValue = (IntValue-p4)/256
		p3 = IntValue - Fix(IntValue/256)*256
		IntValue = (IntValue-p3)/256
		p2 = IntValue - Fix(IntValue/256)*256
		IntValue = (IntValue - p2)/256
		p1 = IntValue
		IntToIP = Cstr(p1) & "." & Cstr(p2) & "." & Cstr(p3) & "." & Cstr(p4)
	End Function
	' ============================================
	' 获取开始IP位置
	' ============================================
	Private Function GetStartIP(RecNo)
		OffSet = FirstStartIP + RecNo * 7
		Stream.Position = OffSet
		Buf = Stream.Read(7)
		
		EndIPOff = AscB(MidB(Buf, 5, 1)) + (AscB(MidB(Buf, 6, 1))*256) + (AscB(MidB(Buf, 7, 1))*256*256) 
		StartIP  = AscB(MidB(Buf, 1, 1)) + (AscB(MidB(Buf, 2, 1))*256) + (AscB(MidB(Buf, 3, 1))*256*256) + (AscB(MidB(Buf, 4, 1))*256*256*256)
		GetStartIP = StartIP
	End Function
	' ============================================
	' 获取结束IP位置
	' ============================================
	Private Function GetEndIP()
		Stream.Position = EndIPOff
		Buf = Stream.Read(5)
		EndIP = AscB(MidB(Buf, 1, 1)) + (AscB(MidB(Buf, 2, 1))*256) + (AscB(MidB(Buf, 3, 1))*256*256) + (AscB(MidB(Buf, 4, 1))*256*256*256) 
		CountryFlag = AscB(MidB(Buf, 5, 1))
		GetEndIP = EndIP
	End Function
	' ============================================
	' 获取地域信息，包含国家和和省市
	' ============================================
	Private Sub GetCountry(IP)
		If (CountryFlag = 1 Or CountryFlag = 2) Then
			Country = GetFlagStr(EndIPOff + 4)
			If CountryFlag = 1 Then
				LocalStr = GetFlagStr(Stream.Position)
				' 以下用来获取数据库版本信息
				If IP >= IPToInt("255.255.255.0") And IP <= IPToInt("255.255.255.255") Then
					LocalStr = GetFlagStr(EndIPOff + 21)
					Country = GetFlagStr(EndIPOff + 12)
				End If
			Else
				LocalStr = GetFlagStr(EndIPOff + 8)
			End If
		Else
			Country = GetFlagStr(EndIPOff + 4)
			LocalStr = GetFlagStr(Stream.Position)
		End If
		' 过滤数据库中的无用信息
		Country = Trim(Country)
		LocalStr = Trim(LocalStr)
		If InStr(Country, "CZ88.NET") Then Country = "GZ110.CN"
		If InStr(LocalStr, "CZ88.NET") Then LocalStr = "GZ110.CN"
	End Sub
	' ============================================
	' 获取IP地址标识符
	' ============================================
	Private Function GetFlagStr(OffSet)
		Dim Flag
		Flag = 0
		Do While (True)
			Stream.Position = OffSet
			Flag = AscB(Stream.Read(1))
			If(Flag = 1 Or Flag = 2 ) Then
				Buf = Stream.Read(3) 
				If (Flag = 2 ) Then
					CountryFlag = 2
					EndIPOff = OffSet - 4
				End If
				OffSet = AscB(MidB(Buf, 1, 1)) + (AscB(MidB(Buf, 2, 1))*256) + (AscB(MidB(Buf, 3, 1))*256*256)
			Else
				Exit Do
			End If
		Loop
		
		If (OffSet < 12 ) Then
			GetFlagStr = ""
		Else
			Stream.Position = OffSet
			GetFlagStr = GetStr() 
		End If
	End Function
	' ============================================
	' 获取字串信息
	' ============================================
	Private Function GetStr() 
		Dim c
		GetStr = ""
		Do While (True)
			c = AscB(Stream.Read(1))
			If (c = 0) Then Exit Do 
			
			'如果是双字节，就进行高字节在结合低字节合成一个字符
			If c > 127 Then
				If Stream.EOS Then Exit Do
				GetStr = GetStr & Chr(AscW(ChrB(AscB(Stream.Read(1))) & ChrB(C)))
			Else
				GetStr = GetStr & Chr(c)
			End If
		Loop 
	End Function
	' ============================================
	' 核心函数，执行IP搜索
	' ============================================
	Public Function QQWry(DotIP)
		Dim IP, nRet
		Dim RangB, RangE, RecNo
		
		IP = IPToInt (DotIP)
		
		Set Stream = CreateObject("ADodb.Stream")
		Stream.Mode = 3
		Stream.Type = 1
		Stream.Open
		Stream.LoadFromFile QQWryFile
		Stream.Position = 0
		Buf = Stream.Read(8)
		
		FirstStartIP = AscB(MidB(Buf, 1, 1)) + (AscB(MidB(Buf, 2, 1))*256) + (AscB(MidB(Buf, 3, 1))*256*256) + (AscB(MidB(Buf, 4, 1))*256*256*256)
		LastStartIP  = AscB(MidB(Buf, 5, 1)) + (AscB(MidB(Buf, 6, 1))*256) + (AscB(MidB(Buf, 7, 1))*256*256) + (AscB(MidB(Buf, 8, 1))*256*256*256)
		RecordCount = Int((LastStartIP - FirstStartIP)/7)
		' 在数据库中找不到任何IP地址
		If (RecordCount <= 1) Then
			Country = "未知"
			QQWry = 2
			Exit Function
		End If
		
		RangB = 0
		RangE = RecordCount
		
		Do While (RangB < (RangE - 1)) 
			RecNo = Int((RangB + RangE)/2) 
			Call GetStartIP (RecNo)
			If (IP = StartIP) Then
				RangB = RecNo
				Exit Do
			End If
			If (IP > StartIP) Then
				RangB = RecNo
			Else 
				RangE = RecNo
			End If
		Loop
		
		Call GetStartIP(RangB)
		Call GetEndIP()

		If (StartIP <= IP) And ( EndIP >= IP) Then
			' 没有找到
			nRet = 0
		Else
			' 正常
			nRet = 3
		End If
		Call GetCountry(IP)

		QQWry = nRet
	End Function
	' ============================================
	' 检查IP地址合法性
	' ============================================
	Public Function IsIp(IP)
		IsIp = True
		If IP = "" Then IsIp = False : Exit Function
		Dim Re
		Set Re = New RegExp
		Re.Pattern = "^(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5])\.(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5])\.(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5])\.(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5])$"
		Re.IgnoreCase = True
		Re.Global = True
		IsIp = Re.Test(IP)
		Set Re = Nothing
	End Function
End Class

%> 

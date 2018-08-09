<%@ Language="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%


Dim KS:Set KS=New PublicCls
Dim KSUser: Set KSUser = New UserCls
Dim ID:ID = Replace(KS.S("ID")," ","")
Dim ChannelID:ChannelID=KS.ChkClng(Request("m"))
If ChannelID="" Then Response.End()
Dim LoginTF,ComeUrl,ClassID,UserName
ID=KS.FilterIDs(ID)
If ID="" Then Response.Write("<script>alert('对不起，您没有选择投票项!');history.back();</script>"):Response.End()

Const UserTF=1         '是否只允许会员投票 1是 0否
Const UserIPNum=3      '每个IP最多投票数，0不限制 3表示限制3票
Const SameVote=0       'UserIPNum如果设置大于0时，是否允许投在同一个选项上，0不允许，1允许
Const UserGroup="0"    '允许投票的会员组，多个会员组请用,号隔开，不想限制请输入0
Const LimitTime=2            '间隔时间设置，单位分种，如2表示同一个IP两分钟后才可以再投,不限制请输入0
Const BeginTime="0"   '投票开始时间，不限制请输入0 格式：YYYY-MM-DD hh:mm:ss
Const EndTime="0"   '投票结束时间，结束后将不能再投票了,不限制请输入0 格式：YYYY-MM-DD hh:mm:ss

'IF Cbool(Request.Cookies(Cstr(ID))("PhotoVote"))<>true Then
' Conn.Execute("Update " & KS.C_S(ChannelID,2) &" Set Score=Score+1 Where ID=" & ID)
' Response.Cookies(Cstr(ID))("PhotoVote")=true
' Response.Write "<script>alert('感谢您的投票！');location.href='" & Request.ServerVariables("HTTP_REFERER") & "';<//script>"
'Else
'Response.Write "<script>alert('你已经投过票，不能再投了！');location.href='" & Request.ServerVariables("HTTP_REFERER") & "';''<//script>"
'End IF

LoginTF=KSUser.UserLoginChecked()
ComeUrl=Request.ServerVariables("HTTP_REFERER")
ClassID=Conn.Execute("Select top 1 Tid From " & KS.C_S(ChannelID,2) & " where ID In(" & ID & ")")(0)

If KS.S("Action")="Show" Then 
 Call ShowVote()
Else
 Call Vote()
End If

Sub Vote()
	If UserTF=1 and LoginTF=False Then
	   Response.Write "<script>alert('对不起，只会登录会员才能投票!');history.back(-1);</script>"
	   Response.End()
	End If
	
	if UserGroup<>"0" and KS.FoundInArr(UserGroup, KSUser.GroupID, ",")=False Then
	   Response.Write "<script>alert('对不起，您所在的会员组不允许投票!');history.back(-1);</script>"
	   Response.End()
	End If
	
	If BeginTime<>"0" Then
	  If DateDiff("s",BeginTime,Now)<0 Then
			Response.Write "<script>alert('对不起，开始投票时间为：" & BeginTime & "！');history.back();</script>"
			Response.End()
	  End If
	End If
	
	If EndTime<>"0" Then
	  If DateDiff("s",EndTime,Now)>0 Then
			Response.Write "<script>alert('对不起，投票已结束，结束时间为：" & EndTime & "！');history.back();</script>"
			Response.End()
	  End If
	End If
	
	If LimitTime<>0 Then
	  Set RS=Server.CreateObject("adodb.recordset")
	  RS.Open "select top 1 * From KS_PhotoVote Where UserIp='" & KS.GetIP &"' and channelid=" & ChannelID &" order by id desc",conn,1,1
	  If Not RS.Eof Then
	    Dim LastvoteTime : LastVoteTime=RS("VoteTime")
		If DateDiff("n",LastVoteTime,now)<LimitTime Then
			Response.Write "<script>alert('对不起，" & LimitTime & "分钟后才可以再参与投票！');history.back();</script>"
			Response.End()
		End If
	  End If
	  RS.Close
	  Set RS=Nothing
	End If
	
	If UserIPNum<>0 Then
	  '判断有没有超过最大投票数了
	  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
	  RS.Open "Select ID From KS_PhotoVote Where UserIp='" & KS.GetIP & "' and ChannelID=" & ChannelID & " And ClassID='" & ClassID & "'",conn,1,1
	  If Not RS.Eof Then
	   If KS.ChkClng(RS.Recordcount)>KS.ChkCLng(UserIPNum)  Then
	    RS.Close:Set RS=Nothing
	    Response.Write "<script>alert('对不起，每人只能投" &UserIPNum &"票！');history.back();</script>"
	    Response.End()
	   End If
	   '判断是不是投了同一选项
	   If SameVote=0 Then
	    Dim RSS:Set RSS=Conn.Execute("Select top 1 ID From KS_PhotoVote Where UserIp='" & KS.GetIP & "' and ChannelID=" & ChannelID & " And ClassID='" & ClassID & "' And InfoID='" & KS.ChkClng(ID) & "'")
		If Not RSS.Eof Then
		  RSS.CLose:Set RSS=Nothing
			Response.Write "<script>alert('对不起，同一个投票项只能投一次！');history.back();</script>"
			Response.End()
		End If
		RSS.CLose:Set RSS=Nothing
	   End If
	  End If
	  RS.Close: Set RS=Nothing
	End If
	
	
	If LoginTF=False Then UserName="游客" Else UserName=KSUser.UserName
	Conn.Execute("Insert Into [KS_PhotoVote]([ChannelID],[ClassID],[InfoID],[VoteTime],[UserName],[UserIP]) Values(" & ChannelID & ",'" & ClassID & "','" & ID & "'," & SqlNowString & ",'" & UserName & "','" & KS.GetIP() & "')")
	Conn.Execute("Update " & KS.C_S(ChannelID,2) &" Set Score=Score+1 Where ID In(" & ID & ")")
	
	KS.AlertHintScript "恭喜，您已成功的投票！"
End Sub

Sub ShowVote()
   Dim TempStr
    TempStr = TempStr & "<table width=""99%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""2"">"
    TempStr = TempStr & "     <tr> "
	TempStr = TempStr & "			<td width=""200"" align=""center""><strong>投票选项</strong></td>"
	TempStr = TempStr & "			<td width=""100"" align=""center""><strong>得票柱状图</strong></td>"
	TempStr = TempStr & "	    	<td  align=""center""><strong>百分比</strong></td>"
	TempStr = TempStr & "	 </tr>"
		
			Dim TotalVote:TotalVote=Conn.Execute("Select sum(score) from " & KS.C_S(ChannelID,2) & " where tid='" & ClassID & "'")(0)
			if totalvote=0 then totalvote=1
			Dim RS:Set RS=Conn.Execute("Select Title,Score From " & KS.C_S(ChannelID,2) & " where tid='" & ClassID & "' Order BY Score Desc")
			Do While Not RS.Eof
			
	TempStr = TempStr & "	  <tr> "
	TempStr = TempStr & "		<td height=""25"" style=""BORDER-BOTTOM: 1px solid"" align=""center"">" & rs(0) & "</td>"
	TempStr = TempStr & "		<td  style=""BORDER-BOTTOM: 1px solid"" align=""center"">" & rs(1) & "</td>"
	TempStr = TempStr & "		<td style=""BORDER-BOTTOM: 1px solid""> "
			
			dim perVote:perVote=round(rs(1)/totalVote,4)
	TempStr = TempStr & "<img src='../images/Default/bar.gif' width='" & round(360*perVote) & "' height='15' align='absmiddle'>"
			perVote=perVote*100
			if perVote<1 and perVote<>0 then
				TempStr = TempStr & "&nbsp;0" & perVote & "%"
			else
				TempStr = TempStr & "&nbsp;" & perVote & "%"
			end if
	
	TempStr = TempStr & "</td>"
	TempStr = TempStr & "</tr>"
			RS.MoveNext 
		Loop
		
	TempStr = TempStr & "</table>"
	Set KSR = New Refresh
	Dim Template
	Template=KSR.LoadTemplate(KS.Setting(3) & KS.Setting(90) & "vote.html")  '模板地址
	Template=Replace(Template,"{$ShowVoteResult}",TempStr)
	Response.Write Template
	Set KSR=Nothing
End Sub


Call CloseConn()
Set KS=Nothing
Set KSUser=Nothing
%> 

<%@ Language="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<%

response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Response.CharSet="utf-8"
Dim KS:Set KS=New PublicCls
Dim Action:Action=KS.S("Action")
Dim MoodID:MoodID=KS.ChkClng(KS.S("ID"))
Dim ChannelID:ChannelID=KS.ChkCLng(KS.S("M_ID"))
Dim InfoID:InfoID=KS.ChkCLng(KS.S("C_ID"))
Dim UserName,LoginTF,ProjectContent,VoteArr,I,VoteStr,VoteItemArr
Dim OnlyUser,UserOnce,AllowGroupID,TimeLimit,StartDate,ExpiredDate,MoodStatus

If ChannelID=0 or InfoID=0 Then Response.End()


Set KSUser = New UserCls

Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
RS.Open "Select Top 1 * From KS_MoodProject Where Status=1 and ID="& MoodID,conn,1,1
If RS.Eof And RS.Bof Then
 RS.Close:Set RS=Nothing::Set KS=Nothing:Set KSUser=Nothing:CloseConn()
 Response.End()
End If
ProjectContent = RS("ProjectContent")
OnlyUser       = RS("OnlyUser")
UserOnce       = RS("UserOnce")
AllowGroupID   = RS("AllowGroupID")
TimeLimit      = RS("TimeLimit")
StartDate      = RS("StartDate")
ExpiredDate    = RS("ExpiredDate")
MoodStatus     = RS("Status")
RS.Close:Set RS=Nothing


if Action="hits" Then
 response.write "var vote={'status':'" &Vote() & "'};"
 response.end
ElseIf Action="ShowPopup" Then
  ShowPopup
Else
  GetVoteStr()
End If

'投票操作
Function Vote()
  LoginTF=KSUser.UserLoginChecked()
  UserName=KSUser.UserName
  If LoginTF=false and OnlyUser="1" Then
   Vote= "nologin":Exit Function
  Else
    If MoodStatus=0 Then
	 Vote= "lock":Exit Function
	Else
	    If UserOnce=1 And KS.C("mood_"&ChannelID&"_"&InfoID)="ok" Then Vote="standoff":Exit Function
	    If TimeLimit="1" then
		 if now<StartDate then Vote= "errstartdate":Exit Function
		 if now>ExpiredDate then Vote= "errexpireddate":Exit Function
	    End If
		if AllowGroupID<>"" then
	     if KS.FoundInArr(AllowGroupID,KSUser.groupid,",")=false then  Vote= "errgroupid":Exit Function
	    end if
		Dim score:score=KS.S("score")
		If Not IsNumeric(score) Then score=0
	    Dim RSI:Set RSI=Server.CreateObject("ADODB.RECORDSET")
		RSI.Open "Select top 1 * From KS_MoodList Where MoodID=" & MoodID & " and ChannelID="& ChannelID &" And InfoID=" & InfoID,Conn,1,3
		If RSI.Eof And RSI.Bof Then
		 RSI.AddNew
		 RSI("MoodID")=MoodID
		 RSI("HitsNum")=1
		 RSI("Score")=score
		 RSI("AvgScore")=score
		 RSI("Title")=KS.CheckXss(KS.LoseHtml(KS.S("Title")))
		 RSI("ChannelID")=ChannelID
		 RSI("InfoID")=InfoID
		Else
		 RSI("HitsNum")=RSI("HitsNum")+1
		 RSI("score")=RSI("score")+score
		 If RSI("HitsNum")<>0 And RSI("Score")<>0 Then
		  RSI("AvgScore")=RSI("Score")/RSI("HitsNum")
		 Else
		  RSI("AvgScore")=0
		 End If
		End If
		 For I=0 To 14 
		  If Trim(I)=KS.S("itemid") Then  RSI("M" & i)=RSI("M"&i)+1
		 Next
		 RSI.Update
		RSI.Close:Set RSI=Nothing
		Response.Cookies(KS.SiteSn)("mood_"&ChannelID&"_"&InfoID)="ok"
	    Vote= "success"
	End If
  End If
End Function

Function GetVoteStr()
     Dim RST,TotalVote,PerVote,ImgSrc,HitsNum,AvgScore,ListId
	Dim ItemVoteNum(15)
	Set RST=Server.CreateObject("ADODB.RECORDSET")
	RST.Open "Select top 1 * From KS_MoodList Where ChannelID=" & ChannelID & " And InfoID=" & InfoID & " And MoodID=" & MoodID,conn,1,1
	If Not RST.Eof Then
	 ListId=RST("ID")
	 TotalVote=RST("Score")
	 AvgScore=RST("AvgScore")
	 For I=0 To 14
	  ItemVoteNum(i)=RST("M" & I)
	 Next
	Else
	 ListId=0
	 TotalVote=0
	 AvgScore=0
	 For I=0 To 14
	  ItemVoteNum(i)=0
	 Next
	End If
	RST.Close
	Set RST=Nothing
	
	If AvgScore<>0 And Instr(AvgScore,".")<>0 Then	AvgScore=Formatnumber(AvgScore,2,-1,0,-1)
	
	
	Dim SQLStr
	'SqlStr="select (select count(avgscore) from KS_MoodList  where ChannelID=" & ChannelID &" and avgscore>=a.avgscore) as 排名 from KS_MoodList a Where A.Id=" & ListId & " And a.ChannelID=" & ChannelID
	
	If DataBaseType=1 Then
	SqlStr="select (select ISNULL(sum(1),0) + 1 from KS_MoodList where avgscore > a.avgscore) as rank from KS_MoodList a Where a.ID=" & ListID & " And a.ChannelID=" & ChannelID
	Else
	SqlStr="select (select iif(isnull(sum(1)), 1, sum(1) + 1) from KS_MoodList where avgscore > a.avgscore) as rank from KS_MoodList a Where a.ID=" & ListID & " And a.ChannelID=" & ChannelID
	End If
	
	Dim PMrs:Set PMrs=conn.Execute(SqlStr)
	If Not PMrs.Eof Then
	 PM=PMRS(0)
	Else
	 PM=Conn.Execute("Select count(1) From " & KS.C_S(ChannelID,2))(0)
	End If
	PMRs.Close : Set PMRs=Nothing
	
    VoteArr=Split(ProjectContent,"$$$")
	
	VoteStr=VoteStr & "<table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0""><tr><td style=""font-weight:bold;color:#2E6493;width:170px;word-wrap:break-word; "">" & KS.CheckXSS(KS.S("Title")) & "</td><td style=""text-align:center;""><div style=""line-height:42px;font-size:18px;color:#fff;font-weight:bold;height:42px;width:52px;background:url(" & KS.GetDomain & "images/rfbg.gif) no-repeat;"">" & AvgScore &"</div></td></tr><tr>"
	VoteStr=VoteStr & "<td colspan=""2"" style=""padding-top:5px;padding-bottom:5px;border-top:1px dashed #2E6493;border-bottom:1px dashed #2E6493""><table border=""0"" cellspacing=""0"" cellpadding=""0""><tr>"
	For I=0 To Ubound(VoteArr)
	 If trim(VoteArr(I))<>"" Then
	   VoteItemArr=Split(VoteArr(i),"|")
	   If Ubound(VoteItemArr)>=1 And VoteItemArr(0)<>"" And VoteItemArr(1)<>"" Then
			ImgSrc=VoteItemArr(1)
			If Left(Lcase(ImgSrc),4)<>"http" Then ImgSrc=KS.Setting(2) & ImgSrc
		 VoteStr=VoteStr & "<td valign=""bottom""  style=""text-align:center;""><div style=""height:89px;padding-top:4px;width:65px;background:url(" & KS.GetDomain & "images/rbg.gif) no-repeat;font-weight:bold;color:#2E6493;""><img alt=""" & VoteItemArr(0) & """ src=""" & ImgSrc & """ border=""0""><br>" & VoteItemArr(0) & "<br>" & ItemVoteNum(I) & " 人</div></td>"
	   End If
	 End If
	Next
	VoteStr=VoteStr & "</tr></table></td>"
	VoteStr=VoteStr & "</tr>"
	VoteStr=VoteStr & "<tr><td style=""height:35px;color:#2E6493;font-weight:bold"">总评分:</td><td nowrap>" & AvgScore &" 分</td></tr>"
	VoteStr=VoteStr & "<tr><td style=""font-weight:bold;color:#2E6493;padding-bottom:5px;"">总排名:</td><td>" & PM & " 名</td></tr>"
	VoteStr=VoteStr & "<tr><td height=""30"" colspan=""2"" style=""border-top:1px dashed #2E6493;text-align:right""> <img src=""" & KS.GetDomain & "images/lx.gif"" align=""absmiddle""/> <a href=""javascript:void(0)"" onclick=""PopRating()"">>> 我也要参与评分</a></td></tr>"
	
	
	VoteStr=VoteStr & "</table>"
    response.write "var data={'str':'" & VoteStr & "'};"
	response.end
End Function

'弹出选项
Sub ShowPopup()
    LoginTF=KSUser.UserLoginChecked()
    If LoginTF=false And OnlyUser=1 Then
	 KS.Die "var popu={'islogin':'false','str':'" & VoteStr & "'};"
	Else
     Dim Title
     Dim RS:Set RS=Conn.Execute("Select top 1 Title From KS_ItemInfo Where ChannelID=" & ChannelID & " And InfoID=" & InfoID)
	 If Not RS.Eof Then
	   Title=RS(0)
	 End If
	 RS.Close : Set RS=Nothing
	 VoteArr=Split(ProjectContent,"$$$")
	 VoteStr=VoteStr & "<table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"">"
	 If Not KS.IsNul(Title) Then
	 
	 End If
	 VoteStr=VoteStr & "<tr><td style=""font-weight:bold;height:35px"">总共10分,您认为此" & KS.C_S(ChannelID,3) & "可以得几分?</td></tr>"
	 VoteStr=VoteStr & "<tr><td><select id=""myscore""><option value=""0"">0</option><option value=""0.5"">0.5</option><option value=""1"">1</option><option value=""1.5"">1.5</option><option value=""2"">2</option><option value=""2.5"">2.5</option><option value=""3"">3</option><option value=""3.5"">3.5</option><option value=""4"">4</option><option value=""4.5"">4.5</option><option value=""5"" selected>5</option><option value=""5.5"">5.5</option><option value=""6"">6</option><option value=""6.5"">6.5</option><option value=""7"">7</option><option value=""7.5"">7.5</option><option value=""8"">8</option><option value=""8.5"">8.5</option><option value=""9"">9</option><option value=""9.5"">9.5</option><option value=""10"">10</option></select> 分</td></tr>"
	 VoteStr=VoteStr & "<tr><td  style=""font-weight:bold;height:35px"">您对此" & KS.C_S(ChannelID,3) & "的评价?</td></tr>"
	 For I=0 To Ubound(VoteArr)
		 If trim(VoteArr(I))<>"" Then
		   VoteItemArr=Split(VoteArr(i),"|")
		   If Ubound(VoteItemArr)>=1 And VoteItemArr(0)<>"" And VoteItemArr(1)<>"" Then
		      If I=0 Then
			    VoteStr=VoteStr & "<tr><td><label><input type=""radio"" name=""myitem"" value=""" & I & """ checked>" & VoteItemArr(0) & "</label></td></tr>"
			  Else
			    VoteStr=VoteStr & "<tr><td><label><input type=""radio"" name=""myitem"" value=""" & I & """>" & VoteItemArr(0) & "</label></td></tr>"
			  End If
		   End If
		 End If
	Next
	 VoteStr=VoteStr & "<tr><td  style=""text-align:center;font-weight:bold;height:35px""><input type=""button"" style=""padding:2px;border:1px solid #999999;background:#fff;color:green"" onclick=""PostMyScore()"" value=""提交我的评分""/> <input type=""button"" value=""取消关闭"" onclick=""closeWindow()"" style=""padding:2px;border:1px solid #999999;background:#fff;color:green""/></td></tr>"
	VoteStr=VoteStr & "</table>"
  End If
    response.write "var popu={'islogin':'true','str':'" & VoteStr & "'};"
	response.end
End Sub



Call CloseConn()
Set KS=Nothing
Set KSUser=Nothing
%>
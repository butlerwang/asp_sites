<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<%

Dim KS
Set KS=New PublicCls
Dim ChannelID,ID,Hits,RS,SqlStr,HitsByDay,HitsByWeek,HitsByMonth,Action

ChannelID=KS.ChkClng(KS.S("M"))
ID = KS.ChkClng(KS.S("ID"))
Action=KS.G("Action")
If Action="commentnum" Then  '评论数
 Dim PostID:PostID=KS.ChkClng(KS.S("PostID"))
 If PostID=0 OR KS.ChkClng(KS.C_S(Channelid,6))<>1 Then
  KS.Die "document.write('<a href=""" & KS.GetDomain & "plus/Comment.asp?ChannelID=" & channelid & "&InfoID=" & id & """ target=""_blank"">我要评论(<span>" & conn.execute("select count(id) as num from ks_comment where channelid=" & channelid & " and infoid=" & id)(0) & "</span>)</a>');"
 Else
  Set RS=Conn.Execute("Select top 1 PostTable From " & KS.C_S(ChannelID,2) & " Where PostID=" & PostID & " and ID=" & ID)
  If Not RS.Eof Then
   Dim PostTable:PostTable=RS(0)
   RS.Close:Set RS=Nothing
   KS.Die "document.write('<a href=""" &KS.GetClubShowUrl(PostId) & """ target=""_blank"">参与跟帖(<span>" & Conn.Execute("Select Count(ID) From "& PostTable & " Where Verific=1 And parentid<>0 and TopicId=" & PostID)(0) & "</span>)</a>');"
  End If
 End If
 KS.Die ""
End If


 If ID = 0 Or ChannelID=0 Then
        Hits = 0
 Else
       Set RS = Server.CreateObject("ADODB.Recordset")
        SqlStr = "SELECT top 1 Hits,HitsByDay,HitsByWeek,HitsByMonth,LastHitsTime FROM [" & KS.C_S(ChannelID,2) & "] Where ID=" & ID
	   If KS.ChkClng(KS.C_S(ChannelID,6))=3 Then
			RS.Open SqlStr, conn, 1,1
			If RS.bof And RS.EOF Then
				Hits = 0
			Else
				Hits=rs(0)
				HitsByDay=rs(1)
				HitsByWeek=rs(2)
				HitsByMonth=rs(3)
			End If
       Else
			RS.Open SqlStr, conn, 1, 3
			If RS.bof And RS.EOF Then
				Hits = 0
			Else
				IF Action="Count" Then
				 rs(0) = rs(0) + 1
				 If KS.ChkClng(DateDiff("Ww", rs(4), Now())) <= 0 Then
					rs(2) = rs(2) + 1
				 Else
					rs(2) = 1
				 End If
				 If DateDiff("M", rs(4), Now()) <= 0 Then
					rs(3) = rs(3) + 1
				 Else
					rs(3) = 1
				 End If
				 If DateDiff("D", rs(4), Now()) <= 0 Then
					rs(1) = rs(1) + 1
				 Else
					rs(1) = 1
					rs(4) = Now()
				 End If
				rs.Update
				Conn.Execute("Update [KS_ItemInfo] Set Hits=" & RS(0) & ",HitsByDay=" & RS(1) & ",HitsByWeek=" & RS(2) & ",HitsByMonth=" & RS(3) & ",LastHitsTime=" & SQLNowString&" Where channelid=" & ChannelID & " and InfoID=" & ID)
			   End IF
				Hits=rs(0)
				HitsByDay=rs(1)
				HitsByWeek=rs(2)
				HitsByMonth=rs(3)
        End If

	 End If
	 rs.Close:Set rs = Nothing 
End If

	Select Case  KS.ChkClng(KS.S("GetFlag"))
	 Case 0
	  Response.Write "document.write('" & Hits & "');"
	 Case 1
	  Response.Write "document.write('" & HitsByDay & "');"
	 Case 2
	  Response.Write "document.write('" & HitsByWeek & "');"
	 Case 3
	  Response.Write "document.write('" & HitsByMonth & "');"
	End Select


Call CloseConn()
Set KS=Nothing
%> 

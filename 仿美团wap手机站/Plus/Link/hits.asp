<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<%

Dim KSCls
Set KSCls = New HitsCls
KSCls.Kesion()
Set KSCls = Nothing

Class HitsCls
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		Dim LinkID, ObjRS,Url,sitename
		LinkID = KS.ChkClng(Request("linkid"))
		If LinkID=0 Then KS.Die ""
		Set ObjRS = Server.CreateObject("Adodb.RecordSet")
		ObjRS.Open "Select top 1 Url,hits,sitename From KS_Link Where LinkID=" & LinkID, Conn, 1, 3
		If Not ObjRS.EOF Then
		  ObjRS(1) = ObjRS(1) + 1
		  ObjRS.Update
		  sitename=ObjRS(2)
		  URL=ObjRS(0)
		 
		 '========点友情链接加积分==================
		 if KS.Setting(168)="1" And KS.ChkClng(KS.Setting(169))>0 Then
		   If KS.C("UserName")<>"" Then
			  If Conn.Execute("Select top 1 * From KS_LogScore Where UserName='" & KS.C("UserName") & "' and year(adddate)=year(" & SQLNowString  &") and month(adddate)=month(" & SQLNowString &") and day(adddate)=day(" & SQLNowString & ") and channelid=1001 and infoid=" & LinkID).Eof Then
			  	 '判断有没有到达每天增加的总限
				 Dim TodayScore:TodayScore=0
				 If KS.ChkClng(KS.Setting(165))<>0 Then
				  TodayScore=KS.ChkClng(Conn.Execute("select sum(Score) from ks_logscore where InOrOutFlag=1 and year(adddate)=year(" & SQLNowString & ") and month(adddate)=month(" & SQLNowString & ") and day(adddate)=day(" & SQLNowString & ") and username='" & ks.c("UserName") & "'")(0))
				 End If
                 If TodayScore+KS.ChkClng(KS.Setting(169))<KS.ChkClng(KS.Setting(165)) Then

                      Conn.Execute("Update KS_User Set Score=Score+" & KS.ChkClng(KS.Setting(169)) & " Where UserName='" & KS.C("UserName") & "'")
					  'on error resume next
					  Dim CurrScore:CurrScore=Conn.Execute("Select top 1 Score From KS_User Where UserName='" & KS.C("UserName") & "'")(0)
					  Conn.Execute("Insert into KS_LogScore(UserName,InOrOutFlag,Score,CurrScore,[User],Descript,Adddate,IP,Channelid,InfoID) values('" & KS.C("UserName") & "',1," & KS.ChkClng(KS.Setting(169)) & ","&CurrScore & ",'系统','点击友情链接[" & sitename & "(" & url & ")]所得!'," & SqlNowString & ",'" & replace(ks.getip,"'","""") & "',1001," & LinkID & ")")

				   
				 End If

			  End If
			  
		   End If
		 End If
		'=====================================
		End If
		  ObjRS.Close
		  Set ObjRS = Nothing
		End Sub

End Class
%>

 

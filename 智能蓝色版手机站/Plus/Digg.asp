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
Dim ChannelID:ChannelID=KS.ChkCLng(KS.S("Channelid"))
Dim InfoID:InfoID=KS.ChkCLng(KS.S("InfoID"))
Dim DigType:DigType=KS.ChkClng(KS.S("DigType"))
Dim UserName,LoginTF,PrintOut,DiggNum,CDiggNum,DigXml
PrintOut=KS.S("PrintOut")
If ChannelID=0 or InfoID=0 Then KS.ECHO "err":Response.End()


Set KSUser = New UserCls

if Action="hits" Then
  LoginTF=KSUser.UserLoginChecked()
  UserName=KSUser.UserName
  If LoginTF=false and KS.C_S(ChannelID,37)="0" Then
   If PrintOut="js" Then
    KS.ECHO "alert('您还没有登录,不能推荐!');"
   Else
    KS.ECHO "nologin"
   End If
   Response.End
  Else
   If UserName="" Then UserName="游客"
   Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
   RS.Open "Select top 1 * From KS_DiggList Where ChannelID=" & ChannelID & " And InfoID=" & InfoID,conn,1,3
   If RS.Eof Then
     RS.AddNew
	  RS("ChannelID")=ChannelID
	  RS("InfoID")=InfoID
	  RS("LastDiggTime")=Now()
	  RS("LastDiggUser")=UserName
	  RS("DiggNum")=0
	  RS("CDiggNum")=0
	 RS.Update
   End IF
   RS.Close
   Dim DiggID:DiggID=Conn.Execute("Select DiggID From KS_DiggList Where ChannelID=" & ChannelID & "And InfoID=" & InfoID)(0)
   RS.Open "Select top 1 * From KS_Digg Where ChannelID=" & ChannelID &" And InfoID=" & InfoID & " And UserIP='" & KS.GetIP() & "'",Conn,1,3
   If Not RS.Eof Then
     If (KS.ChkClng(KS.C_S(ChannelID,39))=0 or (RS("UserIP")=KS.GetIP() and KS.ChkClng(KS.C_S(ChannelID,38))=1)) Then
		If PrintOut="js" Then
		 KS.ECHO ("alert('你已投票过了！');")
		Else
		 KS.ECHO ("over")
		End If
		Response.End()
	  End If
   End If
    RS.AddNew
	 RS("ChannelID")=ChannelID
	 RS("InfoID")=InfoID
	 RS("UserName")=KSUser.UserName
	 RS("UserIP")=KS.GetIP()
	 RS("DiggID")=DiggID
	 RS("DiggTime")=Now
	 RS("DiggType")=DigType
	 RS.Update
	 If DigType=0 Then
	 Conn.Execute("Update KS_DiggList set LastDiggTime=" & SqlNowString & ",DiggNum=DiggNum+" & KS.ChkClng(KS.C_S(ChannelID,40)) &" Where ChannelID=" & ChannelID & " And InfoID="& InfoID)
	 Else
	 Conn.Execute("Update KS_DiggList set LastDiggTime=" & SqlNowString & ",CDiggNum=CDiggNum+" & KS.ChkClng(KS.C_S(ChannelID,40)) &" Where ChannelID=" & ChannelID & " And InfoID="& InfoID)
	 End If
	 RS.Close: Set RS=Nothing
  End If
End If

Show()

Sub Show()
   With KS
     Dim RS,DiggNum,CDiggNum
	 Set RS=Conn.Execute("Select Top 1 DiggNum,CDiggNum From KS_DiggList Where ChannelID=" & ChannelID & " And InfoID=" & InfoID)
     If Not RS.Eof Then
	  Set DigXml=.RsToXml(RS,"row","digroot")
	 End If
	 RS.Close : Set RS=Nothing
	 If IsObject(DigXml) Then
	  DiggNum=DigXml.DocumentElement.SelectSingleNode("//row/@diggnum").text
	  CDiggNum=DigXml.DocumentElement.SelectSingleNode("//row/@cdiggnum").text
	  If PrintOut="js" Then
	   .echo "$('s" & infoid & "').innerHTML='" & DiggNum & "';" & vbcrlf
	   .echo "try{" & vbcrlf
	   .echo "   $('c" & infoid & "').innerHTML=" & CDiggNum & ";" & vbcrlf
	   .echo "   var znum=$('s" & infoid & "').innerHTML;" & vbcrlf
	   .echo " 	 var cnum=$('c" & infoid & "').innerHTML;" & vbcrlf
	   .echo "	 var totalnum=parseInt(znum)+parseInt(cnum);" & vbcrlf
	   .echo " 	$('perz" & infoid & "').innerHTML=((znum*100)/totalnum).toFixed(2)+'%';" & vbcrlf
	   .echo "	$('perc" & infoid & "').innerHTML=((cnum*100)/totalnum).toFixed(2)+'%';" & vbcrlf
	   .echo "	$('digzcimg').style.width = parseInt((znum/totalnum)*55);" & vbcrlf
	   .echo "	$('digcimg').style.width = parseInt((cnum/totalnum)*55);" & vbcrlf
	   .echo "}catch(e){" & vbcrlf
	   .echo "}" & vbcrlf
	  Else
       .echo infoid & "|" & DiggNum &"|" & CDiggNum
	  End If
	 Else
	   .echo infoid & "|0|0"
	 End If
	 Set DigXml=Nothing
  End With
End Sub


Call CloseConn()
Set KS=Nothing
Set KSUser=Nothing
%>

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
  Dim PrintOut:PrintOut=KS.S("PrintOut")
  
          If PrintOut="js" Then
			Response.Write "MoodPositionBack('" & Vote() & "');"
		  Else
	        Response.Write Vote()
		  End iF
		  
  Response.End()
Else
  Response.Write ReplaceJsBr(GetVoteStr())
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
	    Dim RSI:Set RSI=Server.CreateObject("ADODB.RECORDSET")
		RSI.Open "Select * From KS_MoodList Where MoodID=" & MoodID & " and ChannelID="& ChannelID &" And InfoID=" & InfoID,Conn,1,3
		If RSI.Eof And RSI.Bof Then
		 RSI.AddNew
		 RSI("MoodID")=MoodID
		End If
		 On Error Resume Next
		 RSI("Title")=Conn.Execute("Select Title From " & KS.C_S(ChannelID,2) & " Where ID=" & InfoID)(0)
		 If Err Then Vote= "noinfo":RSI.Close:Set RSI=Nothing:Exit Function
		 RSI("ChannelID")=ChannelID
		 RSI("InfoID")=InfoID
		 For I=0 To 14 
		  If Trim(I)=KS.S("itemid") Then
		  RSI("M" & i)=RSI("M"&i)+1
		  End If
		 Next
		 RSI.Update
		RSI.Close:Set RSI=Nothing
		Response.Cookies(KS.SiteSn)("mood_"&ChannelID&"_"&InfoID)="ok"
	    Vote= GetVoteStr()
	End If
  End If

End Function

Function GetVoteStr()
     Dim RST,TotalVote,PerVote,Percentage,ImgSrc
	Dim ItemVoteNum(15)
	TotalVote=0
	Set RST=Server.CreateObject("ADODB.RECORDSET")
	RST.Open "Select top 1 * From KS_MoodList Where ChannelID=" & ChannelID & " And InfoID=" & InfoID & " And MoodID=" & MoodID,conn,1,1
	If Not RST.Eof Then
	 For I=0 To 14
	  TotalVote=TotalVote+RST("M" & I)
	  ItemVoteNum(i)=RST("M" & I)
	 Next
	End If
	RST.Close:Set RST=Nothing
	
    VoteArr=Split(ProjectContent,"$$$")
	VoteStr="<div id=""xinqing"">"
	VoteStr=VoteStr & "<table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0""><tr><td class=""mood_top"" colspan=""15"">您看到此" & KS.C_S(ChannelID,4) & KS.C_S(ChannelID,3) & "时的感受<span>（已有 <b>" & TotalVote & "</b> 人表态）</span></td></tr><tr>"
	For I=0 To Ubound(VoteArr)
	 If trim(VoteArr(I))<>"" Then
	   VoteItemArr=Split(VoteArr(i),"|")
	   If Ubound(VoteItemArr)>=1 And VoteItemArr(0)<>"" And VoteItemArr(1)<>"" Then
	        If TotalVote<>0 Then PerVote=Round(ItemVoteNum(I)/TotalVote,4)
		    Percentage=PerVote*100
			if Percentage<1 and Percentage<>0 then	Percentage= "0" & Percentage
			ImgSrc=VoteItemArr(1)
			If Left(Lcase(ImgSrc),4)<>"http" Then ImgSrc=KS.Setting(2) & ImgSrc
		 VoteStr=VoteStr & "<td valign=""bottom""  style=""text-align:center"">" & ItemVoteNum(I) & "<br><img alt=""" & Percentage & "%"" src=""" & KS.Setting(2) & KS.Setting(3) & "images/default/post.gif"" height=""" & PerVote*50 & """ width=""20""><br><img alt=""" & VoteItemArr(0) & """ src=""" & ImgSrc & """ border=""0""><br>" & VoteItemArr(0) & "<br><input type=""radio"" onclick=""MoodPosition(" & I &");"" name=""votebutton""></td>"
	   End If
	 End If
	Next
	VoteStr=VoteStr & "</tr>"
	VoteStr=VoteStr & "</table>"
	VoteStr=VoteStr &"</div>"
    GetVoteStr=VoteStr
End Function

Function ReplaceJsBr(Content)
		 Dim i
		 Content=Replace(Content,"""","\""")
		 Dim JsArr:JSArr=Split(Content,Chr(13) & Chr(10))
		 For I=0 To Ubound(JsArr)
		   ReplaceJsBr=ReplaceJsBr & "document.writeln('" & JsArr(I) &"');" & vbcrlf 
		 Next
End Function
%>
CreateMoodAjax=function(){
	if(window.XMLHttpRequest){
		return new XMLHttpRequest();
	} else if(window.ActiveXObject){
		return new ActiveXObject("Microsoft.XMLHTTP");
	} 
	throw new Error("XMLHttp object could be created.");
}
ajaxReadText=function(file,fun){
	var xmlObj = CreateMoodAjax();
	
	xmlObj.onreadystatechange = function(){
		if(xmlObj.readyState == 4){
			if (xmlObj.status ==200){
				obj = xmlObj.responseText;
				eval(fun);
			}
			else{
				alert("读取文件出错,错误号为 [" + xmlObj.status  + "]");
			}
		}
	}
	try{
	xmlObj.open ('GET', file, true);
	xmlObj.send (null);
	}
	catch(e){
		var head = document.getElementsByTagName("head")[0];        
		var js = document.createElement("script"); 
		js.src = file+"&printout=js"; 
		head.appendChild(js);   
	}
}
MoodPosition=function(itemid){
ajaxReadText('<%=KS.Setting(2)%>/plus/Mood.asp?action=hits&itemid='+itemid+'&id=<%=MoodID%>&m_ID=<%=ChannelID%>&c_id=<%=InfoID%>','MoodPositionBack(obj)');
}
MoodPositionBack=function(obj){
 switch(obj){
  case "nologin":
   alert('对不起,您还没登录不能表态!');
   break;
  case "standoff":
   alert('您已表态过了, 不能重复表态!');
   break;
  case "lock":
   alert('心情指数已关闭!');
   break;
  case "errstartdate":
   alert('未到表态时间!');
   break;
  case "errexpireddate":
   alert('表态时间已过!');
   break;
  case "errgroupid":
   alert('您没有表态的权限!');
   break;
  case "noinfo":
   alert('找不到您要表态的信息!');
   break;
  default:
   alert('恭喜,您已成功表态!');
   document.getElementById('xinqing').innerHTML=obj;
   break;
 }
}
<%
Call CloseConn()
Set KS=Nothing
Set KSUser=Nothing
%>
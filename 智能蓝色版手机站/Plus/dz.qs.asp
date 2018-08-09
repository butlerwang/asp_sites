<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../Plus/md5.asp"-->
<%

response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Response.CharSet="utf-8"
Dim KS:Set KS=New PublicCls
Set KSUser = New UserCls

Dim LoginTF:LoginTF=KSUser.UserLoginChecked
Dim ChannelID,InfoID,RS,CommentStr,UserIP,Total,TitleStr,N,DomainStr
Dim totalPut, Page, MaxPerPage,PageNum,SqlStr
ChannelID=KS.Chkclng(KS.S("ChannelID"))
IF ChannelID=0 Then KS.Die ""

InfoID=KS.ChkClng(KS.S("InfoID"))
DomainStr=KS.GetDomain
Select Case KS.S("Action")
 Case "savesign"  Call savesign()
 Case Else  Call LoadMain()
 End Select
 Set KS=Nothing
 Set KSUser=Nothing
 
 
 Sub LoadMain
	 Dim RS,SQLStr,Title,Url,SignUser,SignDateLimit,SignDateEnd,xml,node,HasSignUser,NoSignUser,I,MustSignUserArr
	 Set RS=Server.CreateObject("adodb.recordset")
	 SQLStr="Select top 1 ID,Tid,Fname,Title,IsSign,SignUser,SignDateLimit,SignDateEnd From " & KS.C_S(ChannelID,2) & " Where issign=1 and ID=" & InfoID
	 RS.Open SQLStr,conn,1,1
	 If RS.Eof And RS.Bof Then 
	  RS.CLose
	  Set RS=Nothing
	  Exit Sub
	 End If
	 Url=KS.GetItemUrl(ChannelID,RS(1),rs(0),rs(2))
	 Title=RS(3)
	 SignUser=RS("SignUser")
	 SignDateLimit=RS("SignDateLimit")
	 SignDateEnd=RS("SignDateEnd")
	 RS.Close
	 RS.Open "Select top 500 username From KS_ItemSign Where ChannelID=" & ChannelID & " and infoid=" & InfoID,conn,1,1
	 If Not RS.EOf Then
	   SET xml=KS.RsToXml(RS,"row","")
	   for each node in xml.documentelement.selectnodes("row")
	     if HasSignUser="" then 
		   HasSignUser=node.selectSingleNode("@username").text
		 else
		   HasSignUser=HasSignUser& "," & node.selectSingleNode("@username").text
		 end if
	   next
	 End If
	 RS.Close
	 If HasSignUser="" Then
	  NoSignUser=SignUser
	 Else
	   MustSignUserArr=Split(SignUser,",")
	   For I=0 To Ubound(MustSignUserArr)
	      If KS.FoundInArr(HasSignUser,MustSignUserArr(I),",")=false Then
		    if NoSignUser="" then
			  NoSignUser=MustSignUserArr(I)
			else
			  NoSignUser=NoSignUser & "," & MustSignUserArr(I)
			end if
		  End If
	   Next
	 End If
	 
	 if NoSignUser<>"" Then
	 KS.Echo "<div class=""title"">以下是对[<a href=""" & url & """>" & Title & "</a>]的未签收单位,共<font color=red>" & Ubound(Split(NoSignUser,","))+1 & "</font>个用户</div>" &vbcrlf
	 KS.Echo "<div class=""user_name"" style=""word-break:break-all;"">" & "<span>" & replace(NoSignUser,",","</span><span>") &"</span></div>" &vbcrlf
	 KS.Echo "<div style=""clear:both;margin-top:10px;border-bottom:1px solid #999999""></div>" &vbcrlf
	 End If
	 If KS.C("UserName")<>"" Then
	    If KS.FoundInArr(NoSignUser,KS.C("UserName"),",")=true Then
		  KS.Echo "<div class=""title"">用户“" & KS.C("UserName") &"”您还没有签收,请您"
		  If SignDateLimit="1" Then
		    KS.Echo "在[<font color=""red"">" & SignDateEnd & "</font>]完成"
		  End If
		  KS.Echo "签收:<br/><textarea name=""signcontent"" id=""signcontent"" style=""width:95%;height:70px;border:1px solid #ccc""></textarea><div style=""text-align:center""><input type=""button"" id=""btnok"" value=""确定签收"" onclick=""SignOk()""></div></div>"
		  KS.Echo "<div style=""margin-top:10px;border-bottom:1px solid #999999""></div>" &vbcrlf 
		End If
	 End If
	 
	 MaxPerPage=20
	 Page=KS.ChkClng(Request("page"))
	 If Page=0 Then Page=1
	 If Not KS.IsNul(SignUser) Then
	 RS.Open "Select * From KS_ItemSign Where ChannelID=" &ChannelID & " and infoid=" &infoid &" and username in('" & replace(SignUser,",","','")&"') order by id",conn,1,1
	 If Not RS.EOF Then
			totalPut = RS.Recordcount
			If Page < 1 Then Page = 1
			If (totalPut Mod MaxPerPage) = 0 Then
				PageNum = totalPut \ MaxPerPage
			Else
				PageNum = totalPut \ MaxPerPage + 1
			End If

			If (Page - 1) * MaxPerPage < totalPut Then
				RS.Move (Page - 1) * MaxPerPage
			Else
				Page = 1
			End If
			Set XML=KS.ArrayToxml(Rs.GetRows(MaxPerPage),Rs,"row","xml")					
			RS.Close : Set RS=Nothing
	        KS.Echo "<div class=""title"">以下是对[<a href=""" & url & """>" & Title & "</a>]的已签收单位,共<font color=red>" & totalPut & "</font>个用户</div>" &vbcrlf
			For Each Node In  Xml.DocumentElement.SelectNodes("row")
			  KS.Echo "<div class=""unite"">签收单位：" & Node.SelectSingleNode("@username").text &"<span style='color:#999999'>(签收时间:" & Node.SelectSingleNode("@adddate").text & ")</span></div>"
			  KS.Echo "<div class=""unitecontent"">" & Node.SelectSingleNode("@content").text &"</div>"
	          KS.Echo "<div style=""margin-top:10px;border-bottom:1px solid #cccccc; margin-bottom:10px;""></div>" &vbcrlf
			Next
			  KS.Echo "<div class=""qspage"" style=""text-align:right"">分<font color=red>" & PageNum & "</font>页,当前第<font color=red>" & Page & "</font>页"
			  if page>1 then
			  KS.Echo " <a href=""javascript:loading(1);"">首页</a>"
			  KS.Echo " <a href=""javascript:loading(" & page-1 & ");"">上一页</a>"
			  end if
			  
			  If page<>PageNum Then
			  KS.Echo " <a href=""javascript:loading(" & page+1 & ");"">下一页</a>"
			  KS.Echo " <a href=""javascript:loading(" & pagenum & ");"">末页</a>"
			  End If
			  KS.Echo "</div>"
	End If
   End If
End Sub

	
'保存签收
Sub savesign()	
	Dim UserName,signcontent
	If LoginTF=false Then
		   Response.Write("1|对不起,请先登录后再签收!|null")
		   Response.End
	End If
	signcontent=KS.G("signcontent")
	
    UserName=KSUser.UserName
	IF InfoID=0 or channelid=0 Then 
		   Response.Write("1|参数传递有误!|null")
		   Response.End
	End if
		
	if signcontent="" Then 
		 Response.Write("1|请填写签收内容!|null")
		 Response.End
	End if
	
   If DataBaseType=1 Then
	If conn.execute("Select top 1 * From " & KS.C_S(ChannelID,2) & " where issign=1 and ','+cast(signuser as nvarchar(4000))+',' like '%," & KSUser.UserName &",%'").eof then
		 Response.Write("1|对不起,本篇文档不需要签收!|null")
		 Response.End
	end if
  Else
	If conn.execute("Select top 1 * From " & KS.C_S(ChannelID,2) & " where issign=1 and ','+signuser+',' like '%," & KSUser.UserName &",%'").eof then
		 Response.Write("1|对不起,本篇文档不需要签收!|null")
		 Response.End
	end if
  End If

	Set RS=Server.CreateObject("ADODB.RECORDSET")
	RS.Open "select top 1 SignDateLimit,SignDateEnd From " & KS.C_S(ChannelID,2) & " where id=" & infoid,conn,1,1
	if rs.eof then
	     rs.close:set rs=nothing
		 Response.Write("1|对不起,找不到文档!|null")
		 Response.End
	else
	   if rs(0)="1" and now>rs(1) then
	     rs.close
		 set rs=nothing
		 Response.Write("1|对不起,已过签收期限!|null")
		 Response.End
	   end if
	end if
	rs.close
	
	 RS.Open "Select top 1 * From KS_ItemSign Where channelid=" & channelid &" and infoid=" & infoid &" and username='" & KSUser.UserName & "'",Conn,1,3
	 If RS.Eof Then
		RS.AddNew
		 RS("ChannelID")=ChannelID
		 RS("InfoID")=InfoID
		 RS("UserName")=KSUser.UserName
		 RS("Content")=signcontent
		 RS("AddDate")=Now
		 RS.UpDate
		 Response.Write("2|恭喜,签收完毕!|null")
	Else
		 Response.Write("1|您已签收过了!|null")
	End If
		RS.Close
		Set RS=Nothing
End Sub

%>
 

<!--#include file="Kesion.IfCls.asp"-->
<%
Sub Echo(sStr)
	 Response.Write sStr 
	 'Response.Flush()
End Sub
  
public Sub Scan(ByVal sTemplate)
	Dim iPosLast, iPosCur
	iPosLast    = 1
	Do While True 
		iPosCur    = InStr(iPosLast, sTemplate, "[#") 
		If iPosCur>0 Then
			Echo    Mid(sTemplate, iPosLast, iPosCur-iPosLast)
			iPosLast  = Parse(sTemplate, iPosCur+2)
		Else 
			Echo    Mid(sTemplate, iPosLast)
			Exit Do  
		End If 
	Loop
End Sub

 
'=========================扫描会员中心主体框架 增加于2010年6月========================================

Public Sub Kesion()
         Dim LoginTF:LoginTF=Cbool(KSUser.UserLoginChecked)
		 Dim FileContent,MainUrl,RequestItem
		 Dim KSR,ParaList
		 FCls.RefreshType = "Member"   '设置当前位置为会员中心
		 Set KSR = New Refresh
		 Dim TPath:TPath=KS.SSetting(7)  '模板地址
		 FileContent = KSR.LoadTemplate(TPath)
		  FileContent = KSR.KSLabelReplaceAll(FileContent)
		 Set KSR = Nothing
		 ScanTemplate RexHtml_IF(FileContent)
End Sub	
 
public Sub ScanTemplate(ByVal sTemplate)
	Dim iPosLast, iPosCur
	iPosLast    = 1
	Do While True 
		iPosCur    = InStr(iPosLast, sTemplate, "{#") 
		If iPosCur>0 Then
			Echo    Mid(sTemplate, iPosLast, iPosCur-iPosLast)
			iPosLast  = ParseTemplate(sTemplate, iPosCur+2)
		Else 
			Echo    Mid(sTemplate, iPosLast)
			Exit Do  
		End If 
	Loop
End Sub

Function ParseTemplate(sTemplate, iPosBegin)
		Dim iPosCur, sToken, sTemp,MyNode,CheckJS
		iPosCur      = InStr(iPosBegin, sTemplate, "}")
		sTemp        = Mid(sTemplate,iPosBegin,iPosCur-iPosBegin)
		iPosBegin    = iPosCur+1
		select case Lcase(sTemp)
			case "showusermain"  loadMain
			case "showmymenu"  ShowMyMenu
			case "userid"  echo ks.c("userid")
			case "username" echo ksuser.username
			case "groupname" echo KS.U_G(KSUser.GroupID,"groupname")
			case "userface"
			  Dim UserFaceSrc:UserFaceSrc=KSUser.GetUserInfo("UserFace")
			  if left(userfacesrc,1)<>"/" and lcase(left(userfacesrc,4))<>"http" then userfacesrc=KS.Setting(3) & userfacesrc
			  response.write userfacesrc
			case else
			  response.write ksuser.getuserinfo(sTemp)
		end select
		 ParseTemplate=iPosBegin
End Function


 
 Sub ShowMyMenu()
   If KS.SSetting(0)=1 Then  '开通空间则退出
		 If KSUser.CheckPower("s02")=true And KSUser.CheckPower("s01")=true then
			 Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			 Response.Write "<img src=""images/icon7.png"" align=""absmiddle"" /> <a href=""User_Blog.asp"">博文</a>"
			 Response.Write "<span><a href=""User_Blog.asp?Action=Add"">+发表</a></span></li>"
		 End If
		 If KSUser.CheckPower("s05")=true And KSUser.CheckPower("s01")=true Then
			 Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			 Response.Write "<img src=""images/icon2.png"" align=""absmiddle"" /> <a href=""User_Photo.asp"">相册</a>"
			 Response.Write "<span><a href=""User_Photo.asp?Action=Add"">+上传</a></span></li>"
		 End If
		 If KSUser.CheckPower("s06")=true And KSUser.CheckPower("s01")=true Then
			Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write "<img src=""images/icon19.png"" align=""absmiddle"" /> <a href=""User_Team.asp"">圈子</a>"
			Response.Write "<span><a href=""User_Team.asp?action=CreateTeam"">+创建</a></span></li>"
		 End If

	  If Conn.Execute("Select top 1 ID From ks_EnterPrise Where UserName='" &KSUser.UserName & "'").eof Then '个人空间
		 If KSUser.CheckPower("s04")=true And KSUser.CheckPower("s01")=true Then
			Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write "<img src=""images/icon20.png"" align=""absmiddle"" /> <a href=""User_music.asp"">音乐</a>"
			Response.Write "<span><a href=""User_Music.asp?action=addlink"">+添加</a></span></li>"
		 End If
		 If KSUser.CheckPower("s10")=true Then 
			If KS.C_S(5,21)="1" Then
			'Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			'Response.Write "<img src=""images/icon6.png"" align=""absmiddle"" /> <a href=""user_myshop.asp"">商品</a>"
			'Response.Write "<span><a href=""user_myshop.asp?ChannelID=5&Action=Add"">+发布</a></span></li>"
			End If
		 End IF
	  Else   '企业空间
	      if KSUser.CheckPower("s10")=true then
			'Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			'Response.Write "<img src=""images/icon6.png"" align=""absmiddle"" /> <a title='企业产品管理' href=""user_myshop.asp"">产品</a>"
			'Response.Write "<span><a href=""user_myshop.asp?ChannelID=5&Action=Add"">+发布</a></span></li>"
		  End If
		  if KSUser.CheckPower("s11")=true And KSUser.CheckPower("s01")=true then
			Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write "<img src=""images/icon21.png"" align=""absmiddle"" /> <a title='企业新闻管理' href=""user_EnterpriseNews.asp"">动态</a>"
			Response.Write "<span><a href=""user_EnterpriseNews.asp?Action=Add"" title='发布企业新闻'>+发布</a></span></li>"
		  end if
		  if KSUser.CheckPower("s12")=true And KSUser.CheckPower("s01")=true then
			Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write "<img src=""images/icon10.png"" align=""absmiddle"" /> <a title='关键词广告管理管理' href=""user_EnterpriseAD.asp"">广告</a>"
			Response.Write "<span><a href=""user_EnterpriseAD.asp?Action=Add"">+发布</a></span></li>"
		  end if
		  if KSUser.CheckPower("s13")=true And KSUser.CheckPower("s01")=true then
			  	Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
				Response.Write "<img src=""images/icon22.png"" align=""absmiddle"" /> <a title='企业荣誉证书管理' href=""user_Enterprisezs.asp"">荣誉</a>"
				Response.Write "<span><a href=""user_Enterprisezs.asp?Action=Add"">+发布</a></span></li>"
		  End If
	  End If
  End If
   		 If KSUser.CheckPower("s03")=true Then
			Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write "<img src=""images/icon18.png"" align=""absmiddle"" /> <a href=""User_friend.asp"">好友</a>"
			Response.Write "<span><a href=""User_Friend.asp?action=addF"">+寻找</a></span></li>"
		 End If

		 If KSUser.CheckPower("s07")=true Then 
			Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write "<img src=""images/icon7.png"" align=""absmiddle"" /> <a href=""User_Class.asp"">专栏</a>"
			Response.Write "<span><a href=""User_Class.asp?Action=Add"">+创建</a></span></li>"
		 End If

End Sub
 
'------扫描会员中心主体框架------
%>
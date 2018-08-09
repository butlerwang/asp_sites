<script language="jscript" runat="server">
function getjson(str){
        try{
           eval("var jsonStr = (" + str + ")");
        }catch(ex){
           var jsonStr = null;
        }
        return jsonStr;
}
</script>
<%

If GCls.ComeUrl="" Then GCls.ComeUrl=Request.ServerVariables("HTTP_REFERER")  '记录来源页


Dim API_QuickLogin,API_GroupID,API_QQEnable,API_QQAppId,API_QQAppKey,API_QQCallBack 
Dim API_AlipayEnable,API_AlipayPartner,API_AlipayKey,API_AlipayReturnurl
Dim API_SinaEnable,API_SinaId,API_SinaKey,API_SinaCallBack
Dim API_Path,API_Enable,API_ConformKey,API_Urls
Dim API_Debug,API_LoginUrl,API_ReguserUrl,API_LogoutUrl
Dim KS:Set KS=New PublicCls
API_Path = KS.Setting(3) & "API/"
LoadXslConfig()


'根据url读取远程文件的内容
'参数 url 请求处理的URL，method post或get postdata，请求发送的数据
Function file_get_contents(url,method,postdata)
 Dim objXML:Set   objXML   =   server.CreateObject( "MSXML2.ServerXMLHTTP" & MsxmlVersion)     
  if lcase(method)="post" then
	objXML.open   "POST",   url,   False     
	objXML.setRequestHeader "Content-Length", Len(postdata)
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.send(postdata) 
  else
	objXML.open   "GET",   url&"?"&postdata,   False     
	objXML.send() 
  end if
	If objXml.Readystate=4 Then
	 file_get_contents=     objXML.responSetext 
	Else
	 file_get_contents=0
	End If
 Set objXML=Nothing
End Function

'获得用户信息
'参数：apitype 接口类型 1 qq 2新浪
function get_user_info(ApiType,access_token,openid)
    dim url
	if apitype="1" then 'qq
		url = "https://graph.qq.com/user/get_user_info"
		get_user_info = file_get_contents(url,"get","access_token="&access_token&"&oauth_consumer_key="&API_QQAppId &"&openid="&openid&"&format=json")
	elseif apitype="2" then '新浪
	    url = "https://api.weibo.com/2/users/show.json"
        get_user_info = file_get_contents(url,"get","access_token="&access_token&"&uid="&openid)
	end if
end function


'登录成功后调用绑定等操作
'参数：tips 标题提示信息,username 昵称 userface 头像 sex 性别,openid 接口唯一返回的ID
Sub DoBind(tips,username,userface,sex,openid)
	Dim Template,KSR
	Set KSR = New Refresh
	Template=KS.Setting(3) & KS.Setting(90) & "Common/apibind.html"		         	'模板地址
	Template=KSR.LoadTemplate(Template)

	Template=Replace(Template,"{$Title}",tips)
	Template=Replace(Template,"{$OriginUserName}",userName)
	If KS.HasChinese(username) and KS.ChkClng(KS.Setting(175))="0" then
		Dim Ce:Set Ce=new CtoeCls
		username=Ce.CTOE(username)
		Set Ce=Nothing
		elseif isnumeric(username) then
		username="q" &username
	End If
	
	If not conn.execute("select top 1 userid from ks_user where username='" & username & "'").eof then
		username=username &  year(now) & month(now) &day(now)&hour(now)
	End If
	

	'=================================允许快速自动注册登录=============================================
	If Cbool(API_QuickLogin) =True Then
	    Dim UserLen:UserLen=Len(UserName)
		Dim UserNameMaxChar:UserNameMaxChar=Cint(KS.Setting(29))
		If UserLen<UserNameMaxChar Then
		  UserName=UserName & KS.MakeRandom(UserNameMaxChar-UserLen)
		End If
		If KS.ChkClng(Api_GroupID)=0 Then Api_GroupID=2 '默认用户组
		Dim Email:Email=""
		Dim PassWord:PassWord=KS.MakeRandomChar(15)  '生成随机密码
		Call SaveReg(1,UserName,PassWord,Api_GroupID,Sex,Email,UserFace) '保存
		Exit Sub
	End If
    '===================================================================================================
	
	Template=Replace(Template,"{$UserName}",UserName)
	Template=Replace(Template,"{$UserFace}",userface)
	Template=Replace(Template,"{$OpenID}",openid)
	Dim LoginType
	If InStr(tips,"支付宝") Then LoginType="支付宝快捷"
	If InStr(tips,"QQ") Then LoginType="QQ"
	If InStr(tips,"新浪微博") Then LoginType="新浪微博"
	Template=Replace(Template,"{$LoginType}",LoginType)
	dim sexstr
	if sex="男" or sex="" then 
		sexstr="<input type=""radio"" name=""sex"" value=""男"" checked/>男 <input type=""radio"" name=""sex"" value=""女"">女"
	else
		sexstr="<input type=""radio"" name=""sex"" value=""男"" />男 <input type=""radio"" name=""sex"" value=""女"" checked>女"
	end if
	Template=Replace(Template,"{$Sex}",sexstr)
	
	
	Dim Node,UserGroupList
	If  KS.Setting(33)="0" Then 
	 UserGroupList=""
	 Template=Replace(Template,"{$ShowGroupID}"," style='display:none;'")
	Else
		    Call KS.LoadUserGroup()
			If KS.ChkClng(Api_GroupID)=0 Then Api_GroupID=2 '默认用户组

			For Each Node In Application(KS.SiteSN&"_UserGroup").DocumentElement.SelectNodes("row[@showonreg=1 && @id!=1]")
                If KS.ChkClng(Api_GroupID)=KS.ChkClng(Node.SelectSingleNode("@id").text) Then
					UserGroupList=UserGroupList & " <label><input type=""radio""  value=""" & Node.SelectSingleNode("@id").text & """ name=""GroupID"" checked>" & Node.SelectSingleNode("@groupname").text & "</label>"
				Else
					UserGroupList=UserGroupList & " <label><input type=""radio""  value=""" & Node.SelectSingleNode("@id").text & """ name=""GroupID"">" & Node.SelectSingleNode("@groupname").text & "</label>"
				End If
	       Next
	End If
	Template=Replace(Template,"{$UserGroupList}",UserGroupList)
	Template=KSR.KSLabelReplaceAll(Template)
	Response.Write Template
	Set KSR=Nothing
End Sub

'保存注册
'参数 flag 1表示qq绑定注册 2表示新浪微博 3支付宝
Sub DoRegSave(Flag)
         Dim UserNameLimitChar:UserNameLimitChar=Cint(KS.Setting(29))
		 Dim UserNameMaxChar:UserNameMaxChar=Cint(KS.Setting(30))
		 Dim EnabledUserName:EnabledUserName=KS.Setting(31)
         UserName=KS.R(KS.S("UserName"))
		 If UserName = "" Or KS.strLength(UserName) > UserNameMaxChar Or KS.strLength(UserName) < UserNameLimitChar Then
		   	 KS.Die ("<script>alert('请输入用户名(不能大于" & UserNameMaxChar & "小于" & UserNameLimitChar & ")');history.back();</script>")
		 Elseif isnumeric(UserName) then
			 KS.Die ("<script>alert('对不起，会员名不能是纯数字！');history.back();</script>")
		 Elseif KS.HasChinese(username) and KS.ChkClng(KS.Setting(175))="0" then
		   	 KS.Die("<script>alert('对不起，系统设置用户名不能含有中文！');history.back();</script>")
         ElseIF KS.FoundInArr(EnabledUserName, UserName, "|") = True Then
		   	 KS.Die("<script>alert('您输入的用户名为系统禁止注册的用户名');history.back();</script>")
		 ElseIF InStr(UserName, "-") > 0 Or InStr(UserName, "=") > 0 Or InStr(UserName, ".") > 0 Or InStr(UserName, "%") > 0 Or InStr(UserName, Chr(32)) > 0 Or InStr(UserName, "?") > 0 Or InStr(UserName, "&") > 0 Or InStr(UserName, ";") > 0 Or InStr(UserName, ",") > 0 Or InStr(UserName, "'") > 0 Or InStr(UserName, ",") > 0 Or InStr(UserName, Chr(34)) > 0 Or InStr(UserName, Chr(9)) > 0 Or InStr(UserName, "") > 0 Or InStr(UserName, "$") > 0 Or InStr(UserName, "*") Or InStr(UserName, "|") Or InStr(UserName, """") > 0 Then
             KS.Die("<script>alert('用户名中含有非法字符');history.back();</script>")
        End If
		
		 Dim RePassWord,NoMD5_Pass,userface
		 If Session("PassWord")<>"" Then
		   PassWord=Session("PassWord")
		 Else
			 PassWord=KS.R(KS.S("PassWord"))
			 RePassWord=KS.S("RePassWord")
			 If PassWord = "" Then
				 KS.Die("<script>alert('请输入登录密码!');history.back();</script>")
			 ElseIF RePassWord="" Then
				 KS.Die("<script>alert('请输入确认密码');history.back();</script>")
			 ElseIF PassWord<>RePassWord Then
				 KS.Die("<script>alert('两次输入的密码不一致');history.back();</script>")
			 End If
		 End If
		 Dim GroupID:GroupID=KS.ChkClng(Request("GroupID"))
		 If GroupID=0 Then GroupID=2   '默认用户组
		 
		 UserFace=KS.S("UserFace")
		 Call SaveReg(2,UserName,PassWord,GroupID,KS.S("sex"),Email,UserFace)	
End Sub

Sub SaveReg(IsApi,UserName,PassWord,GroupID,Sex,Email,UserFace)		 
		 Dim RS:set rs=server.CreateObject("adodb.recordset")
		 rs.open "select top 1 * from ks_user where username='" & username & "'",conn,1,3
		 if not rs.eof then
		    rs.close:set rs=nothing
			KS.Die("<script>alert('您输入的用户名已被使用，请换个名称！');history.back();</script>")
		 end if
		 RS.AddNew
		 RS("GroupID")=GroupID
		 RS("UserName")=UserName
		 RS("PassWord")=MD5(KS.R(PassWord),16)
		 RS("Question")=""
		 RS("Answer")=""
		 RS("Email")=Email
		 RS("RealName")=""
		 RS("Sex")= Sex
		 RS("UserFace")=UserFace
		 RS("RegDate")=Now
		 RS("BeginDate")=Now '开始计算时间
		 RS("JoinDate")=Now
		 RS("LastLoginTime")=Now
		 RS("RndPassword")=PassWord
		 RS("LoginTimes")=1
		 RS("PostNum")=0
		 '新会员注册，更新相应的数据
		 RS("Money")=0
		 RS("Score")=0
    	 RS("Point")=0
		 RS("Locked")=0
		 RS("ChargeType")=1
		 RS("IsAPI")=IsAPI
		 RS("LastLoginIP")=KS.GetIP
		 
		 RS("QQOpenId")=ks.c("openid")
		 RS("QQToken")=ks.c("access_token")
		 RS("SinaId")=ks.c("sinaId")
		 RS("SinaToken")=ks.c("sina_access_token")
		 RS("alipayID")=session("user_id")
		 RS.Update
		 RS.MoveLast
		 Dim UserID:UserID=RS("UserID")
		 RS.Close
		 if not ks.isnul(userface) then
		   if instr(userface,"noavatar_small.gif")=0 then
			   dim localfile:localfile=KS.Setting(3) & "uploadfiles/user/avatar/" & UserID & ".jpg"
			   if lcase(KS.SaveBeyondFile(localfile,userface))="true" then
				conn.execute("update ks_user set userface='" & localfile & "' where userid=" & userid)
			   end if
		   end if
		 end if
		 
		 Dim NewRegUserMoney:NewRegUserMoney=KS.Setting(38) : If Not IsNumerIc(NewRegUserMoney) Then NewRegUserMoney=0
		 Dim NewRegUserScore:NewRegUserScore=KS.Setting(39) : If Not IsNumeric(NewRegUserScore) Then NewRegUserScore=0
		 Dim NewRegUserPoint:NewRegUserPoint=KS.Setting(40) : If Not IsNumeric(NewRegUserPoint) Then NewRegUserPoint=0
		 If KS.ChkClng(KS.U_G(GroupID,"chargetype"))=1 Then
		  NewRegUserPoint=KS.ChkClng(KS.U_G(GroupID,"grouppoint"))
		 End If
		 
		 Dim RSG:Set RSG=Conn.Execute("select top 1 GradeID,UserTitle from KS_AskGrade where Special=0 and typeflag=1 and ClubPostNum<=0 And score<=" & NewRegUserScore & " order by score desc,ClubPostNum Desc")
		If Not RSG.Eof Then
				 Conn.Execute("Update KS_User Set ClubGradeID=" & RSG(0) & " WHERE UserName='" & UserName & "'")
		End If
		Set RSG=Conn.Execute("select top 1 GradeID,UserTitle from KS_AskGrade where Special=0 and typeflag=0 And score<=" & NewRegUserScore & " order by score desc,gradeid Desc")
		If Not RSG.Eof Then
			 Conn.Execute("Update KS_User Set GradeID=" & RSG(0) & ",GradeTitle='" & RSG(1) & "' WHERE UserName='" & UserName & "'")
		End If


		 If NewRegUserPoint<>0 Then
		  Call KS.PointInOrOut(0,0,UserName,1,NewRegUserPoint,"系统","注册新会员,赠送!",0)
		 End If
		 IF NewRegUserScore<>0 Then
		  Call KS.ScoreInOrOut(UserName,1,NewRegUserScore,"系统","注册新会员,赠送!",0,0)
		 End If
		 If NewRegUserMoney<>0 Then 
		  Call KS.MoneyInOrOut(UserName,UserName,NewRegUserMoney,4,1,now,0,"System","注册新会员,赠送!",0,0,0)
		 End If
		 
		 '===================写入个人空间================
			if KS.SSetting(0)=1 And KS.SSetting(1)=1 then
			 RS.Open "Select top 1 * From KS_Blog Where 1=0",conn,1,3
			 RS.AddNew
			  RS("AddDate")=Now
			  RS("UserID")=UserID
			  RS("UserName")=UserName
			  RS("BlogName")=UserName & "的个人空间"
			  RS("ClassID") = KS.ChkClng(Conn.Execute("Select Top 1 ClassID From KS_BlogClass")(0))
			  If UserType=1 Then
			  RS("TemplateID")=KS.ChkClng(Conn.Execute("Select Top 1 ID From KS_BlogTemplate Where flag=4 and IsDefault='true'")(0))
			  Else
			  RS("TemplateID")=KS.ChkClng(Conn.Execute("Select Top 1 ID From KS_BlogTemplate Where flag=2 and IsDefault='true'")(0))
			  End If
			  RS("Announce")="暂无公告!"
			  RS("ContentLen")=500
			  RS("Recommend")=0
			  if KS.SSetting(2)=1 then
			  RS("Status")=0
			  else
			  RS("Status")=1
			  end if
			 RS.Update
			 RS.Close
			 Set RS=Nothing
		  End If
		  Call DoLogin(userName,MD5(KS.R(PassWord),16))
End Sub

'登录成功后调用登录函数进行登录
Sub DoLogin(userName,Password)
  Dim UserRS:Set UserRS=Server.CreateObject("Adodb.RecordSet")
	 UserRS.Open "Select top 1 * From KS_User Where UserName='" &UserName & "' And PassWord='" & PassWord & "'",Conn,1,3
	 If UserRS.Eof And UserRS.BOf Then
		 UserRS.Close:Set UserRS=Nothing
		 KS.Die "<script>alert('你输入的用户名或密码有误，请重新输入!');history.back();</script>"
	 ElseIf UserRS("Locked")=1 Then
		 KS.Die "<script>alert('您的账号已被管理员锁定，请与管理员联系!');history.back();</script>"
	 ElseIf UserRS("Locked")=2 Then
		 KS.Die "<script>alert('您的账号未审核，不能登录!');history.back();</script>"
	 ElseIf UserRS("Locked")=3 Then
		 KS.Die "<script>alert('您的账号未激活，不能登录!');history.back();</script>"
	 Else
		 GroupID=UserRS("GroupID")
		 '登录成功，更新用户相应的数据
		 Dim RndPassword:RndPassword=KS.R(KS.MakeRandomChar(20))
		
		 Dim ScoreTF:ScoreTF=False
		 If KS.ChkClng(KS.U_S(UserRS("GroupID"),8))>0 and KS.ChkClng(KS.U_S(UserRS("GroupID"),9))>0 And datediff("n",UserRS("LastLoginTime"),now)>=KS.ChkClng(KS.U_S(UserRS("GroupID"),8)) then '判断时间
				ScoreTF=true
		End if
						
		 	UserRS("LastLoginIP") = KS.GetIP
            UserRS("LastLoginTime") = Now()
            UserRS("LoginTimes") = UserRS("LoginTimes") + 1
			UserRS("RndPassWord")=RndPassWord
            UserRS.Update
			
			Dim CurrScore:CurrScore=KS.ChkClng(UserRS("Score"))
			Dim GroupID:GroupID=KS.ChkClng(UserRS("GroupID"))
			Dim PostNum:PostNum=KS.ChkClng(UserRS("PostNum"))
			Dim ClubSpecialPower:ClubSpecialPower=KS.ChkClng(UserRS("ClubSpecialPower"))
			
			'更新论坛等级
			If ClubSpecialPower=0 Then
				Dim RSG:Set RSG=Conn.Execute("select top 1 GradeID,UserTitle from KS_AskGrade where TypeFlag=1 and Special=0 and ClubPostNum<=" & PostNum & " And score<=" & CurrScore & " order by score desc,ClubPostNum Desc")
				If Not RSG.Eof Then
				 Conn.Execute("Update KS_User Set ClubGradeID=" & RSG(0) & " WHERE GroupID<>1 and UserName='" & UserName & "'")
				 End If
			End If
			 '更新问答等级
			Set RSG=Conn.Execute("select top 1 GradeID,UserTitle from KS_AskGrade where typeflag=0 and Special=0 and score<=" & CurrScore & " order by score desc,gradeid Desc")
			If Not RSG.Eof Then
				 Conn.Execute("Update KS_User Set GradeID=" & RSG(0) & ",GradeTitle='" & RSG(1) & "' WHERE UserName='" & UserName & "' and gradeid>5")
			End If
			RSG.Close:Set RSG=Nothing
			
			
			If ScoreTF then 
				Session("PopTips")=KS.U_S(GroupID,8) & "分钟后重新登录，奖励积分 +" & KS.U_S(GroupID,9) & "分！"     '用于在论坛里显示
			   Call KS.ScoreInOrOut(UserName,1,KS.ChkClng(KS.U_S(GroupID,9)),"系统",KS.ChkClng(KS.U_S(GroupID,8)) & "分钟后,重新登录奖励获得",0,0)
			End if
			
			'更新购物车的ID号
			If Not KS.IsNul(KS.C("CartID")) Then
			 Conn.Execute("Update KS_ShopPackageSelect Set UserName='" & UserName & "' where username='" & KS.C("CartID") & "'")
			 Conn.Execute("Update KS_ShoppingCart Set UserName='" & UserName & "' where username='" & KS.C("CartID") & "'")
			End If
			
			If EnabledSubDomain Then
				Response.Cookies(KS.SiteSn).domain=RootDomain					
			Else
                Response.Cookies(KS.SiteSn).path = "/"
			End If
				Response.Cookies(KS.SiteSn).Expires = Date + 365
				Response.Cookies(KS.SiteSn)("UserName") = UserName
				Response.Cookies(KS.SiteSn)("Password") = Password
				Response.Cookies(KS.SiteSN)("RndPassword")= RndPassword
				Response.Cookies(KS.SiteSn)("UserID") = UserRS("UserID")
	End if
	 UserRS.Close : Set UserRS=Nothing
	 
	 
	  Dim TUrl:Turl= "../../index.asp"
	  If Not KS.IsNul(GCls.ComeUrl) Then Turl=GCls.ComeUrl:GCls.ComeUrl=""
      Response.redirect TUrl
End Sub


Class API_Conformity
	Public AppID,Status,GetData,GetAppid
	Private XmlDoc,XmlHttp
	Private MessageCode,ArrUrls,SysKey,XmlPath
	
	Private Sub Class_Initialize()
		GetAppid = ""
		AppID = "KesionCMS"
		ArrUrls = Split(Trim(API_Urls),"|")
		Status = "1"
		SysKey = API_ConformKey
		MessageCode = ""
		XmlPath = API_Path & "api_user.xml"
		XmlPath = Server.MapPath(XmlPath)
		Set XmlDoc = KS.InitialObject("msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
		Set GetData = KS.InitialObject("Scripting.Dictionary")
		XmlDoc.ASYNC = False
		LoadXmlData()
	End Sub
	
	Private Sub Class_Terminate()
		If IsObject(XmlDoc) Then Set XmlDoc = Nothing
		If IsObject(GetData) Then Set GetData = Nothing
	End Sub

	Public Sub LoadXmlData()
		If Not XmlDoc.Load(XmlPath) Then
			XmlDoc.LoadXml "<?xml version=""1.0"" encoding=""utf-8""?><root/>"
		End If
		NodeValue "appID",AppID,1,False
	End Sub
	
	Public Sub NodeValue(Byval NodeName,Byval NodeText,Byval NodeType ,Byval blnEncode)
		Dim ChildNode,CreateCDATASection
		NodeName = Lcase(NodeName)
		If XmlDoc.documentElement.selectSingleNode(NodeName) is nothing Then
			Set ChildNode = XmlDoc.documentElement.appendChild(XmlDoc.createNode(1,NodeName,""))
		Else
			Set ChildNode = XmlDoc.documentElement.selectSingleNode(NodeName)
		End If
		If blnEncode = True Then
			NodeText = AnsiToUnicode(NodeText)
		End If
		If NodeType = 1 Then
			ChildNode.Text = ""
			Set CreateCDATASection = XmlDoc.createCDATASection(Replace(NodeText,"]]>","]]&gt;"))
			ChildNode.appendChild(createCDATASection)
		Else
			ChildNode.Text = NodeText
		End If
	End Sub

	Public Property Get XmlNode(Byval Str)
		If XmlDoc.documentElement.selectSingleNode(Str) is Nothing Then
			XmlNode = "Null"
		Else
			XmlNode = XmlDoc.documentElement.selectSingleNode(Str).text
		End If
	End Property

	Public Property Get GetXmlData()
		Dim GetXmlDoc
		GetXmlData = Null
		If GetAppid <> "" Then
			GetAppid = Lcase(GetAppid)
			If GetData.Exists(GetAppid) Then
				Set GetXmlData = GetData(GetAppid)
			End If
		End If
	End Property

	Public Sub SendHttpData()
		Dim i,GetXmlDoc,LoadAppid
		'On Error Resume Next
		Set Xmlhttp = KS.InitialObject("MSXML2.ServerXMLHTTP" & MsxmlVersion)
		Set GetXmlDoc = KS.InitialObject("msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
		For i = 0 to Ubound(ArrUrls)
			XmlHttp.Open "POST", Trim(ArrUrls(i)), false
			XmlHttp.SetRequestHeader "content-type", "text/xml"
			XmlHttp.Send XmlDoc
			'Response.Write strAnsi2Unicode(xmlhttp.responseBody)
    		'response.end
			If GetXmlDoc.load(XmlHttp.responseXML) Then
				LoadAppid = Lcase(GetXmlDoc.documentElement.selectSingleNode("appid").Text)
				GetData.add LoadAppid,GetXmlDoc
				Status = GetXmlDoc.documentElement.selectSingleNode("status").Text
				MessageCode = MessageCode & LoadAppid & "(" & Status &")：" & GetXmlDoc.documentElement.selectSingleNode("body/message").Text
				If Status = "1" Then '当发生错误时退出
					Exit For
				End If
			Else
				Status = "1"
				MessageCode = "请求数据错误！"
				Exit For
			End If
		Next
		Set GetXmlDoc = Nothing
		Set XmlHttp = Nothing
	End Sub

	Public Property Get Message()
		Message = MessageCode
	End Property
	
	Public Function SetCookie(Byval C_Syskey,Byval C_UserName,Byval C_PassWord,Byval C_SetType)
		Dim i,TempStr
		TempStr = ""
		For i = 0 to Ubound(ArrUrls)
			TempStr = TempStr & vbNewLine & "<script language=""JavaScript"" src="""&Trim(ArrUrls(i))&"?syskey="&Server.URLEncode(C_Syskey)&"&username="&Server.URLEncode(C_UserName)&"&password="&Server.URLEncode(C_PassWord)&"&savecookie="&Server.URLEncode(C_SetType)&"""></script>"
		Next
		SetCookie = TempStr
	End Function

	Public Sub PrintGetXmlData()
		Response.Clear
		Response.ContentType = "text/xml"
		Response.CharSet="utf-8"
		Response.Expires = 0
		Response.Write "<?xml version=""1.0"" encoding=""utf-8""?>"&vbNewLine
		Response.Write GetXmlData.documentElement.XML
	End Sub

	Private Function AnsiToUnicode(ByVal str)
		Dim i, j, c, i1, i2, u, fs, f, p
		AnsiToUnicode = ""
		p = ""
		For i = 1 To Len(str)
			c = Mid(str, i, 1)
			j = AscW(c)
			If j < 0 Then
				j = j + 65536
			End If
			If j >= 0 And j <= 128 Then
				If p = "c" Then
					AnsiToUnicode = " " & AnsiToUnicode
					p = "e"
				End If
				AnsiToUnicode = AnsiToUnicode & c
			Else
				If p = "e" Then
					AnsiToUnicode = AnsiToUnicode & " "
					p = "c"
				End If
				AnsiToUnicode = AnsiToUnicode & ("&#" & j & ";")
			End If
		Next
	End Function

	Private Function strAnsi2Unicode(asContents)
		Dim len1,i,varchar,varasc
		strAnsi2Unicode = ""
		len1=LenB(asContents)
		If len1=0 Then Exit Function
		  For i=1 to len1
			varchar=MidB(asContents,i,1)
			varasc=AscB(varchar)
			If varasc > 127  Then
				If MidB(asContents,i+1,1)<>"" Then
					strAnsi2Unicode = strAnsi2Unicode & chr(ascw(midb(asContents,i+1,1) & varchar))
				End If
				i=i+1
			 Else
				strAnsi2Unicode = strAnsi2Unicode & Chr(varasc)
			 End If	
		Next
	End Function
End Class

Sub LoadXslConfig()
	Dim XslDoc,XslNode,Xsl_Files
	Xsl_Files = API_Path & "api.config"
	Xsl_Files = Server.MapPath(Xsl_Files)
	Set XslDoc = KS.InitialObject("Msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
	If Not XslDoc.Load(Xsl_Files) Then
		Response.Write "初始数据不存在,请检查是否存在api/api.config文件！"
		Response.End
	Else
		Set XslNode = XslDoc.documentElement.selectSingleNode("rs:data/z:row")
		API_Enable		= XslNode.getAttribute("api_enable")
		API_ConformKey		= XslNode.getAttribute("api_conformkey")
		API_Urls		= XslNode.getAttribute("api_urls")
		API_Debug		= XslNode.getAttribute("api_debug")
		API_LoginUrl		= XslNode.getAttribute("api_loginurl")
		API_ReguserUrl		= XslNode.getAttribute("api_reguserurl")
		API_LogoutUrl		= XslNode.getAttribute("api_logouturl")
		
		API_QuickLogin      = XslNode.getAttribute("api_quicklogin")
		API_GroupID         = XslNode.getAttribute("api_groupid")
		
		API_QQEnable        = XslNode.getAttribute("api_qqenable")
		API_QQAppId         = XslNode.getAttribute("api_qqappid")
		API_QQAppKey        = XslNode.getAttribute("api_qqappkey")
		API_QQCallBack      = XslNode.getAttribute("api_qqcallback")
		
		API_AlipayEnable    = XslNode.getAttribute("api_alipayenable")
		API_AlipayPartner   = XslNode.getAttribute("api_alipaypartner")
		API_AlipayKey       = XslNode.getAttribute("api_alipaykey")
		API_AlipayReturnurl = XslNode.getAttribute("api_alipayreturnurl")
		
		API_SinaEnable      = XslNode.getAttribute("api_sinaenable")
		API_SinaId          = XslNode.getAttribute("api_sinaid")
		API_SinaKey         = XslNode.getAttribute("api_sinakey")
		API_SinaCallBack    = XslNode.getAttribute("api_sinacallback")

		Set XslNode = Nothing
	End If
	Set XslDoc = Nothing
End Sub
%>
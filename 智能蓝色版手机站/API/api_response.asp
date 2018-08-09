<!--#include file="../conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../Plus/md5.asp"-->
<!--#include file="Cls_Api.asp"-->
<%

'-- 声明：本程序修改自动网论坛系统Api接口
'=========================================================
Dim XMLDom,XmlDoc,Node,Status,Messenge
Dim UserName,Act,appid
Status = 1
Messenge = ""

If Request.QueryString<>"" And API_Enable Then
	SaveUserCookie()
Else
	Set XmlDoc = KS.InitialObject("msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
	XmlDoc.ASYNC = False
	If API_Enable Then
		If Not XmlDoc.LOAD(Request) Then
			Status = 1
			Messenge = "数据非法，操作中止！"
			appid = "未知"
		Else
			If CheckPost() Then
				Select Case Act
					Case "checkname"
						Checkname()
					Case "reguser"
						UserReguser()
					Case "login"
						UesrLogin()
					Case "logout"
						LogoutUser()
					Case "update"
						UpdateUser()
					Case "delete"
						Deleteuser()
					Case "lock"
						Lockuser()
					Case "getinfo"
						GetUserinfo()
				End Select
			End If
		End If
	Else
		Status = 0
		Messenge = "API接口关闭，操作中止！"
		appid = "KesionCMS"
	End If
	ReponseData()
	Set XmlDoc = Nothing
End If

Sub ReponseData()
	If Act <> "getinfo" Then
		XmlDoc.loadxml "<root><appid>dvbbs</appid><status>0</status><body><message/></body></root>"
	End If
	XmlDoc.documentElement.selectSingleNode("appid").text = "KesionCMS"
	If API_Debug And Act <> "reguser" Then
		XmlDoc.documentElement.selectSingleNode("status").text = 0
		Messenge = ""
	Else
		XmlDoc.documentElement.selectSingleNode("status").text = status
	End If
	XmlDoc.documentElement.selectSingleNode("body/message").text = ""
	Set Node = XmlDoc.createCDATASection(Replace(Messenge,"]]>","]]&gt;"))
	XmlDoc.documentElement.selectSingleNode("body/message").appendChild(Node)
	Response.Clear
	Response.ContentType="text/xml"
	Response.CharSet="utf-8"
	Response.Write "<?xml version=""1.0"" encoding=""utf-8""?>"&vbNewLine
	Response.Write XmlDoc.documentElement.XML
End Sub

Function CheckPost()
	CheckPost = False
	Dim Syskey
	If XmlDoc.documentElement.selectSingleNode("action") is Nothing or XmlDoc.documentElement.selectSingleNode("syskey") is Nothing or XmlDoc.documentElement.selectSingleNode("username")  is Nothing Then
		Status = 1
		Messenge = Messenge & "<li>非法请求。</li>"
		Exit Function
	End If
	UserName = KS.R(XmlDoc.documentElement.selectSingleNode("username").text)
	Syskey = XmlDoc.documentElement.selectSingleNode("syskey").text
	Act = XmlDoc.documentElement.selectSingleNode("action").text
	Appid = XmlDoc.documentElement.selectSingleNode("appid").text
	
	Dim NewMd5,OldMd5
	NewMd5 = Md5(UserName & API_ConformKey,16)
	Md5OLD = 1
	OldMd5 = Md5(UserName & API_ConformKey,16)
	Md5OLD = 0

	If Syskey=NewMd5 or Syskey=OldMd5 Then
		CheckPost = True
	Else
		Status = 1
		Messenge = Messenge & "<li>请求数据验证不通过，请与管理员联系。</li>"
	End If
End Function

Sub GetUserinfo()
	Dim Rs,Sql
	XmlDoc.loadxml "<root><appid>KesionCMS</appid><status>0</status><body><message/><email/><question/><answer/><savecookie/><truename/><gender/><birthday/><qq/><msn/><mobile/><telephone/><address/><zipcode/><homepage/><userip/><jointime/><experience/><ticket/><valuation/><balance/><posts/><userstatus/></body></root>"
	
	Sql = "SELECT TOP 1 * FROM KS_User WHERE UserName='" & KS.R(UserName) & "'"
	Set Rs = Conn.Execute(Sql)
	If Not Rs.Eof And Not Rs.Bof Then
		XmlDoc.documentElement.selectSingleNode("body/email").text = Rs("email") & ""
		XmlDoc.documentElement.selectSingleNode("body/question").text = Rs("question") & ""
		XmlDoc.documentElement.selectSingleNode("body/answer").text = Rs("answer") & ""
		XmlDoc.documentElement.selectSingleNode("body/gender").text = Rs("sex") & ""
		XmlDoc.documentElement.selectSingleNode("body/birthday").text = ""
		XmlDoc.documentElement.selectSingleNode("body/mobile").text = RS("mobile")
		XmlDoc.documentElement.selectSingleNode("body/userip").text = Rs("LastLoginIP") & ""
		XmlDoc.documentElement.selectSingleNode("body/jointime").text = Rs("Joindate") & ""
		XmlDoc.documentElement.selectSingleNode("body/experience").text =""
		XmlDoc.documentElement.selectSingleNode("body/ticket").text = ""
		XmlDoc.documentElement.selectSingleNode("body/valuation").text = Rs("point") & ""
		XmlDoc.documentElement.selectSingleNode("body/balance").text = Rs("Money") & ""
		XmlDoc.documentElement.selectSingleNode("body/posts").text = Rs("zip") & ""
		XmlDoc.documentElement.selectSingleNode("body/userstatus").text = Rs("Locked")
		XmlDoc.documentElement.selectSingleNode("body/homepage").text = Rs("HomePage") & ""
		XmlDoc.documentElement.selectSingleNode("body/qq").text = Rs("qq")
		XmlDoc.documentElement.selectSingleNode("body/msn").text = rs("msn")
		XmlDoc.documentElement.selectSingleNode("body/truename").text = Rs("realName") & ""
		XmlDoc.documentElement.selectSingleNode("body/telephone").text = Rs("OfficeTel") & ""
		XmlDoc.documentElement.selectSingleNode("body/address").text = Rs("address") & ""
		Status = 0
		Messenge = Messenge & "<li>读取用户资料成功。</li>"
	Else
		Status = 1
		Messenge = Messenge & "<li>该用户不存在。</li>"
	End If
	Rs.Close
	Set Rs = Nothing
End Sub

Sub Checkname()
	Dim Rs,SQL,UserEmail
	UserEmail = KS.R(Trim(XmlDoc.documentElement.selectSingleNode("email").text))
	If KS.IsValidEmail(UserEmail) = False Then
		Messenge = "<li>您的Email有错误！</li>"
		Status = 1
		Exit Sub
	End If
	If CInt(KS.Setting(28)) = 1 Then
		Set Rs = Conn.Execute("SELECT userid FROM KS_User WHERE Email='" & UserEmail & "'")
		If Not Rs.EOF Then
			Status = 1
			Messenge = "<li>此邮箱["&UserEmail&"]已经占用，请您换一个邮箱再注册吧。</li>"
			Exit Sub
		End If
		Rs.Close:Set Rs = Nothing
	End If
	Set Rs = Conn.Execute("SELECT top 1 username FROM KS_User WHERE username = '" & UserName & "'")
	If Not (Rs.bof And Rs.EOF) Then
		Status = 1
		Messenge =  "<li>Sorry！此用户已经存在,请换一个用户名再试！</li>"
	Else
		Status = 0
		Messenge =  "<li><font color=red><b>" & UserName & "</b></font> 尚未被人使用，赶紧注册吧！</li>"
	End If
	Rs.Close:Set Rs = Nothing
End Sub

Sub UserReguser()
	Dim nickname,UserPass,UserEmail,Question,Answer,usercookies
	Dim strGroupName,Password,usersex,sex
	Dim Rs,SQL
	UserPass = KS.R(XmlDoc.documentElement.selectSingleNode("password").text)
	UserEmail = KS.R(Trim(XmlDoc.documentElement.selectSingleNode("email").text))
	Question = KS.R(XmlDoc.documentElement.selectSingleNode("question").text)
	Answer = KS.R(XmlDoc.documentElement.selectSingleNode("answer").text)
	sex = KS.R(XmlDoc.documentElement.selectSingleNode("gender").text)
	
	Dim NewRegUserMoney:NewRegUserMoney=KS.Setting(38)
	Dim NewRegUserScore:NewRegUserScore=KS.Setting(39)
	Dim NewRegUserPoint:NewRegUserPoint=KS.Setting(40)

	If sex = "0" Then
		usersex = "女"
	Else
		usersex = "男"
	End If
	usercookies = 1
	If UserName = "" Or UserPass = "" Then
		Status = 1
		Messenge = Messenge & "<li>请填写用户名或密码。"
		Exit Sub
	End If
	If Question = "" Then Question = KS.MakeRandomChar(20)
	If Answer = "" Then Answer = KS.MakeRandomChar(20)
	nickname = UserName
	Password = MD5(KS.R(UserPass),16)
	Answer = Answer
	If KS.IsValidEmail(UserEmail) = False Then
		Messenge = Messenge & "<li>您的Email有错误！</li>"
		Status = 1
		Exit Sub
	End If
	Set Rs = Conn.Execute("SELECT username FROM KS_User WHERE username='" & UserName & "'")
	If Not (Rs.BOF And Rs.EOF) Then
		Status = 1
		Messenge = Messenge & "<li>Sorry！此用户已经存在,请换一个用户名再试！</li>"
		Exit Sub
	End If
	Rs.Close:Set Rs = Nothing
	If CInt(KS.Setting(28)) = 1 Then
		Set Rs = Conn.Execute("SELECT userid FROM KS_User WHERE Email='" & UserEmail & "'")
		If Not Rs.EOF Then
			Status = 1
			Messenge = Messenge & "<li>对不起！本系统已经限制一个邮箱只能注册一个账号。</li><li>此邮箱["&UserEmail&"]已经占用，请您换一个邮箱再注册吧。</li>"
			Exit Sub
		End If
		Rs.Close:Set Rs = Nothing
	End If

	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM KS_User WHERE (userid is null)"
	Rs.Open SQL,Conn,1,3
	Rs.Addnew
		Rs("username") = UserName
		Rs("password") = Password
		RS("GroupID")=2    '设定默认用户类型为个人会员
		Rs("answer") = Answer
		Rs("question") = Question
		Rs("UserFace") = "Images/Face/0.gif"
		Rs("RealName") = UserName
		Rs("sex") = usersex
		Rs("Email") = UserEmail
		Rs("qq") = ""
		RS("RegDate")=Now
		RS("BeginDate")=Now '开始计算时间
		RS("LastLoginIP")=KS.GetIP
		RS("JoinDate")=Now
		RS("LastLoginTime")=Now
		
		 '新会员注册，更新相应的数据
		 RS("Money")=NewRegUserMoney
		 RS("Score")=NewRegUserScore
		 RS("Point")=NewRegUserPoint
		 Call KS.PointInOrOut(0,0,UserName,1,NewRegUserPoint,"系统","注册新会员,赠送的" & KS.Setting(46) & KS.Setting(45),0)
		 RS("Locked")=0
	Rs.update
	RS.movelast
	Conn.Execute("Update KS_User Set ChargeType=" & Conn.Execute("Select ChargeType From KS_UserGroup Where ID=" & RS("GroupID"))(0) & " Where UserID=" & RS("UserID"))
	RS.Close
			  '===================写入个人空间================
			  If KS.SSetting(1)=1 Then
			 RS.Open "Select * From KS_Blog Where Blogid is null",conn,1,3
			 RS.AddNew
			  RS("UserName")=UserName
			  RS("BlogName")=UserName & "的个人空间"
			  RS("ClassID") = KS.ChkClng(Conn.Execute("Select Top 1 ClassID From KS_BlogClass")(0))
			  RS("TemplateID")=KS.ChkClng(Conn.Execute("Select Top 1 ID From KS_BlogTemplate Where flag=2 and IsDefault='true'")(0))
			  RS("Announce")="暂无公告!"
			  RS("ContentLen")=500
			  RS("Recommend")=0
			  RS("Status")=1
			 RS.Update
			 RS.Close
			 end if
			 Set RS=Nothing
		    '==================================

	Status = 0
	Messenge = "用户注册成功。"
End Sub

Sub UesrLogin()
	Dim UserPass
	
	UserPass = XmlDoc.documentElement.selectSingleNode("password").text
	If UserName="" or UserPass="" Then
		Status = 1
		Messenge = Messenge & "<li>请填写用户名或密码。</li>"
		Exit Sub
	End If
	UserPass = Md5(UserPass,16)
	
	If ChkUserLogin(username,UserPass,1) Then
		Status = 0
		Messenge = Messenge & "<li>登陆成功。</li>"
	Else
		Status = 1
		Messenge = Messenge & "<li>登陆失败。</li>"
	End If
End Sub

Sub LogoutUser()
	If EnabledSubDomain Then
		Response.Cookies(KS.SiteSn).domain=RootDomain					
	Else
        Response.Cookies(KS.SiteSn).path = "/"
	End If
	Response.Cookies(KS.SiteSn)("UserName") = ""
	Response.Cookies(KS.SiteSn)("Password") = ""
	Response.Cookies(KS.SiteSn)("RndPassword")=""
End Sub

Sub UpdateUser()
	Dim Rs,SQL
	Dim UserPass,UserEmail,Question,Answer
	UserPass = XmlDoc.documentElement.selectSingleNode("password").text
	UserEmail = Trim(XmlDoc.documentElement.selectSingleNode("email").text)
	Question = XmlDoc.documentElement.selectSingleNode("question").text
	Answer = XmlDoc.documentElement.selectSingleNode("answer").text
	If UserPass <> "" Then
		UserPass = Md5(UserPass,16)
	End If
	If Answer <> "" THen
		Answer = Answer
	End If
	If KS.IsValidEmail(UserEmail) = False Then
		UserEmail = ""
	End If
	Set Rs = Server.CreateObject("Adodb.RecordSet")
	SQL = "SELECT TOP 1 * FROM [KS_User] WHERE Username='" & UserName & "'"
	Rs.Open SQL,Conn,1,3
	If Not Rs.Eof And Not Rs.Bof Then
		If UserPass <> "" Then Rs("password") = UserPass
		If Answer <> "" THen Rs("answer") = Answer
		If UserEmail <> "" Then Rs("email") = UserEmail
		If Question <> "" Then Rs("question") = Question
		Rs.update
		Status = 0
		Messenge = "<li>基本资料修改成功。</li>"
		Response.Cookies(KS.SiteSN)("password") = UserPass
	Else
		Status = 1
		Messenge = "<li>该用户不存在，修改资料失败。</li>"
	End If
	Rs.Close:Set Rs = Nothing
End Sub

Sub Deleteuser()
	Dim Del_Users,i,AllUserID,Del_UserName
	Dim Rs
	Del_Users = Split(UserName,",")
	For i = 0 To UBound(Del_Users)
		Del_UserName = KS.R(Del_Users(i))
		Set Rs = Conn.Execute("SELECT userid,username FROM [KS_User] WHERE UserName='" & Del_UserName & "'")
		If Not (Rs.Eof And Rs.Bof) Then
			Conn.Execute ("DELETE FROM KS_User WHERE UserName='" & Del_UserName & "')")
			Conn.Execute ("DELETE FROM KS_Favorite WHERE UserName='" & Del_UserName & "')")
			Conn.Execute ("DELETE FROM KS_Comment WHERE UserName='" & Del_UserName & "')")
			Messenge = Messenge & "<li>用户（" & Del_UserName & "）删除成功。</li>"
		End If
	Next
	Set Rs = Nothing
	Status = 0
End Sub

Sub Lockuser()
	Dim UserStatus
	If XmlDoc.documentElement.selectSingleNode("userstatus") is Nothing Then
		Messenge = "<li>参数非法，中止请求。</li>"
		Status = 1
		Exit Sub
	ElseIf Not IsNumeric(XmlDoc.documentElement.selectSingleNode("userstatus").text) Then
		Messenge = "<li>参数非法，中止请求。</li>"
		Status = 1
		Exit Sub
	Else
		UserStatus = Clng(XmlDoc.documentElement.selectSingleNode("userstatus").text)
	End If
	If UserStatus = 0 Then
		Conn.Execute ("UPDATE KS_User SET Locked=0 WHERE Username='" & UserName & "'")
	Else
		Conn.Execute ("UPDATE KS_User SET Locked=1 WHERE Username='" & UserName & "'")
	End If
	Status = 0
End Sub

Sub SaveUserCookie()
	Dim S_syskey,Password,usercookies,TruePassWord,userclass,Userhidden
	
	S_syskey = Request.QueryString("syskey")
	UserName = KS.R(Request.QueryString("UserName"))
	Password = Request.QueryString("Password")
	usercookies = Request.QueryString("savecookie")
	If UserName="" or S_syskey="" Then Exit Sub
	Dim NewMd5,OldMd5
	NewMd5 = Md5(UserName & API_ConformKey,16)
	Md5OLD = 1
	OldMd5 = Md5(UserName & API_ConformKey,16)
	Md5OLD = 0
	If Not (S_syskey=NewMd5 or S_syskey=OldMd5) Then
		Exit Sub
	End If
	If usercookies="" or Not IsNumeric(usercookies) Then usercookies = 0
	
	'用户退出
	If Password = "" Then
	    If EnabledSubDomain Then
			Response.Cookies(KS.SiteSn).domain=RootDomain					
		Else
             Response.Cookies(KS.SiteSn).path = "/"
		End If
		Response.Cookies(KS.SiteSn)("UserName") = ""
		Response.Cookies(KS.SiteSn)("Password") = ""
		Response.Cookies(KS.SiteSn)("RndPassword")=""
		Exit Sub
	End If
	ChkUserLogin username,password,usercookies
End Sub

Function ChkUserLogin(username,password,usercookies)
	ChkUserLogin = False
	Dim Rs,SQL,RndPassWord
	RndPassWord=KS.R(KS.MakeRandomChar(20))
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM [KS_User] WHERE username='" & UserName & "'"
	Rs.Open SQL, Conn, 1, 3
	If Not (Rs.BOF And Rs.EOF) Then
		If password <> Rs("password") Then
			ChkUserLogin = False
			Exit Function
		End If
		If Rs("Locked") <> 0 Then
			ChkUserLogin = False
			Exit Function
		End If
		'登录成功，更新用户相应的数据
		If datediff("n",RS("LastLoginTime"),now)>=KS.Setting(36) then '判断时间
		 RS("Score")=RS("Score")+KS.Setting(37)
		end if
		 RS("LastLoginIP") = KS.GetIP
         RS("LastLoginTime") = Now()
         RS("LoginTimes") = RS("LoginTimes") + 1
		 RS("RndPassword")= RndPassWord	
		Rs.Update
		
		Select Case usercookies
		Case 0
			Response.Cookies(KS.SiteSn)("usercookies") = usercookies
		Case 1
			Response.Cookies(KS.SiteSn)("usercookies") = usercookies
			Response.Cookies(KS.SiteSn).Expires=Date+1
		Case 2
			Response.Cookies(KS.SiteSn).Expires=Date+31
			Response.Cookies(KS.SiteSn)("usercookies") = usercookies
		Case 3
			Response.Cookies(KS.SiteSn).Expires=Date+365
			Response.Cookies(KS.SiteSn)("usercookies") = usercookies
		End Select
		If EnabledSubDomain Then
			Response.Cookies(KS.SiteSn).domain=RootDomain					
		Else
           Response.Cookies(KS.SiteSn).path = "/"
		End If
		Response.Cookies(KS.SiteSn)("UserName") = Rs("username")
		Response.Cookies(KS.SiteSn)("Password") = Rs("password")
		Response.Cookies(KS.SiteSn)("RndPassword")=RndPassWord
		ChkUserLogin = True
	End If
	Rs.Close:Set Rs = Nothing
End Function

%>
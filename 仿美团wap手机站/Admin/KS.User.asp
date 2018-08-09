<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../Plus/md5.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.BaseFunCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="Include/Session.asp"-->
<!--#include file="../API/cls_api.asp"-->
<%

Dim KSCls
Set KSCls = New Admin_User
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_User
        Private KS,KSCls
		Private MaxPerPage
		Private rsAdmin,sqlAdmin
		Private UserID,UserSearch,Keyword,strField,Page,sql,FoundErr,RS,TotalPut,TotalPages,I
		Private Action,ComeUrl,strFileName
		Private ValidDays,tmpDays,BeginID,EndID
		Private ErrMsg
		Private Sub Class_Initialize()
		  MaxPerPage=20
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub
       Sub Kesion()
	   
			If KS.G("Action")="CheckUserName" Then Call CheckUserName():Response.End()
            Response.Write "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd"">"
            Response.WRite "<html xmlns=""http://www.w3.org/1999/xhtml"">"
			Response.Write"<head>"
			Response.Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			Response.Write"<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			Response.Write"<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>" & vbCrLf
			Response.Write"<script language=""JavaScript"" src=""../KS_Inc/jquery.js""></script>" & vbCrLf
			Response.Write"</head>"
			Response.Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			Response.Write"	<ul id='mt'> "
			Response.Write "<div id='mtl'>快速查找用户：</div><li><a href=""KS.User.asp?Action=Search"">搜索用户</a></li>&nbsp;|&nbsp;<a href=""?UserSearch=12"">所有用户</a>&nbsp;|&nbsp;<a href=""?UserSearch=13"" style='color:red'>在线用户</a>&nbsp;|&nbsp;<a href=""?UserSearch=1"">被锁住</a>&nbsp;|&nbsp;<a href=""?UserSearch=2"">管理员</a>&nbsp;|&nbsp;<a href=""?UserSearch=3"">待审批</a>&nbsp;|&nbsp;<a href=""?UserSearch=4"">待邮件激活</a>&nbsp;|&nbsp;<a href=""?UserSearch=5"">24小时内登录</a>&nbsp;|&nbsp;<a href=""?UserSearch=6"">24小时内注册</a>&nbsp;|&nbsp;<a href=""?UserSearch=15"">已过期</a>&nbsp;|&nbsp;<a href=""?UserSearch=16"">未过期</a>"
			Response.Write	" </ul>"
		     If Not KS.ReturnPowerResult(0, "KMUA10002") Then
			  Response.Write ("<script>$(parent.document).find('#BottomFrame')[0].src='javascript:history.back();';</script>")
			  Call KS.ReturnErr(1, "")
			End If

		ComeUrl=Cstr(Request.ServerVariables("HTTP_REFERER"))
		keyword		= Trim(request("keyword"))
		strField	= Trim(request("Field"))
		UserSearch	= KS.ChkClng(request("UserSearch"))
		Action		= Trim(request("Action"))
		UserID		= Trim(Request("UserID"))
		strFileName	= "KS.User.asp"
		Page	= KS.ChkClng(request("page"))
		if keyword<>"" then keyword=KS.R(keyword)
		%>
		<SCRIPT language=javascript>
		function unselectall()
		{
			if(document.myform.chkAll.checked){
			document.myform.chkAll.checked = document.myform.chkAll.checked&0;
			} 	
		}
		
		function CheckAll(form)
		{
		  for (var i=0;i<form.elements.length;i++)
			{
			var e = form.elements[i];
			if (e.Name != "chkAll"  && e.disabled==false)
			   e.checked = form.chkAll.checked;
			}
		}
		</SCRIPT>
		</head>
		<%
		Select Case Action
		Case "Add"            call AddUser()
		Case "SaveAdd"	      call SaveAdd()
		Case "Modify"         call Modify()
		Case "SaveModify"     call SaveModify()
		Case "Del"            call DelUser()
		Case "Lock"	          call locked()
		Case "UnLock"         call Unlocked()
		Case "Verify"	      call verify(0)
		Case "UnVerify"       call verify(2)
		Case "Active"         call Unlocked()
		Case "Move"	          call MoveUser()
		Case "AddMoney"	      call AddMoney()
		Case "AddScore"       call AddScore()
		Case "SaveAddMoney"	  call SaveAddMoney()
		Case "SaveAddScore"   call SaveAddScore()
		Case "AddZJ"          call AddZJ()
		Case "SaveAddZJ"      call SaveAddZJ
		Case "Search"         Call ShowSearch()
		Case "ShowDetail"	  Call ShowDetail()
		Case Else	call main()
		End Select
		if FoundErr=True then KS.ShowError(ErrMsg)
		If Action<>"ShowDetail" Then
		Response.Write ""
		End If
		End Sub
		
		Sub Main()
		    Dim GroupID:GroupID=KS.G("GroupID")
				dim strGuide ,sSQL,Param
				strGuide="<table style='margin-top:0px' width='100%' align='center' border='0' cellpadding='0' cellspacing='1'><tr class='list'><td align='center' height='25'>&nbsp;"
				sSQL = " UserID,UserName,GroupID,ChargeType,Point,BeginDate,LastLoginIP,LastLoginTime,LoginTimes ,locked,Edays,IsOnline,Money"
				Select Case UserSearch
				 Case 1
				    Param="locked=1"
					strGuide=strGuide & "所有被锁住的用户"
                Case 2
					Param="groupid=1"
					strGuide=strGuide & "所有管理员身份的用户"
                Case 3
					Param="locked=2"
					strGuide=strGuide & "待管理员认证用户"
                Case 4
					Param="locked=3"
					strGuide=strGuide & "待邮件验证的用户"
				Case 5
				   Param="datediff(" & DataPart_H & ",LastLoginTime," & SqlNowString & ")<25"
				   strGuide=strGuide & "最近24小时内登录的用户"
				Case 6
				    Param="datediff(" & DataPart_H & ",RegDate," & SqlNowString & ")<25"
					strGuide=strGuide & "最近24小时内注册的用户"
				Case 10
					param="GroupID=" & GroupID
					strGuide=strGuide & KS.GetUserGroupName(GroupID)
				Case 13
				    param="IsOnline=1"
					strGuide=strGuide & "在线用户列表"
				Case 14
				    param="ClubGradeID=" & KS.ChkClng(Request("ClubGradeID"))
					strGuide=strGuide & "论坛等级为" & KS.A_G(KS.ChkClng(Request("ClubGradeID")),"usertitle")& "的用户列表"
				Case 15
				    Param="chargetype=2 and datediff(" & DataPart_d & ",beginDate," & SqlNowString & ")>edays"
					strGuide=strGuide & "已过服务期的用户列表"
				Case 16
				    Param="chargetype=2 and datediff(" & DataPart_d & ",beginDate," & SqlNowString & ")<=edays"
					strGuide=strGuide & "未过服务期的用户列表"
				Case 11
					UserID = KS.ChkClng(UserID)
					if UserID>0 then
						param="UserID="&UserID&""
					else 
						Dim strsql
						strsql=""
						if request("username")<>"" then
							if request("usernamechk")="yes" then
								strsql=strsql & " username='"&request("username")&"'"
							else
								strsql=strsql &" username like '%"&request("username")&"%'"
							end if
						end if
						if cint(request("GroupID"))>0 then
							if strsql="" then
								strsql=strsql & " GroupID="&request("GroupID")&""
							else
								strsql=strsql & " and GroupID="&request("GroupID")&""
							end if
						end if
						if request("Email")<>"" then
							if strsql="" then
								strsql=strsql & " Email like '%"&request("Email")&"%'"
							else
								strsql=strsql & " and Email like '%"&request("Email")&"%'"
							end if
						end if
		            '======特殊搜索=======
						dim Tsqlstr
						if request("loginT")<>"" then
							if request("loginR")="more" then
								Tsqlstr=" LoginTimes >= "&KS.Chkclng(request("loginT"))
							else
								Tsqlstr=" LoginTimes <= "&KS.Chkclng(request("loginT"))
							end if 	
							if strsql="" then 
								strsql=Tsqlstr
							else
								strsql=strsql & " and" & Tsqlstr
							end if
						end if
		
						if request("vanishT")<>"" then
							if request("vanishR")="more" then
								Tsqlstr=" datediff(" & DataPart_D & ",LastLoginTime,"&SqlNowString&") >= "&KS.Chkclng(request("vanishT"))&""
							else
								Tsqlstr=" datediff(" & DataPart_D & ",LastLoginTime,"&SqlNowString&") <= "&KS.Chkclng(request("vanishT"))&""
							end if 	
							if strsql="" then 
								strsql=Tsqlstr
							else
								strsql=strsql & " and" & Tsqlstr
							end if
						end if
		
						if request("regT")<>"" then
							if request("regR")="more" then
								Tsqlstr=" datediff(" & DataPart_D & ",RegDate,"&SqlNowString&") >= "&KS.Chkclng(request("regT"))
							else
								Tsqlstr=" datediff(" & DataPart_D & ",RegDate,"&SqlNowString&") <= "&KS.Chkclng(request("regT"))
							end if 	
							if strsql="" then 
								strsql=Tsqlstr
							else
								strsql=strsql & " and" & Tsqlstr
							end if
						end if
		
						if request("artcleT")<>"" then
							if request("artcleR")="more" then
								Tsqlstr=" (select count(id) from ks_iteminfo where inputer=ks_user.username) >= "&KS.Chkclng(request("artcleT"))
							else
								Tsqlstr=" (select count(id) from ks_iteminfo where inputer=ks_user.username) <= "&KS.Chkclng(request("artcleT"))
							end if 	
							if strsql="" then 
								strsql=Tsqlstr
							else
								strsql=strsql & " and" & Tsqlstr
							end if
						end if
		              '======特殊搜索结束======
						If strsql = "" Then
							FoundErr=True
							ErrMsg=ErrMsg & "<br><li>请指定搜索参数！</li>"
							Exit Sub
						End If
						If Request("Searchmax") = "" Or Not Isnumeric(Request("Searchmax")) Then
							param=strsql
						Else
						    param=strsql
							'Sql = "Select top "&Request("Searchmax")&" "&sSQL&" From KS_User Where " & strsql & " order by UserID desc"
						End If
					end if '''ID
					strGuide=strGuide & "查询结果："
				Case Else
					Param="1=1"
					strGuide=strGuide & "所有用户"
				End Select
				strGuide=strGuide & "</td><td width='150' align='center'>"
				if FoundErr=True then Exit Sub
				
				If KS.C("SuperTF")<>"1" then Param=Param & " and (groupid<>1 or username='" & KS.C("AdminName") & "')"
				
				if Page < 1 then Page=1
				
						SQL=KS.GetPageSQL("KS_User","UserID",MaxPerPage,Page,1,Param,sSQL)
						Set RS = Server.CreateObject("AdoDb.RecordSet")
						RS.Open SQL, conn, 1, 1
				
				'Set rs=Server.CreateObject("Adodb.RecordSet")
				'rs.Open sql,Conn,1,1
				if rs.eof and rs.bof then
					TotalPut=0
					Response.Write strGuide & "共找到 <font color=#ff6600>0</font> 个用户&nbsp;&nbsp;&nbsp;&nbsp;</td></tr></table>"
					rs.Close:set rs=Nothing
				else
					'TotalPut=rs.recordcount
					TotalPut=Conn.Execute("Select count(userid) from [KS_User] where " & Param)(0)
					'Response.Write strGuide & "共找到 <font color=#ff6600>" & TotalPut & "</font> 个用户&nbsp;&nbsp;&nbsp;&nbsp;</td></tr></table>"
					if (TotalPut mod MaxPerPage)=0 then
						TotalPages = TotalPut \ MaxPerPage
					else
						TotalPages = TotalPut \ MaxPerPage + 1
					end if
					SQL = rs.GetRows(MaxPerPage)
					rs.Close:set rs=Nothing
					ShowContent
				end if
		End Sub
		
		Sub ShowContent()
		%>
		  <table width="100%" style="border-top:1px #CCCCCC solid" border="0" align="center" cellspacing="0" cellpadding="0">
		  		  <form name="myform" method="Post" action="KS.User.asp" onSubmit="return confirm('确定要执行选定的操作吗？');">
				  <tr class='sort'>
					<td width="30" align="center">选中</td>
					<td width="30" align="center">ID</td>
					<td width="80" align="center"> 用户名</td>
					<td height="22" align="center">所属用户组</td>
					<td align="center"><%=KS.Setting(45)%>/天数</td>
					<td height="22" align="center">最后登录IP</td>
					<td align="center">最后活动时间</td>
					<td width="60" align="center">登录</td>
					<td width="40" align="center">状态</td>
					<td align="center">操作</td>
				  </tr>
			  <%
				For i=0 To Ubound(SQL,2)
			 %>
				  <tr height="23" class='list' id='u<%=SQL(0,i)%>' onclick="chk_iddiv('<%=SQL(0,i)%>')" onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
					<td class='splittd' width="30" align="center"><input <%If KS.C("SuperTF")<>1 and SQL(2,i)=4 and  SQL(1,i)<>KS.C("AdminName") Then Response.Write " Disabled" %> name="UserID" type="checkbox"  onclick="chk_iddiv('<%=SQL(0,i)%>')" id='c<%=SQL(0,i)%>'  value="<%=SQL(0,i)%>"></td>
					<td class='splittd' width="30" align="center"><%=SQL(0,i)%></td>
					<td class='splittd' width="80" align="center"><%
					If KS.C("SuperTF")<>1 and SQL(2,i)=4 and  SQL(1,i)<>KS.C("AdminName") then
					 response.write "<font color=red>" & SQL(1,i) & "</font>"
					else
					Response.Write "<a href='KS.User.asp?Action=ShowDetail&UserID=" & SQL(0,i) & "'>" & SQL(1,i) & "</a>"
					end if
					if SQL(11,i)="1" Then
					 response.write "<font color=red>在线</font>"
					end if
					%>
					</td>
					<td class='splittd' align="center"><font color=blue><%=KS.GetUserGroupName(SQL(2,i))%></font></td>
					<td class='splittd' align="center">
					<%
				if SQL(3,i)=1 then
					if SQL(4,i)<=0 then
						Response.Write "<font color=#ff6600>" & SQL(4,i) & "</font> " & KS.Setting(46)
					else
						if SQL(4,i)<=10 then
							Response.Write "<font color=blue>" & SQL(4,i) & "</font> " & KS.Setting(46)
						else
							Response.Write SQL(4,i) & " " & KS.Setting(46)
						end if
					end if
				elseif SQL(3,i)=2 then
				    ValidDays=SQL(10,i)
					tmpDays = ValidDays-DateDiff("D",SQL(5,i),now())
					if tmpDays<=10 then
						Response.Write "<font color=#ff0033>" & tmpDays & "</font> 天"
					else
						Response.Write "<font color=#0000ff>" & tmpDays & "</font> 天"
					end if
				else
				   response.write "<font color=red>无限期</font>"
				end if
				   response.write " <font color=#999999>余额 <font color=green>" & sql(12,i) &"</font> 元</font>"
				%></td>
					<td class='splittd' align="center"> <%
				if SQL(6,i)<>"" then
					Response.Write SQL(6,i)
				else
					Response.Write "&nbsp;"
				end if%> </td>
					<td class='splittd' align="center"> <%=SQL(7,i)%> </td>
					<td class='splittd' width="60" align="center"><%=SQL(8,i)%> 次</td>
					<td class='splittd' width="40" align="center"><%
				select case SQL(9,i)
				   case 1 Response.Write "<font color=#ff6600>已锁定</font>"
				   case 2 response.write "<font color=blue>待审核</font>"
				   case 3 response.write "<font color=green>待激活</font>"
				   case else
					Response.Write "正常"
				end select%></td>
					<td class='splittd' align="center"><%
				If KS.C("SuperTF")<>1 and SQL(2,i)=4 and  SQL(1,i)<>KS.C("AdminName") then
					 response.write "---"
				else	
					Response.Write "<a href='KS.User.asp?Action=Modify&UserID=" & SQL(0,i) & "'>改</a>&nbsp;"
					if SQL(2,i)<>1 then  '管理员判断
						if SQL(9,i)=0 then
							Response.Write "<a href='KS.User.asp?Action=Lock&UserID=" & SQL(0,i) & "'>锁</a>&nbsp;"
						else
							Response.Write "<a href='KS.User.asp?Action=UnLock&UserID=" & SQL(0,i) & "'>解</a>&nbsp;"
						end if
						Response.Write "<a href='KS.User.asp?Action=Del&UserID=" & SQL(0,i) & "' onClick='return confirm(""确定要删除此用户吗？"");'>删</a>&nbsp;"
						 If SQL(3,i)=1 Then
						   Response.Write "<a href='KS.User.asp?Action=AddMoney&UserID=" & SQL(0,i) & "'>续" & KS.Setting(45) & "</a>"
						 ElseIf SQL(3,I)=2 Then
						   Response.Write "<a href='KS.User.asp?Action=AddMoney&UserID=" & SQL(0,i) & "'>续天数</a>"
						 End IF
					end if
					%> <a href='KS.User.asp?Action=AddZJ&UserID=<%=SQL(0,I)%>'>续费</a>
					<a href='KS.User.asp?Action=AddScore&UserID=<%=SQL(0,I)%>'>加积分</a>
				<%end if%>
				</td>
				  </tr>
			<%
			Next
			%>

		  <tr class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
			<td colspan=10 height="30">&nbsp;<label><input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">
					  选中</label>&nbsp;<strong>操作：</strong> 
					  <input name="Action" type="radio" value="Del" checked onClick="this.form.GroupID.disabled=true">删除 
					  <input name="Action" type="radio" value="Lock" onClick="this.form.GroupID.disabled=true">锁定 
					  <input name="Action" type="radio" value="UnLock" onClick="this.form.GroupID.disabled=true">解锁 
					  <input name="Action" type="radio" value="Verify" onClick="this.form.GroupID.disabled=true">审核 
					  <input name="Action" type="radio" value="UnVerify" onClick="this.form.GroupID.disabled=true">待审 
					  <input name="Action" type="radio" value="Active" onClick="this.form.GroupID.disabled=true">激活 
					  <input name="Action" type="radio" value="Move" onClick="this.form.GroupID.disabled=false">移动到
					  <select name="GroupID" id="GroupID" disabled>
						<%=KS.GetUserGroup_Option(3)%>
					  </select>
					  &nbsp;<input type="submit" name="Submit" class='button' value=" 执 行 " >&nbsp;<input class='button' type="button" name="Submit" value="发送邮件" onclick="this.form.action='KS.UserMail.asp?InceptType=2';this.form.submit()" >&nbsp;<input class='button' type="button" name="Submit" value="发送短信" onclick="this.form.action='KS.UserMessage.asp?action=new';this.form.submit()" > </td>
		  </tr></form>
		<tr valign=middle class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
			<td colspan="10" align="right">
			<%
			 Call KS.ShowPage(totalput, MaxPerPage, "", Page,true,true)
			%>
		</td>
		</tr>
		</table>
		<%
		End Sub
		
		Sub ShowSearch()
		%>
		<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1">
		<form name="form2" method="get" action="KS.User.asp">
		<tr Class="sort">
			<td height="25" colspan="2" align="center"><strong>高级查询</strong></td>
		</tr>
		<tr class="tdbg">
			<td width="100" height="25" class="clefttitle" align="right"><strong>注意事项:</strong></td>
			<td>在记录很多的情况下搜索条件越多查询越慢，请尽量减少查询条件；最多显示记录数也不宜选择过大</td>
		</tr>
		<!--
		<tr class="tdbg">
			<td class="clefttitle" align="right"><strong>最多查询记录数:</strong></td>
			<td><input class="textbox" size="45" name="searchMax" type="text" value="100"></td>
		</tr>
		-->
		<tr class="tdbg">
			<td class="clefttitle" align="right"><strong>用户ID:</strong></td>
			<td><input class="textbox" size="45" name="userid" type="text"></td>
		</tr>
		<tr class="tdbg">
			<td class="clefttitle" align="right"><strong>用户名:</strong></td>
			<td><input class="textbox" size="45" name="username" type="text">&nbsp;<input type="checkbox" name="usernamechk" value="yes" checked>用户名完整匹配</td>
		</tr>
		<tr class="tdbg">
			<td class="clefttitle" align="right"><strong>用户组:</strong></td>
			<td>
			<select size="1" name="GroupID">
			<option value="0" selected>任意</option>
			<%=KS.GetUserGroup_Option(0)%>
			</select>
		  </td>
		</tr>
		<tr class="tdbg">
			<td class="clefttitle" align="right"><strong>Email包含:</strong></td>
			<td><input class="textbox" size="45" name="Email" type=text></td>
		</tr>
		</table>
		<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1">
		<!--特殊搜索-->
		<tr class="sort">
			<td height="23" colspan="2" align="left">特殊查询&nbsp;（注意： 
		  <多于> 或 <少于> 已默认包含 <等于>；条件留空则不使用此条件 ）</td>
		</tr>
		<tr class="tdbg">
			<td>登录次数:
		  <input type=radio value=more name="loginR" checked ID="Radio1">&nbsp;多于&nbsp;<input type=radio value=less name="loginR" ID="Radio2">&nbsp;少于&nbsp;&nbsp;<input class="textbox" size=5 name="loginT" type=text ID="Text1"> 次&nbsp;&nbsp;</td>
			<td>消失天数:
		  <input type=radio value=more name="vanishR" checked ID="Radio3">&nbsp;多于&nbsp;<input type=radio value=less name="vanishR" ID="Radio4">&nbsp;少于&nbsp;&nbsp;<input class="textbox" size=5 name="vanishT" type=text ID="Text2"> 天&nbsp;&nbsp;</td>
		</tr>
		<tr class="tdbg">
			<td width="50%">注册天数:
		  <input type=radio value=more name="regR" checked ID="Radio5">&nbsp;多于&nbsp;<input type=radio value=less name="regR" ID="Radio6">&nbsp;少于&nbsp;&nbsp;<input class="textbox" size=5 name="regT" type=text ID="Text3"> 天&nbsp;&nbsp;</td>
			<td width="50%">发表文章:
		  <input type=radio value=more name="artcleR" checked ID="Radio7">&nbsp;多于&nbsp;<input type=radio value=less name="artcleR" ID="Radio8">&nbsp;少于&nbsp;&nbsp;<input class="textbox" size=5 name="artcleT" type=text ID="Text4"> 篇&nbsp;&nbsp;</td>
		</tr>
		<!--特殊搜索结束-->
		<tr class="tdbg">
		  <td width="100%" colspan="2" align="center"><input name="submit" class='button' type=submit value="   搜  索   "></td>
		</tr>
		<input name="UserSearch" type="hidden" id="UserSearch" value="11">
		</table>
		</form>
		<%
		end sub
		
		sub AddUser()
		 Dim GroupID:GroupID=KS.ChkClng(KS.G("GroupID"))
		 If GroupID=0 Then GroupID=2
		%>
		<SCRIPT language=javascript>
		function CheckForm()
		{
		  if(document.myform.UserName.value=="")
			{
			  alert("用户名不能为空！");
			  document.myform.UserName.focus();
			  return false;
			}
		  if(document.myform.Password.value=="")
			{
			  alert("用户密码不能为空！");
			  document.myform.Password.focus();
			  return false;
			}
		
		  if(document.myform.Question.value=="")
			{
			  alert("密码问题不能为空！");
			  document.myform.Question.focus();
			  return false;
			}
		  if(document.myform.Answer.value=="")
			{
			  alert("密码答案不能为空！");
			  document.myform.Answer.focus();
			  return false;
			}
		  if(document.myform.Email.value=="")
			{
			  alert("用户Email不能为空！");
			  document.myform.Email.focus();
			  return false;
			}

		}
		 checkaccount=function(val){
		  if(val=='')
		  {
			alert('请输入用户名称!');
			$('input[name=UserName]').focus();
			return false;
		  }
		  if(val.length<<%=KS.Setting(29)%>||val.length><%=KS.Setting(30)%>)
		  {
			alert('用户长度必须大于等于<%=KS.Setting(29)%>位小于等于<%=KS.Setting(30)%>位!');
			$('input[name=UserName]').focus();
			return false;
		  }
		  window.open('?action=CheckUserName&username='+val,'','width=0,height=0');
		 }
		</script>
		<script type="text/javascript" src="../editor/ckeditor.js" mce_src="../editor/ckeditor.js"></script>
		<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1">
		<FORM name="myform" action="KS.User.asp" method="post" onsubmit="return(CheckForm());">
			<TR class="sort">
			  <TD height="22" colspan="4" align="center"><span style="margin-top:10px;text-align:center;font-weight:bold">添加新用户</span></TD>
		  </TR>
			<TR class="tdbg"> 
				<TD width="80" height="25" align="right"  class="clefttitle"><strong>用户等级：</strong></TD>
			  <TD height="25"><select name="GroupID" id="GroupID" onchange="location.href='KS.User.asp?action=Add&amp;GroupID='+this.value;">
                <%=KS.GetUserGroup_Option(GroupID)%>
              </select>
			   论坛头衔
			  <select name="clubgradeid">
			  <%KS.LoadAskGrade
			  dim node,xml,master,masterarr,i
			   set xml=Application(KS.SiteSN&"_AskGrade")
			   If IsObject(XML) Then
			     
			    for each node in xml.documentelement.selectnodes("row[@typeflag='1']")
				  if node.selectsinglenode("@gradeid").text=20 then
					  response.write "<option value='" & node.selectsinglenode("@gradeid").text &"' selected>" & node.selectsinglenode("@usertitle").text & "</option>"
				  else
					  response.write "<option value='" & node.selectsinglenode("@gradeid").text &"'>" & node.selectsinglenode("@usertitle").text & "</option>"
				  end if
			    next
			   end if
			  %>
			  </select>
			  问吧头衔
			  <select name="gradeid">
			  <%
			   set xml=Application(KS.SiteSN&"_AskGrade")
			   If IsObject(XML) Then
			     
			    for each node in xml.documentelement.selectnodes("row[@typeflag='0']")
				   if node.selectsinglenode("@gradeid").text=6 then
					  response.write "<option value='" & node.selectsinglenode("@gradeid").text &"' selected>" & node.selectsinglenode("@usertitle").text & "</option>"
				   else
					  response.write "<option value='" & node.selectsinglenode("@gradeid").text &"'>" & node.selectsinglenode("@usertitle").text & "</option>"
				   end if
			    next
			   end if
			  %>
			  </select>
			  </TD>
				<TD width="80" height="25" align="right" class="clefttitle"><strong>用户状态：</strong></TD>
			  <TD><input type="radio" name="locked" value="0" checked="checked" />
正常&nbsp;&nbsp;
<input type="radio" name="locked" value="1" />
锁定</TD>
			</TR>
			<TR class="tdbg"> 
				<TD width="80" height="25" align="right" class="clefttitle"><strong>用户名称：</strong></TD>
			  <TD> <Input class="textbox" Name="UserName" id="UserName" type=text size=20> <font color="red">*</font><input type="button" name="Submit22" value="检测帐号"  onClick="checkaccount($('#UserName').val())" class="button"></TD>
				<TD width="80" height="25" align="right" class="clefttitle"><strong>电子邮箱：</strong></TD>
				<TD>  <Input class="textbox" Name="Email" type=text size=30 Value=""><font color="red">*</font></TD>
			</TR>
				<TR class="tdbg"> 
					<TD width="80" height="25" align="right" class="clefttitle"><strong>用户密码：</strong></TD>
				  <TD height="25"><INPUT class="textbox" type="password" name="Password" value="" size="30" maxLength="12"><font color="red">*</font> <font class=tips>用户登录时的密码</font></TD>
					<TD width="80" height="25" align="right" class="clefttitle"><strong>手机号码：</strong></TD>
					<TD height="25"><INPUT class="textbox" type="text" name="mobile" value="" size="30" maxLength="12"/></TD>
				</TR>
			<TR  class="tdbg"> 
				<TD width="80" height="25" align="right" class="clefttitle"><strong>密码问题：</strong></TD>
				<TD><Select id=Question name=Question>
                    <Option value="" selected>--请您选择--</Option>
                    <Option value="我的宠物名字？">我的宠物名字？</Option>
                    <Option value="我最好的朋友是谁？">我最好的朋友是谁？</Option>
                    <Option value="我最喜爱的颜色？">我最喜爱的颜色？</Option>
                    <Option value="我最喜爱的电影？">我最喜爱的电影？</Option>
                    <Option value="我最喜爱的影星？">我最喜爱的影星？</Option>
                    <Option value="我最喜爱的歌曲？">我最喜爱的歌曲？</Option>
                    <Option value="我最喜爱的食物？">我最喜爱的食物？</Option>
                    <Option value="我最大的爱好？">我最大的爱好？</Option>
                    <Option value="我中学校名全称是什么？">我中学校名全称是什么？</Option>
                    <Option value="我的座右铭是？">我的座右铭是？</Option>
                    <Option value="我最喜欢的小说的名字？">我最喜欢的小说的名字？</Option>
                    <Option value="我最喜欢的卡通人物名字？">我最喜欢的卡通人物名字？</Option>
                    <Option value="我母亲/父亲的生日？">我母亲/父亲的生日？</Option>
                    <Option value="我最欣赏的一位名人的名字？">我最欣赏的一位名人的名字？</Option>
                    <Option value="我最喜欢的运动队全称？">我最喜欢的运动队全称？</Option>
                    <Option value="我最喜欢的一句影视台词？">我最喜欢的一句影视台词？</Option>
                  </Select>
			  <font color="#FF6600">*</font> </TD>
				<TD width="80" height="25" align="right" class="clefttitle"><strong>问题答案：</strong></TD>
				<TD><INPUT class="textbox" type=text maxLength=20 size=30 name="Answer">			</TD>
			</TR>
			<TR class="tdbg">
				<TD width="80" height="25" align="right" class="clefttitle"><strong>计费方式：</strong></TD>
				<TD Colspan=3><input name="ChargeType" type="radio" value="1" checked>
				扣点数<font color="#0066CC">（推荐）</font>
				<input type="radio" name="ChargeType" value="2">
			  有效期(在有效期内，用户可以任意阅读收费内容)
			  <input type="radio" name="ChargeType" value="3"/>
无限期</TD>
			</TR>
			

			<tr class="sort">
			  <td colspan="4" align="center">====自定义选项====</td>
			</tr>
			<tr class="tdbg">
			  <td colspan="4">
			 <%
			            
					Dim Template:Template=LFCls.GetSingleFieldValue("Select top 1 Template From KS_UserForm Where ID=" & KS.U_G(GroupID,"formid"))
							
						   Dim FieldsList:FieldsList=LFCls.GetSingleFieldValue("Select top 1 FormField From KS_UserForm Where ID=" & KS.U_G(GroupID,"formid"))
						   Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
						   RS.Open "Select FieldID,FieldType,FieldName,DefaultValue,Width,Height,Options,EditorType,ShowUnit,UnitOptions from KS_Field Where ChannelID=101 Order By OrderID",conn,1,1
						   Dim SQL,K,N,InputStr,O_Arr,O_Len,F_V,O_Value,O_Text,BRStr
						   If Not RS.EOf Then SQL=RS.GetRows(-1):RS.Close():Set RS=Nothing
						   If IsArray(SQL) Then
						   For K=0 TO Ubound(SQL,2)
							 If KS.FoundInArr(FieldsList,SQL(0,k),",") Then
							  InputStr=""
							  If replace(lcase(SQL(2,K)),"&","")="provincecity" Then
								 InputStr="<script language=""javascript"" src=""../plus/area.asp""></script>"
							  Else
							  Select Case SQL(1,K)
								Case 2:InputStr="<textarea style=""width:" & SQL(4,K) & "px;height:" & SQL(5,K) & "px;"" rows=""5"" class=""textbox"" name=""" & SQL(2,K) & """>" &SQL(3,K) & "</textarea>"
								Case 3,11
								  If SQL(1,K)=11 Then
					               InputStr= "<select style=""width:" & SQL(4,K) & """ name=""" & SQL(2,K) & """ onchange=""fill" & SQL(2,K) &"(this.value)""><option value=''>---请选择---</option>"
								  Else
								   InputStr="<select style=""width:" & SQL(4,K) & """ name=""" & SQL(2,K) & """>"
								  End If
								  O_Arr=Split(SQL(6,K),vbcrlf): O_Len=Ubound(O_Arr)
								  For N=0 To O_Len
									 F_V=Split(O_Arr(N),"|")
									 If Ubound(F_V)=1 Then
										O_Value=F_V(0):O_Text=F_V(1)
									 Else
										O_Value=F_V(0):O_Text=F_V(0)
									 End If						   
									 If Trim(SQL(3,K))=O_Value Then
										InputStr=InputStr & "<option value=""" & O_Value& """ selected>" & O_Text & "</option>"
									 Else
										InputStr=InputStr & "<option value=""" & O_Value& """>" & O_Text & "</option>"
									 End If
								  Next
									InputStr=InputStr & "</select>"
									'联动菜单
									If SQL(1,K)=11  Then
										Dim JSStr
										InputStr=InputStr &  GetLDMenuStr("",101,SQL,SQL(2,k),JSStr) & "<script type=""text/javascript"">" &vbcrlf & JSStr& vbcrlf &"</script>"
									End If
								Case 6
									 O_Arr=Split(SQL(6,K),vbcrlf): O_Len=Ubound(O_Arr)
									 If O_Len>1 And Len(SQL(6,K))>50 Then BrStr="<br>" Else BrStr=""
									 For N=0 To O_Len
									  IF O_Arr(N)<>"" Then
										F_V=Split(O_Arr(N),"|")
										If Ubound(F_V)=1 Then
										 O_Value=F_V(0):O_Text=F_V(1)
										Else
										 O_Value=F_V(0):O_Text=F_V(0)
										End If
										If Trim(SQL(3,K))=O_Value Then
											InputStr=InputStr & "<input type=""radio"" name=""" & SQL(2,K) & """ value=""" & O_Value& """ checked>" & O_Text & BRStr
										Else
											InputStr=InputStr & "<input type=""radio"" name=""" & SQL(2,K) & """ value=""" & O_Value& """>" & O_Text & BRStr
										 End If
									  End If
									 Next
							  Case 7
									O_Arr=Split(SQL(6,K),vbcrlf): O_Len=Ubound(O_Arr)
									 For N=0 To O_Len
										  F_V=Split(O_Arr(N),"|")
										  If Ubound(F_V)=1 Then
											O_Value=F_V(0):O_Text=F_V(1)
										  Else
											O_Value=F_V(0):O_Text=F_V(0)
										  End If						   
										  If KS.FoundInArr(Trim(SQL(3,K)),O_Value,",")=true Then
												 InputStr=InputStr & "<input type=""checkbox"" name=""" & SQL(2,K) & """ value=""" & O_Value& """ checked>" & O_Text
										 Else
										  InputStr=InputStr & "<input type=""checkbox"" name=""" & SQL(2,K) & """ value=""" & O_Value& """>" & O_Text
										 End If
								   Next
							  Case 10
							        on error resume next

									InputStr=InputStr & "<textarea style=""display:none"" id=""" & SQL(2,K) &""" name=""" & SQL(2,K) &""">"& Server.HTMLEncode(SQL(3,K)) &"</textarea><script type=""text/javascript"">CKEDITOR.replace('" & SQL(2,K) &"', {width:""" & SQL(4,K) &""",height:""" & SQL(5,K) & """,toolbar:""" & SQL(7,K) & """,filebrowserBrowseUrl :""Include/SelectPic.asp?from=ckeditor&Currpath="& KS.GetUpFilesDir() &""",filebrowserWindowWidth:650,filebrowserWindowHeight:290});</script>"

									
													
							  Case Else:InputStr="<input type=""text"" class=""textbox"" style=""width:" & SQL(4,K) & """ name=""" & SQL(2,K) & """ value=""" & SQL(3,K) & """>"
							  End Select
							  End If
							  If SQL(8,K)="1" Then 
									  InputStr=InputStr & " <select name=""" & SQL(2,K) & "_Unit"" id=""" & SQL(2,K) & "_Unit"">"
									  If Not KS.IsNul(SQL(9,k)) Then
									   Dim KK,UnitOptionsArr:UnitOptionsArr=Split(SQL(9,k),vbcrlf)
									   For KK=0 To Ubound(UnitOptionsArr)
										  InputStr=InputStr & "<option value='" & UnitOptionsArr(KK) & "'>" & UnitOptionsArr(KK) & "</option>"                 
									   Next
									  End If
									  InputStr=InputStr & "</select>"
							End If
							  
							  'if SQL(1,K)=9 Then InputStr=InputStr & "<div><iframe id='UpPhotoFrame' name='UpPhotoFrame' src='User_UpFile.asp?Type=Field&FieldID=" & SQL(0,K) & "&ChannelID=101' frameborder=0 scrolling=no width='100%' height='26'></iframe></div>"
				              Template=Replace(Template,"[@NoDisplay(" & SQL(2,K) & ")]","")
							  Template=Replace(Template,"[@" & replace(SQL(2,K),"&","") & "]",InputStr)
							 End If
						   Next
						End If	
							Response.Write Template
			  
			  %>			  </td>
			</tr>
			
			<TR class="tdbg"> 
			  <TD height="30" colspan="4" style="text-align:center"><input name="Action" type="hidden" id="Action" value="SaveAdd"> 
			  <input name=Submit class='button'  type=submit id="Submit" value="&nbsp;确定添加&nbsp;" > </TD>
			</TR>
		</form>
	    </TABLE>
		<%
		end sub
		
		Sub CheckUserName()
		  Dim UserName:UserName=KS.G("UserName")
		  If UserName="" Then
		   Response.Write "<script>alert('请输入用户名称!');window.close();</script>"
		  Else
		   If Conn.Execute("select top 1 userid from ks_user where username='" & UserName & "'").eof Then
		    Response.Write "<script>alert('恭喜，用户" & UserName & "可以使用！');window.close();</script>"
		   Else
		    Response.Write "<script>alert('对不起，用户" & UserName & "不可以使用，请重输！');window.close();</script>"
		   End If
		  End If
		 End Sub

		
		sub SaveAdd()
		    dim UserID,UserName,RealName,Password,PwdConfirm,Question,Answer,Sex,Email,HomePage,QQ,MSN,GroupID,locked,DataCount,ChargeType,Point,Money,BeginDate,Edays,province,clubgradeid,gradeid,GradeTitle
			dim rsUser,sqlUser
			dim OfficeTel,Address,Sign
			dim IDCard,BirthDay,City,Zip,ICQ,UC
			Dim SchoolAge,UserWorking,HomeTel,Mobile
			Action=Trim(request("Action"))
			Password  = KS.S("Password")
			PwdConfirm= KS.S("PwdConfirm")
			Question  = Trim(request("Question"))
			Answer    = Trim(request("Answer"))
			Sex       = Trim(Request("Sex"))
			Email     = Trim(request("Email"))
			HomePage  = Trim(request("HomePage"))
			QQ        = Trim(request("QQ"))
			MSN       = Trim(request("MSN"))
			ICQ       =KS.G("ICQ")
			UC        =KS.G("UC")
			GroupID   = Trim(request("GroupID"))
			locked    = Trim(request("locked"))
			ChargeType=Trim(request("ChargeType"))
			Point=KS.ChkClng(Trim(request("Point")))
			Money=KS.ChkClng(Trim(request("Money")))
			BeginDate=Trim(request("BeginDate"))
			Edays=    KS.ChkClng(Trim(request("Edays")))
			province= KS.G("province")
			city=     KS.G("city")
			UserName=Trim(request("UserName"))
			RealName=Trim(request("RealName"))
			OfficeTel=Trim(request("OfficeTel"))
			Address=Trim(request("Address"))
			Sign=Trim(request("Sign"))
			IDCard=Trim(request("IDCard"))
			BirthDay=Trim(request("BirthDay"))
			Zip=Trim(Request("Zip"))
			HomeTel=Trim(request("HomeTel"))
			Mobile=Trim(request("Mobile"))
            clubgradeid=KS.ChkClng(Request("clubgradeid"))
			gradeid=KS.ChkClng(Request("gradeid"))
			GradeTitle=Conn.Execute("Select top 1 UserTitle from KS_AskGrade Where GradeID=" & gradeid)(0)

			if Question="" then
				'founderr=true
				'errmsg=errmsg & "<br><li>密码提示问题不能为空</li>"
			end if
			if Email="" then
				founderr=true
				errmsg=errmsg & "<br><li>Email不能为空</li>"
			else
				if KS.IsValidEmail(Email)=false then
					errmsg=errmsg & "<br><li>您的Email有错误</li>"
					founderr=true
				end if
			end if
			 Dim EmailMultiRegTF:EmailMultiRegTF=KS.ChkClng(KS.Setting(28))
				If EmailMultiRegTF=0 Then
					Dim EmailRSCheck:Set EmailRSCheck = Conn.Execute("select top 1 UserID from KS_User where Email='" & Email & "'")
					If Not (EmailRSCheck.BOF And EmailRSCheck.EOF) Then
						EmailRSCheck.Close:Set EmailRSCheck = Nothing
						Response.Write("<script>alert('您注册的Email已经存在！请更换Email再试试！');history.back();</script>")
						Exit Sub
					End If
					EmailRSCheck.Close:Set EmailRSCheck = Nothing
			 End If
		
			
			if GroupID="" then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>请指定用户级别！</li>"
			else
				GroupID=CLng(GroupID)
			end if
			if locked<>0 then locked=1
			if ChargeType="" then
				ChargeType=1
			else
				ChargeType=Clng(ChargeType)
			end if
			if BeginDate="" then
				BeginDate=Now()
			else
				BeginDate=Cdate(BeginDate)
			end if
			
			if BirthDay<>"" then
			    BirthDay=Split(BirthDay," ")(0)
				if Not IsDate(BirthDay) then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>出生日期错误！</li>"
				end if
			Else
			   Birthday=now
			end if
			if IDCard<>"" then
				if len(Cstr(IDCard))<15 then
					FoundErr=True
					ErrMsg=ErrMsg & "<br><li>身份证号码错误！</li>"
				end if
			end if
			Set rsUser=Server.CreateObject("Adodb.RecordSet")
			sqlUser="select top 1 * from KS_User where username='" & username & "'"
			rsUser.Open sqlUser,Conn,1,3
			if not rsUser.Eof Then
			  rsUser.Close:Set rsUser=nothing
					FoundErr=True
					ErrMsg=ErrMsg & "<br><li>用户名已存在！</li>"
			End If
			
		 Dim SQL,K
		 Dim FieldsList:FieldsList=LFCls.GetSingleFieldValue("Select top 1 FormField From KS_UserForm Where ID=" & KS.U_G(GroupID,"formid"))
		 If FieldsList="" Then FieldsList="0"
	     Set RS = Server.CreateObject("ADODB.RECORDSET")
		 If KS.FilterIDs(FieldsList)="" Then
		 RS.Open "Select FieldName,MustFillTF,Title,FieldType,ShowUnit,UnitOptions,ParentFieldName From KS_Field Where ChannelID=101 and (ParentFieldName<>'0' and ParentFieldName is not null)",conn,1,1
		 Else
		 RS.Open "Select FieldName,MustFillTF,Title,FieldType,ShowUnit,UnitOptions,ParentFieldName From KS_Field Where ChannelID=101 and (FieldID In(" & KS.FilterIDs(FieldsList) & ") or (ParentFieldName<>'0' and ParentFieldName is not null))",conn,1,1
		 End If
		 If Not RS.Eof Then SQL=RS.GetRows(-1)
		 RS.Close
		 If IsArray(SQL) Then
		  For K=0 To UBound(SQL,2)
		    If SQL(6,K)="0" Then
		  	  If SQL(1,K)="1" Then 
			     if lcase(SQL(0,K))<>"province&city" and KS.S(SQL(0,K))="" then
					 FoundErr=True
					 ErrMsg=ErrMsg & "<br><li>" & SQL(2,K) & "必须填写!</li>"
				 elseif lcase(SQL(0,K))="province&city" and (KS.S("province")="" or ks.s("city")="") then
					 FoundErr=True
					 ErrMsg=ErrMsg & "<br><li>" & SQL(2,K) & "必须选择!</li>"
				 end if
			   End If

			   
			   If SQL(3,K)="4" And Not Isnumeric(KS.S(SQL(0,K))) Then
			     FoundErr=True
				 ErrMsg=ErrMsg & "<br><li>" & SQL(2,K) & "必须填写数字!</li>"
			   End If
			   If SQL(3,K)="5" And Not IsDate(KS.S(SQL(0,K))) Then 
			     FoundErr=True
				 ErrMsg=ErrMsg & "<br><li>" & SQL(2,K) & "必须填写正确的日期!</li>"
			   End If
			   If SQL(3,K)="8" And Not KS.IsValidEmail(KS.S(SQL(0,K))) and SQL(1,K)="1" Then 
			     FoundErr=True
				 ErrMsg=ErrMsg & "<br><li>" & SQL(2,K) & "必须填写正确的Email格式!</li>"
			   End If 
		  End If
		Next
       End If

			if founderr=true then exit sub
            rsUser.AddNew		
			rsUser("UserFace")=KS.GetDomain & "Images/Face/boy.jpg"	
			rsUser("UserName")=UserName
			rsUser("RealName")=RealName
			rsUser("Password")=MD5(KS.R(Password),16)
			rsUser("Question")=Question
			rsUser("Answer")=Answer
			rsUser("Email")=Email
			rsUser("HomePage")=HomePage
			rsUser("Sex")=Sex
			rsUser("GroupID")=GroupID
			rsUser("locked")=locked
			rsUser("ChargeType")=ChargeType
			rsUser("Point")=Point
			rsUser("Money")=Money
			rsUser("BeginDate")=BeginDate
			rsUser("Edays")=Edays
			rsUser("Sign")=Sign
			rsUser("Birthday")=Birthday
			rsUser("IDCard")=IDCard
			rsUser("province")=province
			rsUser("City")=City
			rsUser("Address")=Address
			rsUser("Zip")=Zip
			rsUser("MSN")=MSN
			rsUser("QQ")=QQ
			rsUser("ICQ")=ICQ
			rsUser("UC")=UC
			rsUser("HomeTel") = HomeTel
			rsUser("Mobile") = Mobile
			rsUser("OfficeTel")=OfficeTel
			rsUser("LastLoginIP")=KS.GetIP()
			rsUser("logintimes")=1
			rsUser("UserType")=KS.ChkClng(KS.U_G(GroupID,"usertype"))
			rsUser("lastlogintime")=now()
			rsUser("RegDate")=Now
			rsUser("clubgradeid")=clubgradeid
			rsUser("ClubSpecialPower")=KS.ChkClng(KS.A_G(clubgradeid,"special"))
			rsUser("GradeTitle")=GradeTitle
			rsUser("GradeID")=GradeID

				 '自定义字段
			If IsArray(SQL) Then
				 For K=0 To UBound(SQL,2)
				  If left(Lcase(SQL(0,K)),3)="ks_" Then
				   rsUser(SQL(0,K))=KS.S(SQL(0,K))
				  End If
				  If SQL(4,K)="1" Then
				   RSUser(SQL(0,K)&"_Unit")=KS.S(SQL(0,K)&"_Unit")
				  End If
				 Next
			End If
			rsUser.update
			rsUser.Close
			set rsUser=Nothing
			Call KS.Alert("恭喜您，用户添加成功！",ComeUrl)
		End Sub
		
		
		Sub Modify()
			dim rsUser,sqlUser,sSex,ChargeType,GroupID
			UserID=KS.ChkClng(UserID)
			GroupID=KS.ChkClng(KS.S("GroupID"))
			if UserID=0 then Response.Write("<script>alert('参数不足！');history.back();</script>")
			sqlUser="select top 1 * from KS_User where UserID=" & UserID
			Set rsUser=Server.CreateObject("ADODB.RECORDSET")
			rsUser.Open sqlUser,conn,1,1
			if rsUser.bof and rsUser.eof then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>找不到指定的用户！</li>"
				rsUser.close:set rsUser=Nothing
				exit sub
			end if
			If GroupID=0 Then GroupID=rsUser("GroupID")
		ChargeType=rsUser("ChargeType")
		%>
		<SCRIPT language=javascript>
		function CheckFrom()
		{
		  if(document.myform.UserName.value=="")
			{
			  alert("用户名不能为空！");
			  document.myform.UserName.focus();
			  return false;
			}
		
		  if(document.myform.Question.value=="")
			{
			  alert("密码问题不能为空！");
			  document.myform.Question.focus();
			  return false;
			}
		  if(document.myform.Answer.value=="")
			{
			  alert("密码答案不能为空！");
			  document.myform.Answer.focus();
			  return false;
			}
		  if(document.myform.Email.value==""){
			  alert("用户Email不能为空！");
			  document.myform.Email.focus();
			  return false;
			}

		}
		</script>
		<script type="text/javascript" src="../editor/ckeditor.js" mce_src="../editor/ckeditor.js"></script>
		<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1">
		<FORM name="myform" action="KS.User.asp" method="post">
			<TR class="sort">
			  <TD height="22" colspan="4" align="center"><span style="margin-top:10px;text-align:center;font-weight:bold">修改注册用户信息</span></TD>
		  </TR>
			<TR class="tdbg"> 
				<TD width="80" height="25" align="right" class="clefttitle"><strong>用户名称：</strong></TD>
			  <TD> <Input class="textbox" Name="UserName" type=text size=30 Value="<%=rsUser("UserName")%>" readonly> <font color="red">*</font></TD>
				<TD width="80" height="25" align="right" class="clefttitle"><strong>用户邮箱：</strong></TD>
				<TD><input class="textbox" name="Email" type="text" size="30" value="<%=rsUser("Email")%>" /></TD>
			</TR>
				<TR class="tdbg"> 
					<TD width="80" height="25" align="right" class="clefttitle"><strong>用户密码：</strong></TD>
				  <TD height="25" class=tips><INPUT class="textbox" type="password" name="Password" value="" size="20" maxLength="52"> 如果不想修改，请留空</TD>
					<TD width="80" height="25" align="right" class="clefttitle"><strong>手机号码：</strong></TD>
					<TD height="25"><INPUT class="textbox" type="text" size=30 name="Mobile" value="<%=rsUser("Mobile")%>" maxLength="14"><br><font color="#FF6600"></font></TD>
				</TR>
			<TR  class="tdbg"> 
				<TD width="80" height="25" align="right" class="clefttitle"><strong>密码问题：</strong></TD>
				<TD><INPUT class="textbox" type=text maxLength=50 size=30 name="Question" value="<%=rsUser("Question")%>"> 
			  <font color="#FF6600">*</font> </TD>
				<TD width="80" height="25" align="right" class="clefttitle"><strong>问题答案：</strong></TD>
				<TD><INPUT class="textbox" type=text maxLength=20 size=30 name="Answer" value="<%=rsUser("Answer")%>"></TD>
			</TR>
			<TR class="tdbg"> 
				<TD width="80" height="25" align="right"  class="clefttitle"><strong>用 户 组：</strong></TD>
			  <TD height="25"> 
			  <select name="GroupID" id="GroupID" onchange="location.href='?Action=Modify&Groupid='+this.value+'&UserID=<%=UserID%>';">
				<%=KS.GetUserGroup_Option(GroupID)%>
			  </select>
			  论坛头衔
			  <select name="clubgradeid">
			  <%KS.LoadAskGrade
			  dim node,xml,master,masterarr,i
			   set xml=Application(KS.SiteSN&"_AskGrade")
			   If IsObject(XML) Then
			     
			    for each node in xml.documentelement.selectnodes("row[@typeflag='1']")
				     if trim(node.selectsinglenode("@gradeid").text)=trim(rsUser("ClubGradeID")) Then
					  response.write "<option value='" & node.selectsinglenode("@gradeid").text &"' selected>" & node.selectsinglenode("@usertitle").text & "</option>"
					 else
					  response.write "<option value='" & node.selectsinglenode("@gradeid").text &"'>" & node.selectsinglenode("@usertitle").text & "</option>"
					 end if
			    next
			   end if
			  %>
			  </select>
			  问吧头衔
			  <select name="gradeid">
			  <%
			   set xml=Application(KS.SiteSN&"_AskGrade")
			   If IsObject(XML) Then
			     
			    for each node in xml.documentelement.selectnodes("row[@typeflag='0']")
				     if trim(node.selectsinglenode("@gradeid").text)=trim(rsUser("GradeID")) Then
					  response.write "<option value='" & node.selectsinglenode("@gradeid").text &"' selected>" & node.selectsinglenode("@usertitle").text & "</option>"
					 else
					  response.write "<option value='" & node.selectsinglenode("@gradeid").text &"'>" & node.selectsinglenode("@usertitle").text & "</option>"
					 end if
			    next
			   end if
			  %>
			  </select>
			  </TD>
			  <TD height="25" align="right" class="clefttitle"><strong>用户状态：</strong></TD>
			  <TD height="25"><input type="radio" name="locked" value="0" <%if rsUser("locked")=0 then Response.Write "checked"%> />
正常&nbsp;&nbsp;
<input type="radio" name="locked" value="1" <%if rsUser("locked")=1 then Response.Write "checked"%> />
锁定<input type="radio" name="locked" value="2" <%if rsUser("locked")=2 then Response.Write "checked"%> />
待审核<input type="radio" name="locked" value="3" <%if rsUser("locked")=3 then Response.Write "checked"%> />
待激活</TD>
			</TR>
			
			<TR class="tdbg">
				<TD width="80" height="25" align="right" class="clefttitle"><strong>计费方式：</strong></TD>
				<TD><input name="ChargeType" type="radio" value="1" <%if ChargeType=1 then Response.Write " checked"%>>
				扣点数<font color="#0066CC">（推荐）</font>
				<input type="radio" name="ChargeType" value="2" <%if ChargeType=2 then Response.Write " checked"%>>
			  有效期
			  <input type="radio" name="ChargeType" value="3" <%if ChargeType=3 then Response.Write " checked"%> />
无限期</TD>
			    <TD align="right" class="clefttitle"><strong>头像地址：</strong></TD>
			    <TD><input class="textbox" type="text" maxlength="255" size="30" name="UserFace" value="<%=rsUser("UserFace")%>" /></TD>
			</TR>
			<TR  class="tdbg">
				<TD width="80" height="25" align="right" class="clefttitle"><strong>有效期限：</strong></TD>
				<TD height="50" Colspan=3>开始日期：
				<input  class="textbox" name="BeginDate" type="text" id="BeginDate" readonly value="<%=FormatDateTime(rsUser("BeginDate"),2)%>" size="20" maxlength="20"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;有 效 期：
			  <input  class="textbox" name="EDays" readonly type="text" id="EDays" value="<%=rsUser("EDays")%>" size="10" maxlength="10">
			 天
			  <br>
				若超过此期限，则用户不能阅读收费内容此功能只有当计费方式为“有效期限”时才有效			  </TD>
			</TR>
			<TR class="tdbg"> 
				<TD width="80" height="25" align="right" class="clefttitle"><strong>用户点数：</strong></TD>
				<TD> <Input class="textbox" readonly Name="Point" type=text size=30 Value="<%=rsUser("Point")%>"></TD>
				<TD width="80" height="25" align="right" class="clefttitle"><strong>用户资金：</strong></TD>
			  <TD> <Input class="textbox" readonly Name="Money" type=text size=8 Value="<%=rsUser("Money")%>">元</TD>
			</TR>
						
			<tr class="sort">
			  <td colspan="4" align="center">====自定义选项====</td>
			</tr>
			<tr class="tdbg">
			  <td colspan="4">
			  <%
			            
					Dim Template:Template=LFCls.GetSingleFieldValue("Select top 1 Template From KS_UserForm Where ID=" & KS.ChkClng(KS.U_G(GroupID,"formid")))
							
						   Dim FieldsList:FieldsList=LFCls.GetSingleFieldValue("Select top 1 FormField From KS_UserForm Where ID=" & KS.ChkClng(KS.U_G(GroupID,"formid")))
						   Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
						   RS.Open "Select FieldID,FieldType,FieldName,DefaultValue,Width,Height,Options,EditorType,ShowUnit,UnitOptions from KS_Field Where ChannelID=101 Order By OrderID",conn,1,1
						   Dim SQL,K,N,InputStr,O_Arr,O_Len,F_V,O_Value,O_Text,BRStr
						   If Not RS.EOf Then SQL=RS.GetRows(-1):RS.Close():Set RS=Nothing
						If IsArray(SQL) Then
						   For K=0 TO Ubound(SQL,2)
							 If KS.FoundInArr(FieldsList,SQL(0,k),",") Then
							  InputStr=""
							  If lcase(replace(SQL(2,K),"&",""))="provincecity" Then
							  	  InputStr="<script>try{setCookie(""pid"",'" & rsuser("province") & "');}catch(e){}</script>" & vbcrlf
								 InputStr=InputStr & "<script src='../plus/area.asp'></script><script language=""javascript"">" &vbcrlf
								 If rsUser("Province")<>"" And Not ISNull(rsUser("Province")) Then
						         InputStr=InputStr & "$('#Province').val('" & rsUser("province") &"');" &vbcrlf
								 End If
						         If rsUser("City")<>"" And Not ISNull(rsUser("City")) Then
								  InputStr=InputStr & "$('#City').val('" & rsUser("City") & "');" &Vbcrlf
						         end if
						          InputStr=InputStr & "</script>" &vbcrlf
							  Else
							  Select Case SQL(1,K)
								Case 2:InputStr="<textarea style=""width:" & SQL(4,K) & "px;height:" & SQL(5,K) & "px;"" rows=""5"" class=""textbox"" name=""" & SQL(2,K) & """>" &rsUser(SQL(2,K)) & "</textarea>"
								Case 3,11
								If SQL(1,K)=11 Then
					               InputStr= "<select style=""width:" & SQL(4,K) & """ name=""" & SQL(2,K) & """ onchange=""fill" & SQL(2,K) &"(this.value)""><option value=''>---请选择---</option>"
								  Else
								   InputStr="<select style=""width:" & SQL(4,K) & """ name=""" & SQL(2,K) & """>"
								  End If
								  O_Arr=Split(SQL(6,K),vbcrlf): O_Len=Ubound(O_Arr)
								  For N=0 To O_Len
									 F_V=Split(O_Arr(N),"|")
									 If Ubound(F_V)=1 Then
										O_Value=F_V(0):O_Text=F_V(1)
									 Else
										O_Value=F_V(0):O_Text=F_V(0)
									 End If						   
									 If Trim(rsUser(SQL(2,K)))=O_Value Then
										InputStr=InputStr & "<option value=""" & O_Value& """ selected>" & O_Text & "</option>"
									 Else
										InputStr=InputStr & "<option value=""" & O_Value& """>" & O_Text & "</option>"
									 End If
								  Next
									InputStr=InputStr & "</select>"
									'联动菜单
									If SQL(1,K)=11  Then
										Dim JSStr
										InputStr=InputStr &  GetLDMenuStr(RSUser,101,SQL,SQL(2,k),JSStr) & "<script type=""text/javascript"">" &vbcrlf & JSStr& vbcrlf &"</script>"
									End If
								Case 6
									 O_Arr=Split(SQL(6,K),vbcrlf): O_Len=Ubound(O_Arr)
									 If O_Len>1 And Len(SQL(6,K))>50 Then BrStr="<br>" Else BrStr=""
									 For N=0 To O_Len
										F_V=Split(O_Arr(N),"|")
										If Ubound(F_V)=1 Then
										 O_Value=F_V(0):O_Text=F_V(1)
										Else
										 O_Value=F_V(0):O_Text=F_V(0)
										End If
										If Trim(rsUser(SQL(2,K)))=O_Value Then
											InputStr=InputStr & "<input type=""radio"" name=""" & SQL(2,K) & """ value=""" & O_Value& """ checked>" & O_Text & BRStr
										Else
											InputStr=InputStr & "<input type=""radio"" name=""" & SQL(2,K) & """ value=""" & O_Value& """>" & O_Text & BRStr
										 End If
									 Next
							  Case 7
									O_Arr=Split(SQL(6,K),vbcrlf): O_Len=Ubound(O_Arr)
									 For N=0 To O_Len
										  F_V=Split(O_Arr(N),"|")
										  If Ubound(F_V)=1 Then
											O_Value=F_V(0):O_Text=F_V(1)
										  Else
											O_Value=F_V(0):O_Text=F_V(0)
										  End If						   
										  If KS.FoundInArr(Trim(rsUser(SQL(2,K))),O_Value,",")=true Then
												 InputStr=InputStr & "<input type=""checkbox"" name=""" & SQL(2,K) & """ value=""" & O_Value& """ checked>" & O_Text
										 Else
										  InputStr=InputStr & "<input type=""checkbox"" name=""" & SQL(2,K) & """ value=""" & O_Value& """>" & O_Text
										 End If
								   Next
							  Case 10
							        Dim H_Value:H_Value=rsUser(SQL(2,K))
									If IsNull(H_Value) Then H_Value=" "
									InputStr=InputStr & "<textarea  style=""display:none"" id=""" & SQL(2,K) &""" name=""" & SQL(2,K) &""">"& Server.HTMLEncode(H_Value) &"</textarea><script type=""text/javascript"">CKEDITOR.replace('" & SQL(2,K) &"', {width:""" & SQL(4,K) &""",height:""" & SQL(5,K) & """,toolbar:""" & SQL(7,K) & """,filebrowserBrowseUrl :""Include/SelectPic.asp?from=ckeditor&Currpath="& KS.GetUpFilesDir() &""",filebrowserWindowWidth:650,filebrowserWindowHeight:290});</script>"
									
							  Case Else
								  InputStr="<input type=""text"" class=""textbox"" style=""width:" & SQL(4,K) & """ name=""" & SQL(2,K) & """ value=""" & rsUser(SQL(2,K)) & """>"

							  End Select
							  End If

							  If SQL(8,K)="1" Then 
								  InputStr=InputStr & " <select name=""" & SQL(2,K) & "_Unit"" id=""" & SQL(2,K) & "_Unit"">"
								  If Not KS.IsNul(SQL(9,k)) Then
								   Dim KK,UnitOptionsArr:UnitOptionsArr=Split(SQL(9,k),vbcrlf)
								   For KK=0 To Ubound(UnitOptionsArr)
								      If Trim(RSUser(SQL(2,K) & "_Unit"))=Trim(UnitOptionsArr(KK)) Then
									  InputStr=InputStr & "<option value='" & UnitOptionsArr(KK) & "' selected>" & UnitOptionsArr(KK) & "</option>"                 
									  Else
									  InputStr=InputStr & "<option value='" & UnitOptionsArr(KK) & "'>" & UnitOptionsArr(KK) & "</option>"                 
									  End If
								   Next
								  End If
								  InputStr=InputStr & "</select>"
			                  End If
							  
							  'if SQL(1,K)=9 Then InputStr=InputStr & "<div><iframe id='UpPhotoFrame' name='UpPhotoFrame' src='User_UpFile.asp?Type=Field&FieldID=" & SQL(0,K) & "&ChannelID=101' frameborder=0 scrolling=no width='100%' height='26'></iframe></div>"
				              Template=Replace(Template,"[@NoDisplay(" & SQL(2,K) & ")]","")
							  Template=Replace(Template,"[@" & replace(SQL(2,K),"&","") & "]",InputStr)
							  
							 End If
						   Next
					End IF	
							Response.Write Template
			  
			  %>			  </td>
			</tr>
			

			<TR class="tdbg"> 
			  <TD height="30" colspan="4" style="text-align:center"><input name="Action" type="hidden" id="Action" value="SaveModify"> 
			  <input name=Submit class='button'  type=submit id="Submit" value="&nbsp;保存修改结果&nbsp;" > &nbsp;&nbsp;<input type='button' onclick="location.href='KS.User.asp?userid=<%=rsUser("UserID")%>&action=ShowDetail';" value="查看打印" class='button'><input name="UserID" type="hidden" id="UserID" value="<%=rsUser("UserID")%>"></TD>
			</TR>
		</form>
	    </TABLE>
		<%
			rsUser.close:set rsUser=Nothing
		end sub
		
		'取得联动菜单
		   Function GetLDMenuStr(RSU,ChannelID,F_Arr,byVal ParentFieldName,JSStr)
		     Dim OptionS,OArr,I,VArr,V,F,Str
		     Dim RSL:Set RSL=Conn.Execute("Select Top 1 FieldName,Title,Options,Width From KS_Field Where ChannelID=" & ChannelID & " and ParentFieldName='" & ParentFieldName & "'")
			 If Not RSL.Eof Then
			     Str=Str & " <select name='" & RSL(0) & "' id='" & RSL(0) & "' onchange='fill" & RSL(0) & "(this.value)' style='width:" & RSL(3) & "px'><option value=''>--请选择--</option>"
				 JSStr=JSStr & "var sub" &ParentFieldName & " = new Array();"
				  Options=RSL(2)
				  OArr=Split(Options,Vbcrlf)
				  For I=0 To Ubound(OArr)
				    Varr=Split(OArr(i),"|")
					If Ubound(Varr)=1 Then 
					 V=Varr(0):F=Varr(1)
					Else
					 V=trim(OArr(i))
					 F=trim(OArr(i))
					End If
				    JSStr=JSStr & "sub" & ParentFieldName&"[" & I & "]=new Array('" & V & "','" & F & "')" &vbcrlf
				  Next
				 Str=Str & "</select>"
				 JSStr=JSStr & "function fill"& ParentFieldName&"(v){" &vbcrlf &_
							   "$('#"& RSL(0)&"').empty();" &vbcrlf &_
							   "$('#"& RSL(0)&"').append('<option value="""">--请选择--</option>');" &vbcrlf &_
							   "for (i=0; i<sub" &ParentFieldName&".length; i++){" & vbcrlf &_
							   " if (v==sub" &ParentFieldName&"[i][0]){document.getElementById('" & RSL(0) & "').options[document.getElementById('" & RSL(0) & "').length] = new Option(sub" &ParentFieldName&"[i][1], sub" &ParentFieldName&"[i][1]);}}" & vbcrlf &_
							   "}"
				 If IsObject(RSU) Then
				 Dim DefaultVAL:DefaultVAL=RSU(trim(RSL(0)))
                 If Not KS.IsNul(DefaultVAL) Then
				  str=str & "<script>$(document).ready(function(){fill"&ParentFieldName&"($('select[name=" &ParentFieldName&"] option:selected').val()); $('#"& RSL(0)&"').val('" & DefaultVAL & "');})</script>" &vbcrlf
				 End If
				 End If
				 GetLDMenuStr=str & GetLDMenuStr(RSU,ChannelID,F_Arr,RSL(0),JSStr)
			 Else
			     JSStr=JSStr & "function fill" & ParentFieldName &"(v){}"				 
			 End If
			     
		   End Function
		
		Sub ShowDetail()
			dim rsUser,sqlUser,sSex,ChargeType,GroupID
			UserID=KS.ChkClng(UserID)
			GroupID=KS.ChkClng(KS.S("GroupID"))
			if UserID=0 then Response.Write("<script>alert('参数不足！');history.back();</script>")
			sqlUser="select * from KS_User where UserID=" & UserID
			Set rsUser=Server.CreateObject("ADODB.RECORDSET")
			rsUser.Open sqlUser,conn,1,1
			if rsUser.bof and rsUser.eof then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>找不到指定的用户！</li>"
				rsUser.close:set rsUser=Nothing
				exit sub
			end if
			If GroupID=0 Then GroupID=rsUser("GroupID")
		ChargeType=rsUser("ChargeType")
		%>

		<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1">
			<TR class="sort">
			  <TD height="22" colspan="4" align="center"><span style="margin-top:10px;text-align:center;font-weight:bold">用户详细资料</span></TD>
		  </TR>
			<TR class="tdbg"> 
				<TD height="25" align="center" colspan="4">
				<strong>用户名称：</strong><%=rsUser("UserName")%>&nbsp;&nbsp;<strong>注册时间：</strong><%=rsUser("RegDate")%>&nbsp;&nbsp;<strong>推荐人：</strong><%=rsUser("AllianceUser")%>&nbsp;&nbsp;<strong>可用资金：</strong><%=rsUser("money")%>元&nbsp;&nbsp;<strong>可用点券：</strong><%=rsUser("point")%>点&nbsp;&nbsp;<strong>总积分：</strong><%=rsUser("score")%>分&nbsp;<strong>已消费积分：</strong><%=rsUser("scorehasuse")%>分&nbsp;<strong>可用积分:</strong><%=rsuser("score")-rsUser("scorehasuse")%>分
				</TD>
			</TR>
	
			<tr class="tdbg">
			  <td colspan="4">
			  <%
			            
					Dim Template:Template=LFCls.GetSingleFieldValue("Select Template From KS_UserForm Where ID=" & KS.U_G(GroupID,"formid"))
						   Dim FieldsList:FieldsList=LFCls.GetSingleFieldValue("Select FormField From KS_UserForm Where ID=" & KS.U_G(GroupID,"formid"))
						   Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
						   RS.Open "Select FieldID,FieldType,FieldName,DefaultValue,Width,Options from KS_Field Where ChannelID=101 Order By OrderID",conn,1,1
						   Dim SQL,K,N,InputStr,O_Arr,O_Len,F_V,O_Value,O_Text,BRStr
						   If Not RS.EOf Then SQL=RS.GetRows(-1):RS.Close():Set RS=Nothing
						   For K=0 TO Ubound(SQL,2)
							 If KS.FoundInArr(FieldsList,SQL(0,k),",") Then
							  InputStr=""
							  If (SQL(2,K))="Province&City" Then
								 InputStr=rsUser("Province") & "" &  rsUser("City") & ""
							  ElseIf SQL(1,K)=11 Then 
							    InputStr=rsUser(SQL(2,K)) &GetLDValue(SQL(2,K),RSuSER)
							  Else
							     InputStr=rsUser(SQL(2,K))
							     If KS.IsNul(InputStr) Then InputStr=" "
							  End If
							  Template=Replace(Template,"[@NoDisplay(" & SQL(2,K) & ")]","")
							  Template=Replace(Template,"[@" & replace(SQL(2,K),"&","") & "]",InputStr)

							End If
						   Next
							Template=Replace(Template,"{@NoDisplay}"," style='display:none'")
							Response.Write Template
			  
			  %>			  </td>
			</tr>
			

			<TR class="tdbg"> 
			  <TD height="30" colspan="4" style="text-align:center"> 
			  <input name=Submit class='button'  type="button" onclick="this.style.display='none';document.getElementById('modifybutton').style.display='none';document.getElementById('backbutton').style.display='none';document.getElementById('mt').style.display='none';window.print();" id="Submit" value=" 打 印 " >&nbsp;&nbsp;<input type="button" class="button" name="modifybutton" id="modifybutton" value=" 修 改 " onclick="location.href='KS.User.asp?action=Modify&userid=<%=rsUser("UserID")%>';">&nbsp;&nbsp;<input name='backbutton' id='backbutton' type='button' onclick='history.back();' value=' 返 回 ' class='button'></TD>
			</TR>
	    </TABLE>
		<br><br>
		<%
			rsUser.close:set rsUser=Nothing
	  End Sub
	 
	  Function GetLDValue(ParentFieldName,rsUser)
	    Dim RSL:Set RSL=Conn.Execute("Select Top 1 FieldName,Title,Options,Width From KS_Field Where ChannelID=101 and ParentFieldName='" & ParentFieldName & "'")
	    If Not RSL.Eof Then
		 GetLDValue=" " & rsUser(trim(rsl(0))) & " " & GetLDValue(RSL(0),rsUser)
		End If
	  End Function
		
	sub AddScore()
			dim rsUser,sqlUser
			UserID=KS.ChkClng(UserID)
			if UserID=0 then Response.Write("<script>alert('参数不足！');history.back();</script>")
			Set rsUser=Conn.Execute("select top 1 * from KS_User where UserID=" & UserID)
			if rsUser.bof and rsUser.eof then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>找不到指定的用户！</li>"
				rsUser.close:set rsUser=Nothing
				exit sub
			end if

		%>
		<table width="80%" style="margin-top:10px" border="0" align="center" cellpadding="3" cellspacing="1" class="ctable">
		<FORM name="myform" action="KS.User.asp?ComeUrl=<%=Server.URLEncode(Request.ServerVariables("HTTP_REFERER"))%>" method="post">
			<TR class="sort">
			  <TD height="28" colspan="2" align="center"><b>给 用 户 加 积 分</b></TD>
		   </TR>
			<TR class='tdbg'> 
			  <TD width="25%" height="28" align="right" class="clefttitle"><b>用户名：</b></TD>
			  <TD width="75%"><%=rsUser("UserName")%></TD>
			</TR>
			<TR class='tdbg'> 
			  <TD width="25%" height="28" align="right" class="clefttitle"><strong>用户级别：</strong></TD>
			  <TD width="75%"><%=KS.GetUserGroupName(rsUser("GroupID"))%></TD>
			</TR>
			<TR class='tdbg'> 
			  <TD width="25%" height="28" align="right" class="clefttitle"><strong>当前积分：</strong></TD>
			  <TD width="75%"><%=rsUser("Score")%> 分</TD>
			</TR>

			<TR  class='tdbg'>
			  <TD height="28" align="right" class="clefttitle"><strong>追加积分：</strong></TD>
			  <TD> <input name="Score" type="text" id="Score" value="100" size="10" maxlength="10">
			  分</TD>
			</TR>
			<TR class='tdbg'>
			  <TD height="28" align="right" class="clefttitle"><strong>请输入原因：</strong></TD>
			  <TD> <input name="Reason" type="text" id="Reason" value="积分奖励" size="55"></TD>
			</TR>
			<TR  class='tdbg'>
			  <TD height="28" align="right" class="clefttitle"><strong>站内通知：</strong></TD>
			  <TD> <label><input name="MTF" type="checkbox" value="1" checked>发送短消息通知用户</label>
			  <br/>消息内容:
			  <textarea name="message" style="width:300px;height:60px">您好{$UserName},本站管理员手工给您增加了{$Score}分积分奖励!
			  </textarea>
			  </TD>
			</TR>
			
			<TR class='tdbg'> 
			  <TD height="40" colspan="2" style="text-align:center"><input name="Action" type="hidden" id="Action" value="SaveAddScore"> 
			  <input name=Submit class="button"  type=submit id="Submit" value="&nbsp;确定追加积分&nbsp;" > <input name="UserName" type="hidden" id="UserName" value="<%=rsUser("UserName")%>"></TD>
			</TR>
		</form>
	    </TABLE>
		<%
			rsUser.close : set rsUser=Nothing
		end sub
		
		
	 sub AddMoney()
			dim rsUser,sqlUser
			UserID=KS.ChkClng(UserID)
			if UserID=0 then Response.Write("<script>alert('参数不足！');history.back();</script>")
			Set rsUser=Conn.Execute("select top 1 * from KS_User where UserID=" & UserID)
			if rsUser.bof and rsUser.eof then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>找不到指定的用户！</li>"
				rsUser.close:set rsUser=Nothing
				exit sub
			end if
			if rsUser("ChargeType")=3 Then
			  rsUser.Close:Set rsUser=Nothing
			  Call KS.Alert("无限期用户无需续费操作!",Request.ServerVariables("HTTP_REFERER"))
			  Exit Sub
			End if
		%>
		<table width="80%" style="margin-top:10px" border="0" align="center" cellpadding="3" cellspacing="1" class="ctable">
		<FORM name="myform" action="KS.User.asp?ComeUrl=<%=Server.URLEncode(Request.ServerVariables("HTTP_REFERER"))%>" method="post">
			<TR class="sort">
			  <TD height="28" colspan="2" align="center"><b>用 户 续 费</b></TD>
		   </TR>
			<TR class='tdbg'> 
			  <TD width="25%" height="28" align="right" class="clefttitle"><b>用户名：</b></TD>
			  <TD width="75%"><%=rsUser("UserName")%></TD>
			</TR>
			<TR class='tdbg'> 
			  <TD width="25%" height="28" align="right" class="clefttitle"><strong>用户级别：</strong></TD>
			  <TD width="75%"><%=KS.GetUserGroupName(rsUser("GroupID"))%></TD>
			</TR>
			<TR  class='tdbg'>
			  <TD height="28" align="right" class="clefttitle"><strong>计费方式：</strong></TD>
			  <TD><%
			  if rsUser("ChargeType")=1 then
				Response.Write "扣点数"
			  else
				Response.Write "有效期"
			  end if
			  %>
				<input name="ChargeType" type="hidden" id="ChargeType" value="<%=rsUser("ChargeType")%>">			  </TD>
			</TR>
			<TR class='tdbg'> 
			  <TD width="25%" height="28" align="right" class="clefttitle"><strong>可用资金：</strong></TD>
			  <TD width="75%"><%=rsUser("Money")%>元人民币</TD>
			</TR>
			<%if rsUser("ChargeType")=1 then%>
			<TR  class='tdbg'>
			  <TD height="28" align="right" class="clefttitle"><strong>目前的用户<%=KS.Setting(45)%>：</strong></TD>
			  <TD><%=rsUser("Point")%> <%=KS.Setting(46)%></TD>
			</TR>
			<TR  class='tdbg'>
			  <TD height="28" align="right" class="clefttitle"><strong>追加<%=KS.Setting(45)%>：</strong></TD>
			  <TD> <input name="Point" type="text" id="Point" value="100" size="10" maxlength="10">
			  <%=KS.Setting(46)%></TD>
			</TR>
			<%else%>
			<TR class='tdbg'>
			  <TD height="28" align="right" class="clefttitle"><strong>目前的有效期限信息：</strong></TD>
			  <TD>
			  <%
			  Response.Write "开始计算日期" & FormatDateTime(rsUser("BeginDate"),2) & "&nbsp;&nbsp;&nbsp;&nbsp;有 效 期：" & rsUser("Edays")
			 
				Response.Write "天"
			 
			  Response.Write "<br>"
			  tmpDays=rsUser("Edays")-DateDiff("D",rsUser("BeginDate"),now())
			  if tmpDays>=0 then
				Response.Write "尚有 <font color=blue>" & tmpDays & "</font> 天到期"
			  else
				Response.Write "已经过期 <font color=#ff6600>" & abs(tmpDays) & "</font> 天"
			  end if
			  %>			  </TD>
			</TR>
			<tr class='tdbg'>
			  <td height="60" align="right" class="clefttitle"><strong>追加天数：</strong><br></td>
			  <td>
			  <input name="Edays" type="text" id="Edays" value="100" size="10" maxlength="10">
			  天<br />
			  若目前用户尚未到期，则追加相应天数<br />
若目前用户已经过了有效期，则有效期从续费之日起重新计数。</td>
			</tr>
			<%end if%>
			<tr class='tdbg'>
			  <td height="30" align="right" class="clefttitle"><strong>同时减去：</strong><br></td>
			  <td>
			  <input name="Money" type="text" id="Money" value="100" size="10" maxlength="10"> 元人民币
			  <font color=red>
			  <%if rsUser("ChargeType")=1 then %>
			   资金与点券的默认比率：<%=KS.Setting(43)%>:1
			  <%else%>
			  资金与有效期的默认比率：<%=KS.Setting(44)%>:1
			  <%end if%>
			  </font> 不想扣除资金，请输入0
			  </td>
			</tr>
			<TR class='tdbg'>
			  <TD height="28" align="right" class="clefttitle"><strong>请输入原因：</strong></TD>
			  <TD> <input name="Reason" type="text" id="Reason" value="<%If rsUser("ChargeType")=1 Then Response.Write "续" & KS.Setting(45) & "操作" Else Response.Write "续有效天数操作"%>" size="55"></TD>
			</TR>
			<TR class='tdbg'> 
			  <TD height="40" colspan="2" style="text-align:center"><input name="Action" type="hidden" id="Action" value="SaveAddMoney"> 
			  <input name=Submit class="button"  type=submit id="Submit" value="&nbsp;保存续费结果&nbsp;" > <input name="UserID" type="hidden" id="UserID" value="<%=rsUser("UserID")%>"></TD>
			</TR>
		</form>
	    </TABLE>
		<%
			rsUser.close : set rsUser=Nothing
		end sub
		
		
		
		sub SaveModify()
			dim UserID,UserName,RealName,Password,PwdConfirm,Question,Answer,Sex,Email,HomePage,QQ,MSN,GroupID,locked,DataCount,ChargeType,Point,Money,BeginDate,Edays,province,fax,UserFace,clubgradeid,gradeid,GradeTitle
			dim rsUser,sqlUser
			dim OfficeTel,Address,Sign
			dim IDCard,BirthDay,City,Zip,ICQ,UC
			Dim SchoolAge,UserWorking,HomeTel,Mobile
			Action=Trim(request("Action"))
			UserID=KS.ChkClng(request("UserID"))
			if UserID=0 then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
				exit sub
			end if
			Password  =KS.S("Password")
			PwdConfirm= Trim(request("PwdConfirm"))
			Question  = Trim(request("Question"))
			Answer    = Trim(request("Answer"))
			Sex       = Trim(Request("Sex"))
			Email     = Trim(request("Email"))
			HomePage  = Trim(request("HomePage"))
			QQ        = Trim(request("QQ"))
			MSN       = Trim(request("MSN"))
			ICQ       =KS.G("ICQ")
			UC        =KS.G("UC")
			GroupID   = Trim(request("GroupID"))
			locked    = Trim(request("locked"))
			ChargeType=Trim(request("ChargeType"))
			Point=KS.ChkClng(Trim(request("Point")))
			Money=KS.ChkClng(Trim(request("Money")))
			BeginDate=Trim(request("BeginDate"))
			Edays=    KS.ChkClng(Trim(request("Edays")))
			province= KS.G("province")
			city=     KS.G("city")
			UserName=Trim(request("UserName"))
			RealName=Trim(request("RealName"))
			OfficeTel=Trim(request("OfficeTel"))
			Fax=Trim(Request("Fax"))
			Address=Trim(request("Address"))
			Sign=Trim(request("Sign"))
			IDCard=Trim(request("IDCard"))
			BirthDay=Trim(request("BirthDay"))
			Zip=Trim(Request("Zip"))
			HomeTel=Trim(request("HomeTel"))
			Mobile=Trim(request("Mobile"))
			UserFace=Trim(Request("UserFace"))
			clubgradeid=KS.ChkClng(Request("clubgradeid"))
			gradeid=KS.ChkClng(Request("gradeid"))
			GradeTitle=Conn.Execute("Select top 1 UserTitle from KS_AskGrade Where GradeID=" & gradeid)(0)
			Set rsUser=Server.CreateObject("Adodb.RecordSet")
			sqlUser="select top 1 * from KS_User where UserID=" & UserID
			rsUser.Open sqlUser,Conn,1,3
			if rsUser.bof and rsUser.eof then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>找不到指定的用户！</li>"
				rsUser.close:set rsUser=Nothing
				exit sub
			end if
			if Password<>"" then
				if Instr(Password,"=")>0 or Instr(Password,"%")>0 or Instr(Password,chr(32))>0 or Instr(Password,"?")>0 or Instr(Password,"&")>0 or Instr(Password,";")>0 or Instr(Password,",")>0 or Instr(Password,"'")>0 or Instr(Password,",")>0 or Instr(Password,chr(34))>0 or Instr(Password,chr(9))>0 or Instr(Password,"")>0 or Instr(Password,"$")>0 then
					errmsg=errmsg+"<br><li>密码中含有非法字符，如果你不想修改密码，请保持为空。</li>"
					founderr=true
				end if
			end if

			if Question="" then
				'founderr=true
				'errmsg=errmsg & "<br><li>密码提示问题不能为空</li>"
			end if
			
			if Email="" then
				founderr=true
				errmsg=errmsg & "<br><li>Email不能为空</li>"
			else
				if KS.IsValidEmail(Email)=false then
					errmsg=errmsg & "<br><li>您的Email有错误</li>"
					founderr=true
				end if
			end if
			
			 Dim EmailMultiRegTF:EmailMultiRegTF=KS.ChkClng(KS.Setting(28))
				If EmailMultiRegTF=0 Then
					Dim EmailRSCheck:Set EmailRSCheck = Conn.Execute("select top 1 UserID from KS_User where UserName<>'" & rsUser("UserName") & "' And Email='" & Email & "'")
					If Not (EmailRSCheck.BOF And EmailRSCheck.EOF) Then
						EmailRSCheck.Close:Set EmailRSCheck = Nothing
						Response.Write("<script>alert('您注册的Email已经存在！请更换Email再试试！');history.back();</script>")
						Exit Sub
					End If
					EmailRSCheck.Close:Set EmailRSCheck = Nothing
			 End If
		
			
			if GroupID="" then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>请指定用户级别！</li>"
			else
				GroupID=CLng(GroupID)
			end if
			if locked<>0 then locked=1
			if ChargeType="" then
				ChargeType=1
			else
				ChargeType=Clng(ChargeType)
			end if
			if BeginDate="" then
				BeginDate=Now()
			else
				BeginDate=Cdate(BeginDate)
			end if
			
			if BirthDay<>"" then
			    BirthDay=Split(BirthDay," ")(0)
				if Not IsDate(BirthDay) then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>出生日期错误！</li>"
				end if
			end if

         Dim SQL,K
		 Dim FieldsList:FieldsList=LFCls.GetSingleFieldValue("Select FormField From KS_UserForm Where ID=" & KS.U_G(GroupID,"formid"))
		 If FieldsList="" Then FieldsList="0"
	     Set RS = Server.CreateObject("ADODB.RECORDSET")
		 If FieldsList<>"0" Then
		 RS.Open "Select FieldName,MustFillTF,Title,FieldType,ShowUnit,UnitOptions,ParentFieldName From KS_Field Where ChannelID=101 and (FieldID In(" & KS.FilterIDs(FieldsList) & ") or (ParentFieldName<>'0' and ParentFieldName is not null))",conn,1,1
		 If Not RS.Eof Then SQL=RS.GetRows(-1)
		 RS.Close
		  For K=0 To UBound(SQL,2)
		    If SQL(6,K)="0" Then
			   If SQL(1,K)="1" Then 
			     if lcase(SQL(0,K))<>"province&city" and KS.S(SQL(0,K))="" then
					 FoundErr=True
					 ErrMsg=ErrMsg & "<br><li>" & SQL(2,K) & "必须填写!</li>"
				 elseif lcase(SQL(0,K))="province&city" and (KS.S("province")="" or ks.s("city")="") then
					 FoundErr=True
					 ErrMsg=ErrMsg & "<br><li>" & SQL(2,K) & "必须选择!</li>"
				 end if
			   End If
			   If SQL(3,K)="4" And Not Isnumeric(KS.S(SQL(0,K))) Then
			     FoundErr=True
				 ErrMsg=ErrMsg & "<br><li>" & SQL(2,K) & "必须填写数字!</li>"
			   End If
			   If SQL(3,K)="5" And Not IsDate(KS.S(SQL(0,K))) Then 
			     FoundErr=True
				 ErrMsg=ErrMsg & "<br><li>" & SQL(2,K) & "必须填写正确的日期!</li>"
			   End If
			   If SQL(3,K)="8" And Not KS.IsValidEmail(KS.S(SQL(0,K))) and SQL(1,K)="1" Then 
			     FoundErr=True
				 ErrMsg=ErrMsg & "<br><li>" & SQL(2,K) & "必须填写正确的Email格式!</li>"
			   End If 
		  End If
		 Next
	End If	
			if founderr=true then exit sub
			
			rsUser("RealName")=RealName
			if Password<>"" then rsUser("Password")=MD5(KS.R(Password),16)
			rsUser("Question")=Question
			if Answer<>"" then rsUser("Answer")=Answer
			rsUser("Email")=Email
			rsUser("HomePage")=HomePage
			rsUser("Sex")=Sex
			rsUser("GroupID")=GroupID
			rsUser("locked")=locked
			rsUser("ChargeType")=ChargeType
			'rsUser("Point")=Point
			'rsUser("Money")=Money
			rsUser("BeginDate")=BeginDate
			rsUser("Edays")=Edays
			rsUser("Sign")=Sign
			rsUser("UserFace")=UserFace
			if not isdate(birthday) then
			rsUser("Birthday")=now
			else
			rsUser("Birthday")=Birthday
			end if
			rsUser("IDCard")=IDCard
			rsUser("province")=province
			rsUser("City")=City
			rsUser("Address")=Address
			rsUser("Zip")=Zip
			rsUser("MSN")=MSN
			rsUser("Fax")=Fax
			rsUser("QQ")=QQ
			rsUser("ICQ")=ICQ
			rsUser("UC")=UC
			rsUser("HomeTel") = HomeTel
			rsUser("OfficeTel")=OfficeTel
			rsUser("clubgradeid")=clubgradeid
			rsUser("ClubSpecialPower")=KS.ChkClng(KS.A_G(clubgradeid,"special"))
			rsUser("GradeTitle")=GradeTitle
			rsUser("GradeID")=GradeID
			
			'自定义字段
		If IsArray(SQL) Then
			 For K=0 To UBound(SQL,2)
				If left(Lcase(SQL(0,K)),3)="ks_" Then
				   rsUser(SQL(0,K))=KS.G(SQL(0,K))
				End If
				  If SQL(4,K)="1" Then
				   RSUser(SQL(0,K)&"_Unit")=KS.S(SQL(0,K)&"_Unit")
				  End If
			 Next
	   End If
			rsUser("Mobile") = Mobile
			rsUser.update
			rsUser("UserType")=KS.ChkClng(KS.U_G(rsUser("GroupID"),"usertype"))
			rsUser.Update
			
			
			Dim FieldsXml:Set FieldsXml=LFCls.GetXMLFromFile("SpaceFields")
			If IsObject(FieldsXml) Then
				   	 Dim objNode,i,j,objAtr
					 Set objNode=FieldsXml.documentElement 
					If objNode.Attributes.item(0).Text<>"0" Then
					   If Not Conn.Execute("Select UserName From KS_EnterPrise Where UserName='" &rsUser("UserName") & "'").Eof Then
						 For i=0 to objNode.ChildNodes.length-1 
								set objAtr=objNode.ChildNodes.item(i) 
								on error resume next
								Conn.Execute("UPDATE KS_EnterPrise Set " & objAtr.Attributes.item(0).Text & "='" & rsUser(objAtr.Attributes.item(1).Text) & "' Where UserName='" & rsUser("UserName") & "'")
						 Next
					   End If
					End If
			 End If
				
				'-----------------------------------------------------------------
				'系统整合
				'-----------------------------------------------------------------
				Dim API_KS,API_SaveCookie,SysKey
				If API_Enable Then
					Set API_KS = New API_Conformity
					API_KS.NodeValue "action","update",0,False
					API_KS.NodeValue "username",rsUser("UserName"),1,False
					Md5OLD = 1
					SysKey = Md5(API_KS.XmlNode("username") & API_ConformKey,16)
					Md5OLD = 0
					API_KS.NodeValue "syskey",SysKey,0,False
					API_KS.NodeValue "email",rsUser("Email"),1,False
					API_KS.NodeValue "mobile",rsUser("Mobile"),1,False
					API_KS.NodeValue "homepage",rsUser("homepage"),1,False
					API_KS.NodeValue "address",rsUser("Address"),1,False
					API_KS.NodeValue "zipcode",rsUser("zip"),1,False
					API_KS.NodeValue "qq",rsUser("qq"),1,False
					API_KS.NodeValue "icq",rsUser("icq"),1,False
					API_KS.NodeValue "msn",rsUser("msn"),1,False

					If KS.S("PassWord")<>"" Then
					API_KS.NodeValue "password",KS.R(KS.S("PassWord")),1,False
					End If
					API_KS.SendHttpData
					If API_KS.Status = "1" Then
						Response.Write "<script>alert('" &  API_KS.Message  & "');</script>"
						Exit Sub
					End If
					Set API_KS = Nothing
				End If
				'-----------------------------------------------------------------
			
			If KS.C_S(8,21)="1" Then
				  Conn.Execute("Update KS_GQ Set ContactMan='" & RealName &"',Tel='" &OfficeTel & "',Address='" & Address & "',Zip='" & Zip & "',Fax='" & Fax & "',Homepage='" & HomePage & "' where inputer='" & rsUser("UserName") & "'")
			End If
			rsUser.Close
			set rsUser=Nothing

			
			
			Call KS.Alert("恭喜您，修改成功！请按确定返回！",ComeUrl)
		end sub
		
		sub SaveAddMoney()
			dim UserID,ChargeType,Point,Edays,rsUser,sqlUser,Reason
			Dim Money:Money=KS.G("Money")
			If Not IsNumeric(Money) Then 
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>减去的资金有误！</li>"
				exit sub
			End if
			Action=Trim(request("Action"))
			UserID=KS.ChkClng(request("UserID"))
			if UserID=0 then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
				exit sub
			end if
			ChargeType=Trim(request("ChargeType"))
			Point=KS.ChkClng(Trim(request("Point")))
			Edays=KS.ChkClng(Trim(request("Edays")))
			Reason=KS.G("Reason")
		
			if ChargeType="" then
				ChargeType=1
			else
				ChargeType=Clng(ChargeType)
			end if
			
			if ChargeType=1 and Point=0 then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>请输入要追加的用户点数！</li>"
			end if
			if ChargeType=2 and Edays=0 then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>请输入要追加的天数</li>"
			end if
		    if Reason="" Then
			  FoundErr=True
			  ErrMsg=ErrMsg & " <br><li>请输入操作原因</li>"
			end if
			if founderr=true then exit sub
			Set rsUser=Server.CreateObject("Adodb.RecordSet")
			sqlUser="select top 1 * from KS_User where UserID=" & UserID
			rsUser.Open sqlUser,Conn,1,3
			if rsUser.bof and rsUser.eof then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>找不到指定的用户！</li>"
				rsUser.close:set rsUser=Nothing
				exit sub
			end if
			If Round(rsUser("Money"),2)<Round(Money,2) and money>0 Then
			  FoundErr=True
			  ErrMsg=ErrMsg & " <br><li>该用户的可用资金不足！</li>"
			  rsUser.close:set rsUser=Nothing
			  exit sub
			End If
			'rsUser("Money")=rsUser("Money")-Money
			if ChargeType=1 then
				'rsUser("Money")=rsUser("Money")-Money
			else
				ValidDays=rsUser("Edays")
				tmpDays=ValidDays-DateDiff("D",rsUser("BeginDate"),now())
				if tmpDays>0 then
					rsUser("Edays")=rsUser("Edays")+Edays
				else
					rsUser("BeginDate")=now
					rsUser("Edays")=Edays
				end if
			end if
			rsUser.update
			
			'消费记录
			If Money>0 Then
			 if ChargeType=2 Then
			  Call KS.MoneyInOrOut(rsUser("UserName"),rsUser("RealName"),Money,4,2,now,0,KS.C("AdminName"),"用于兑换有效天数",0,0,0)
			 else
			  Call KS.MoneyInOrOut(rsUser("UserName"),rsUser("RealName"),Money,4,2,now,0,KS.C("AdminName"),"用于兑换点券",0,0,0)
			 end if
			end if
			
			
			if ChargeType=1 then
			 Call KS.PointInOrOut(0,0,rsUser("UserName"),1,Point,KS.C("AdminName"),Reason,0)
			else
			 Call KS.EdaysInOrOut(rsUser("UserName"),1,Edays,KS.C("AdminName"),Reason)
			end if
			rsUser.Close:set rsUser=Nothing
			IF Request("ComeUrl")<>"" Then
			Call KS.Alert("操作成功!",Request("ComeUrl"))
			Else
			Call KS.Alert("操作成功!","KS.User.asp")
			End IF
		end sub
		
		
		sub SaveAddScore()
			dim UserName,Score,MTF,Message,Reason
			Action=Trim(request("Action"))
            UserName=KS.G("UserName")
			Score=KS.ChkClng(Trim(request("Score")))
			Reason=KS.G("Reason")
			MTF=KS.ChkClng(Request("MTF"))
			Message=Request("Message")
		
			
			if Score=0 then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>请输入要追加的积分！</li>"
			end if

		    if Reason="" Then
			  FoundErr=True
			  ErrMsg=ErrMsg & " <br><li>请输入操作原因</li>"
			end if
			if founderr=true then exit sub
			
	
			Call KS.ScoreInOrOut(UserName,1,Score,KS.C("AdminName"),Reason,0,0)
			If Mtf=1 and Not KS.IsNul(Message) Then
			 Message=Replace(Message,"{$UserName}",username)
			 Message=Replace(Message,"{$Score}",score)
			 Call KS.SendInfo(UserName,KS.C("AdminName"),"获得积分通知",Message)
			End If


			IF Request("ComeUrl")<>"" Then
			Call KS.Alert("操作成功!",Request("ComeUrl"))
			Else
			Call KS.Alert("操作成功!","KS.User.asp")
			End IF
		end sub
		
		'添加会员资金
		Sub AddZJ()
		dim rsUser,sqlUser
			UserID=KS.ChkClng(UserID)
			if UserID=0 then Response.Write("<script>alert('参数不足！');history.back();</script>")
			Set rsUser=Conn.Execute("select top 1 * from KS_User where UserID=" & UserID)
			if rsUser.bof and rsUser.eof then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>找不到指定的用户！</li>"
				rsUser.close:set rsUser=Nothing
				exit sub
			end if
		%>
		<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1">
		<FORM name="myform" action="KS.User.asp?ComeUrl=<%=Server.URLEncode(Request.ServerVariables("HTTP_REFERER"))%>" method="post">
			<TR class="sort">
			  <TD colspan="2" align="center"><b>用 户 续 费(增加资金)</b></TD>
		   </TR>
			<TR class='tdbg'> 
			  <TD width="25%" height="28" align="right" class="clefttitle"><b>用户名：</b></TD>
			  <TD width="75%"><%=rsUser("UserName")%></TD>
			</TR>
			<TR class='tdbg'> 
			  <TD width="25%" height="28" align="right" class="clefttitle"><strong>可用资金：</strong></TD>
			  <TD width="75%"><%=rsUser("Money")%> 元</TD>
			</TR>
			<TR class='tdbg'> 
			  <TD width="25%" height="28" align="right" class="clefttitle"><strong>用户级别：</strong></TD>
			  <TD width="75%"><%=KS.GetUserGroupName(rsUser("GroupID"))%></TD>
			</TR>
			<TR  class='tdbg'>
			  <TD height="28" align="right" class="clefttitle"><strong>资金来源：</strong></TD>
			  <TD><label><input name="MoneyType" type="radio" id="ChargeType" checked onclick="document.all.Remark.value='银行汇款';" value="2">银行汇款</label>
			      <label><input name="MoneyType" type="radio" id="ChargeType" onclick="document.all.Remark.value='现金收取';" value="1">其它（如：现金）</label>
		      </TD>
			</TR>
			
			<TR  class='tdbg'>
			  <TD height="28" align="right" class="clefttitle"><strong>操作日期：</strong></TD>
			  <TD><input name="PayTime" type="text" id="PayTime" value="<%=formatdatetime(now,2)%>" size="15" class="textbox"></TD>
			</TR>
			<TR  class='tdbg'>
			  <TD height="28" align="right" class="clefttitle"><strong>操作方式：</strong></TD>
			  <TD><label><input name="InOrOut" type="radio" id="InOrOut" onclick="document.all.Remark.value='续费';" checked  value="1">续费</label>
			      <label><input name="InOrOut" type="radio" id="InOrOut" onclick="document.all.Remark.value='扣费';" value="2">扣费</label>
		      </TD>
			</TR>
			<TR  class='tdbg'>
			  <TD height="28" align="right" class="clefttitle"><strong>金额：</strong></TD>
			  <TD> <input name="Money" type="text" id="Money" value="100" size="15" class="textbox">
			  元</TD>
			</TR>
			
			
			<TR class='tdbg'>
			  <TD height="28" align="right" class="clefttitle"><strong>备注：</strong></TD>
			  <TD> <input name="Remark" type="text" id="Remark" value="银行汇款" size="55"></TD>
			</TR>
			<TR class='tdbg'> 
			  <TD height="40" colspan="2" style="text-align:center"><input name="Action" type="hidden" id="Action" value="SaveAddZJ"> 
			  <input name=Submit  class='button' type=submit id="Submit" value="&nbsp;保存续费结果&nbsp;" > <input name="UserID" type="hidden" id="UserID" value="<%=rsUser("UserID")%>"><input class='button' type='button' value=' 返回 ' onclick='javascript:history.back();'></TD>
			</TR>
		</form>
	    </TABLE>
		<%
			rsUser.close : set rsUser=Nothing
		End Sub
		
		'保存续费
		sub SaveAddZJ()
			dim UserID,MoneyType,Money,PayTime,Remark,sqlUser,rsUser,InOrOut
			Action=Trim(request("Action"))
			UserID=KS.ChkClng(request("UserID"))
			InOrOut=KS.ChkClng(Request("InOrOut"))
			if UserID=0 then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
				exit sub
			end if
			MoneyType=Trim(request("MoneyType"))
			Money=KS.G("Money")
			PayTime=KS.G("PayTime")
			Remark=KS.G("Remark")
            If Not IsDate(PayTime) Then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>日期格式有误</li>"
			end if
			if KS.ChkClng(Money)=0 then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>请输入金额</li>"
			end if
			If Money<0 Then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>金额比须大于0</li>"
			End If
		    if Remark="" Then
			  FoundErr=True
			  ErrMsg=ErrMsg & " <br><li>请输入备注</li>"
			end if
			if founderr=true then exit sub
			Set rsUser=Server.CreateObject("Adodb.RecordSet")
			sqlUser="select top 1 * from KS_User where UserID=" & UserID
			rsUser.Open sqlUser,Conn,1,3
			if rsUser.bof and rsUser.eof then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>找不到指定的用户！</li>"
				rsUser.close:set rsUser=Nothing
				exit sub
			end if
				'rsUser("Money")=rsUser("Money")+Money
			rsUser.update
			
			Call KS.MoneyInOrOut(rsUser("UserName"),rsUser("RealName"),Money,MoneyType,InOrOut,PayTime,"0",KS.C("AdminName"),Remark,0,0,0)
			IF Request("ComeUrl")<>"" Then
			Call KS.Alert("操作成功!",Request("ComeUrl"))
			Else
			Call KS.Alert("操作成功!","KS.User.asp")
			End IF
		end sub
		
		
		sub DelUser()
			if UserID="" then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>请指定要删除的用户</li>"
				exit sub
			end if
			UserID=replace(UserID," ","")
			
		    Dim rsUser:Set rsUser=Conn.Execute("select username from KS_User Where UserID In(" & UserID & ") and GroupID<>1")					
			Do While Not rsUser.Eof
				Conn.Execute("Delete From KS_Blog Where UserName='" & rsUser(0) & "'")
				Conn.Execute("Delete From KS_BlogInfo Where UserName='" & rsUser(0) & "'")
				Conn.Execute("Delete From KS_BlogMusic Where UserName='" & rsUser(0) & "'")
				Conn.Execute("Delete From KS_Enterprise Where UserName='" & rsUser(0) & "'")
				Conn.Execute("Delete From KS_EnterpriseNews Where UserName='" & rsUser(0) & "'")
				Conn.Execute("Delete From KS_UserLog Where UserName='" & rsUser(0) & "'")
			
			
				'-----------------------------------------------------------------
				'系统整合
				'-----------------------------------------------------------------
				Dim API_KS,API_SaveCookie,SysKey
				If API_Enable Then
					Set API_KS = New API_Conformity
					API_KS.NodeValue "action","delete",0,False
					API_KS.NodeValue "username",rsUser("UserName"),1,False
					Md5OLD = 1
					SysKey = Md5(API_KS.XmlNode("username") & API_ConformKey,16)
					Md5OLD = 0
					API_KS.NodeValue "syskey",SysKey,0,False
					API_KS.SendHttpData
					If API_KS.Status = "1" Then
						Response.Write "<script>alert('" &  API_KS.Message  & "');history.back();</script>"
						Exit Sub
					End If
					Set API_KS = Nothing
				End If
				'-----------------------------------------------------------------
				
		  rsUser.MoveNext
		  Loop
		  rsUser.Close
		  Set rsUser=Nothing
			
			Conn.Execute ("Delete From KS_UploadFiles Where channelid=1023 and infoid in(" & UserID &")")
			Conn.Execute ("Delete From KS_UploadFiles Where channelid=1024 and infoid in(" & UserID &")")
			if instr(UserID,",")>0 then
				sql="delete from KS_User where UserID in (" & UserID & ") and GroupID<>1"
			else
				UserID=KS.ChkClng(UserID)
				sql="delete from KS_User where UserID=" & UserID & "  and GroupID<>1"
			end if
			Conn.Execute sql
			Response.Redirect ComeUrl
		end sub
		
		sub locked()
			if UserID="" then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>请选择要锁定的用户</li>"
				exit sub
			end if
			if instr(UserID,",")>0 then
				UserID=replace(UserID," ","")
				sql="Update KS_User set locked=1 where UserID in (" & UserID & ")"
			else
				UserID=KS.ChkClng(UserID)
				sql="Update KS_User set locked=1 where UserID=" & UserID
			end if
			Conn.Execute sql
			Response.Redirect ComeUrl
		end sub
		
		
		sub Unlocked()
			if UserID="" then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>请指定要操作的用户</li>"
				exit sub
			end if
			if instr(UserID,",")>0 then
				UserID=replace(UserID," ","")
				sql="Update KS_User set locked=0 where UserID in (" & UserID & ")"
			else
				UserID=KS.ChkClng(UserID)
				sql="Update KS_User set locked=0 where UserID=" & UserID
			end if
			Conn.Execute sql
			Response.Redirect ComeUrl
		end sub
		
		sub verify(v)
			if UserID="" then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>请选择要审核的用户</li>"
				exit sub
			end if
			if instr(UserID,",")>0 then
				UserID=replace(UserID," ","")
				sql="Update KS_User set locked=" & v & " where UserID in (" & UserID & ")"
			else
				UserID=KS.ChkClng(UserID)
				sql="Update KS_User set locked=" & v & " where UserID=" & UserID
			end if
			Conn.Execute sql
			Response.Redirect ComeUrl
		end sub
		

		
		sub MoveUser()
			Dim RsGroup
			Dim sGroupName,sChargeType,sValidDays,sGroupPoint
			Dim GroupID
			if UserID="" then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>请指定要移动的用户</li>"
				Exit Sub
			end if
			GroupID=KS.ChkClng(request("GroupID"))
			if GroupID=0 then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>请指定目标用户组</li>"
				Exit Sub
			end if
			UserID=replace(UserID," ","")
			Set RsGroup=Conn.Execute("Select GroupName,ChargeType,ValidDays,GroupPoint From KS_UserGroup Where ID="&GroupID&"")
			if Not (RsGroup.Bof and RsGroup.Eof) then
				sGroupName	= RsGroup(0)
				sChargeType	= RsGroup(1)
				sValidDays	= RsGroup(2)
				sGroupPoint	= RsGroup(3)
			else
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>请指定目标用户组</li>"
				Exit Sub
			end if
			RsGroup.Close : Set RsGroup=Nothing
			ErrMsg = "&nbsp;&nbsp;&nbsp;&nbsp;已经成功将选定用户设为“<font color=blue>"&sGroupName&"</font>”！并且按照用户组中预设值给设定了这些用户的计费方式、初始点数、有效期等数据。"
			ErrMsg = ErrMsg & "<br><br>计费方式："
			if sChargeType=1 then
				ErrMsg = ErrMsg & "扣点数<br>初始点数：" & sGroupPoint & "点"
			else
				ErrMsg = ErrMsg & "有效期<br>开始日期：" & Formatdatetime(now(),2) & "<br>有 效 期：" & sValidDays & "天"
			end if
			Dim UserType:UserType=KS.U_G(GroupID,"usertype")
				Conn.Execute("Update KS_User set  UserType=" & UserType & ",GroupID=" & GroupID & ",ChargeType =" & sChargeType & ",Point =" & sGroupPoint & ",BeginDate =#" & formatdatetime(now(),2) & "#,EDays=" & sValidDays & " where UserID in (" & UserID & ")")
			Response.Write KS.ShowError(ErrMsg)
		end sub

		
End Class
%> 

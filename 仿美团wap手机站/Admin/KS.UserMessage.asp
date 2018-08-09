<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

'response.buffer=false
Dim KSCls
Set KSCls = New Admin_UserMessage
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_UserMessage
        Private KS
		Private Action,RSObj,MaxPerPage,TotalPut,CurrentPage
		Private Title, Message, Numc

		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		   MaxPerPage = 20
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		    If Not KS.ReturnPowerResult(0, "KMUA10003") Then
			  Response.Write "<script src='../ks_inc/jquery.js'></script>"
			  Response.Write ("<script>$(parent.document).find('#BottomFrame')[0].src='javascript:history.back();';</script>")
			  Call KS.ReturnErr(1, "")
			End If
	        Response.Write "<html>"
			Response.Write"<head>"
			Response.Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			Response.Write"<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			Response.Write "<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"
			Response.Write"</head>"
			Response.Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			Response.Write"<table width=""100%""  height=""25"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
			Response.Write " <tr>"
			Response.Write"	<td height=""25"" class=""sort""> "
			Response.Write " &nbsp;&nbsp;<strong>用户短信管理：</strong><a href='?action=new'>发送短信</a>"
			Response.Write	" </td>"
			Response.Write " </tr>"
			Response.Write"</TABLE>"
		Action=Trim(Request("Action"))
		Select Case Action
		Case "new","edit"
		    call SendMsg()
		Case "add"
			call savemsg()
		Case "saveedit"
		    Call editsavemsg()
		Case "delall"
			call delall()
		Case "delchk"
			call delchk()
		Case "del"
		    call delbyid()
		Case else
			call main()
		end Select
		Response.Write ""%>
		</body>
		</html>
		<%
		End Sub
		
		Sub Main()
		If Request("page") <> "" Then
			  CurrentPage = CInt(Request("page"))
		Else
			  CurrentPage = 1
		End If
         %>
		<table width="100%" style="border-top:1px #CCCCCC solid" border="0" align="center" cellspacing="0" cellpadding="0">
		  	  <form name="myform" method="Post" action="KS.UserMessage.asp">
				  <tr class='sort'>
					<td height="22" width="30" align="center">选中</td>
					<td align="center">标题</td>
					<td width="80" align="center">发送者</td>
					<td align="center" width="80">接收者</td>
					<td width="100" align="center">发送时间</td>
					<td width="40" align="center">状态</td>
					<td width="120" align="center">操作</td>
				  </tr>
			<%
		           Set RSObj = Server.CreateObject("ADODB.RecordSet")
				   Dim Param:Param=" where 1=1"
				   If KS.S("KeyWord")<>"" Then
				     select case KS.ChkClng(KS.S("condition"))
					   case 1
					    Param=Param & " and title like '%" & KS.S("KeyWord") & "%'"
					   case 2
					    Param=Param & " and Sender like '%" & KS.S("KeyWord") & "%'"
					   case 3
					    Param=Param & " and Incept like '%" & KS.S("KeyWord") & "%'"
					 end select 
				   End If
				   RSObj.Open "SELECT * FROM KS_Message " & Param & " order by id Desc", Conn, 1, 1
				 If RSObj.EOF Then
				    Response.Write "<tr><td colspan=8 height='30' align='center'>找不到任何短消息！</td></tr>"
				 Else
					totalPut = RSObj.RecordCount
		
							If CurrentPage < 1 Then
								CurrentPage = 1
							End If
		
							If (CurrentPage - 1) * MaxPerPage > totalPut Then
								If (totalPut Mod MaxPerPage) = 0 Then
									CurrentPage = totalPut \ MaxPerPage
								Else
									CurrentPage = totalPut \ MaxPerPage + 1
								End If
							End If
		
							If CurrentPage = 1 Then
								Call showContent
							Else
								If (CurrentPage - 1) * MaxPerPage < totalPut Then
									RSObj.Move (CurrentPage - 1) * MaxPerPage
									Call showContent
								Else
									CurrentPage = 1
									Call showContent
								End If
							End If
			End If
				 %>	
		<tr class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
			<td colspan=8 height="30">
			<input type="hidden" value="del" name="action">
			<input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">选中本页显示的所有记录&nbsp;<input type="submit" value="删除选中的记录" onclick="return(confirm('确定删除选中的记录吗？'))" class="button">
			&nbsp
			<input type="button" value="发送短信" onclick="location.href='?action=new';" class="button">
					 </td>
		  </tr> 
		  <%
		  Response.Write "<tr><td colspan='7' align='right'>"
		  			 Call KS.ShowPageParamter(totalPut, MaxPerPage, "KS.UserMessage.asp", True, "条", CurrentPage, "KeyWord=" & KS.S("KeyWord") &"&condition=" & ks.s("condition"))

			Response.Write "</td></tr>"

		  %> 
		</form>
		</table>
		<div>
		<form action="KS.UserMessage.asp" name="myform" method="post">
		   <div style="border:1px dashed #cccccc;margin:3px;padding:4px">
			  &nbsp;<strong>快速搜索=></strong>
			 &nbsp;关键字:<input type="text" class='textbox' name="keyword">&nbsp;条件:
			 <select name="condition">
			  <option value=1>短信标题</option>
			  <option value=2>发送用户</option>
			  <option value=3>接收用户</option>
			 </select>
			  &nbsp;<input type="submit" value="开始搜索" class="button" name="s1">
			  </div>
		</form>
		</div>
		
		<table width="100%" border="0" align=center cellpadding="3" cellspacing="1" class="ctable">
		  <tr align="center" class="sort"> 
			<td height="25" colspan="2">短消息管理(批量删除)</td>
		  </tr>
		  <form action="KS.UserMessage.asp?action=del" method=post>
		  </form>
		  <form action="KS.UserMessage.asp?action=delall" method=post>
			<tr> 
			  <td colspan="2" bgcolor="#FFFFFF" class=tdbg> 批量删除用户指定日期内短消息（默认为删除已读信息）：<br>
				<select name="delDate" size=1>
				  <option value=7>一个星期前</option>
				  <option value=30>一个月前</option>
				  <option value=60>两个月前</option>
				  <option value=180>半年前</option>
				  <option value="all">所有信息</option>
				</select>
				&nbsp; 
				<input type="checkbox" name="isread" value="yes">
				包括未读信息 
				<input type="submit" name="Submit" class="button" value="提 交">
			  </td>
			</tr>
		  </form>
		  <form action="KS.UserMessage.asp?action=delchk" method=post>
			<tr> 
			  <td colspan="2" bgcolor="#FFFFFF" class=tdbg> 批量删除含有某关键字短信（注意：本操作将删除所有已读和未读信息）：<br>
				关键字： 
				<input class="textbox" type="text" name="keyword" size=30>
				&nbsp;在 
				<select name="selaction" size=1>
				  <option value=1>标题中</option>
				  <option value=2>内容中</option>
				</select>
				&nbsp; 
				<input type="submit" name="Submit" value="提 交" class='button'>
			  </td>
			</tr>
		  </form>
		</table>
		<%
		End Sub
		
		Sub ShowContent()
		 Dim i:i=1
		 Do While Not RSObj.Eof
		 %>
		  <tr height="23" class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
		    <td><input name="ID" type="checkbox" onClick="unselectall()" id="ID" value="<%=RSObj("ID")%>"></td>
			<td><img src="images/Announce.gif" align="absmiddle"><a href="?action=edit&id=<%=rsobj("id")%>"><%=KS.Gottopic(rsobj("title"),35)%></a></td>
			<td><%=rsobj("sender")%></td>
			<td><%=rsobj("Incept")%></td>
			<td><%=rsobj("sendtime")%></td>
			<td align="center">
			<%if rsobj("flag")=0 then
			   response.write "<font color=red>未读</font>"
			  else
			   response.write "<font color=blue>已读</font>"
			  end if
			 %>
			</td>
			<td align="center"><a href="?action=edit&id=<%=rsobj("id")%>">修改</a> | <a onclick="return(confirm('删除后不可恢复，确定删除吗?'))" href="?action=del&id=<%=rsobj("id")%>">删除</a></td>
		  </tr>
		   <tr><td colspan='8' background='images/line.gif'></td></tr>
		 <%if i>=maxperpage then exit do
		   i=I+1
		  RSObj.MoveNext
		 Loop
		End Sub
		
		Sub SendMsg()
		  dim flag,display,Incept,title,content,sendtime
		  If KS.S("Action")="edit" then
		    flag="saveedit"
			display=" style='display:none'"
			dim rs:set rs=server.createobject("adodb.recordset")
			rs.open "select * from ks_message where id="& ks.chkclng(ks.s("id")),conn,1,1
			if not rs.eof then
			 Incept=rs("Incept")
			 title=rs("title")
			 content=rs("content")
			 sendtime=rs("sendtime")
			end if
			rs.close:set rs=nothing
		  else
		    flag="add"
			
			 dim userid:userid=KS.FilterIds(replace(request("userid")," ",""))
			 dim usernamelist
			 if userid<>"" then
				 set rs=KS.InitialObject("adodb.recordset")
				 rs.open "select userid,username from ks_user where userid in("& userid & ")",conn,1,1
				 do while not rs.eof
				  if usernamelist="" then
				   usernamelist=rs(1)
				  else
				   usernamelist=usernamelist &"," & rs(1)
				  end if
				  rs.movenext
				 loop
				 rs.close:set rs=nothing
			 end if

			
			
			
		  end if
		%>
		<table width="100%" border="0" style="margin-top:3px" align=center cellpadding="3" cellspacing="1" class="ctable">
		  
		  <form action="KS.UserMessage.asp?action=<%=flag%>" method=post name="myform" id="myform">
		   <input type="hidden" value="<%=KS.S("id")%>" name="id">
			<tr class="sort">
			  <td height="25" colspan="2" align="center">发送短消息</td>
		    </tr>
			<tr class="tdbg"<%=display%>>
				<td height="25" align="right" class="clefttitle">用户类别：</td>
				<td>
				<Input type="radio" name="UserType" value="1" checked onclick="UType(this.value)">用户名单
				<Input type="radio" name="UserType" value="2" onclick="UType(this.value)">用户组
				<Input type="radio" name="UserType" value="0" onclick="UType(this.value)">所有用户				</td>
			</tr>
			<%if ks.s("action")="edit" then%>
			<tr class="tdbg" id="ToUserName">
				<td height="25" align="right" class="clefttitle">接收用户：</td>
				<td> 
				<%=Incept%>
				</td>
			</tr>
            <tr class="tdbg">
			   <td height="25" align="right" class="clefttitle">发送时间：</td>
			   <td><input type="text" name="sendtime" value="<%=sendtime%>"> <font color=red>格式：0000-00-00 00:00</font></td>
			</tr>
			<%else%>
			<tr class="tdbg" id="ToUserName">
				<td height="25" align="right" class="clefttitle">用 户 名：</td>
				<td> <INPUT class="textbox" TYPE="text" value="<%=usernamelist%>" NAME="UserName" size="80"><br>
				请输入用户名：(多个用户名请以英文逗号“,”分隔,注意区分大小写)</td>
			</tr>
			<%end if%>
			<tr class="tdbg" id="ToGroupID" style="display:none;">
				<td height="25" align="right" class="clefttitle">用 户 组：</td>
				<td>
					<table width="100%" border="0" cellspacing="1" cellpadding="3" align=center>
					<tr><td>
					<%=KS.GetUserGroup_CheckBox("GroupID",0,5)
					%>
					</td></tr>
					<tr><td height=20><input type="button" value="打开高级设置" NAME="OPENSET" onclick="openset(this,'UpSetting')"></td></tr>
					<tr><td height=20 ID="UpSetting" style="display:NONE">
						<table width="100%" border="0" cellspacing="1" cellpadding="3" align=center>
						<tr><td height=20 colspan="4">符合条件设置(以下条件将对选择的用户组生效)</td></tr>
						<tr>
							<td width="15%">最后登陆时间：</td>
							<td width="35%">
							<input class="textbox" type="text" name="LoginTime" onkeyup="CheckNumber(this,'天数')" size=6>天 &nbsp;<INPUT TYPE="radio" NAME="LoginTimeType" checked value="0">多于 <INPUT TYPE="radio" NAME="LoginTimeType" value="1">少于							</td>
							<td width="15%">注册时间：</td>
							<td width="35%">
							<input class="textbox" type="text" name="RegTime" onkeyup="CheckNumber(this,'天数')" size=6>天 &nbsp;<INPUT TYPE="radio" NAME="RegTimeType" checked value="0">多于 <INPUT TYPE="radio" NAME="RegTimeType" value="1">少于							</td>
						</tr>
						<tr>
							<td>登陆次数：</td>
							<td><input class="textbox" type="text" name="Logins" size=6 onkeyup="CheckNumber(this,'次数')">次 &nbsp;<INPUT TYPE="radio" NAME="LoginsType" checked value="0">多于 <INPUT TYPE="radio" NAME="LoginsType" value="1">少于							</td>
							<td>发表文章：</td>
							<td><input class="textbox" type="text" name="UserArticle" size=6 onkeyup="CheckNumber(this,'篇数')">篇 &nbsp;<INPUT TYPE="radio" NAME="UserArticleType" checked value="0">多于 <INPUT TYPE="radio" NAME="UserArticleType" value="1">少于</td>
						</tr></table>
					</td></tr></table>				</td>
			</tr>
			<tr class=tdbg> 
			  <td width="20%" height="25" align="right" class="clefttitle">消息标题：</td>
			  <td width="80%"> 
				<input class="textbox" type="text"  value="<%=title%>" name="title" size="80">			  </td>
			</tr>
			<tr class=tdbg> 
			  <td width="20%" height="25" align="right" class="clefttitle">消息内容：</td>
			  <td width="80%"> 
				 <textarea id="message" name="message"  style="display:none"><%=server.htmlencode(content)%></textarea>
				 <script type="text/javascript" src="../editor/ckeditor.js" mce_src="../editor/ckeditor.js"></script>
			    <script type="text/javascript">
                CKEDITOR.replace('message', {width:"690",height:"200px",toolbar:"NewsTool",filebrowserBrowseUrl :"Include/SelectPic.asp?from=ckeditor&Currpath=<%=KS.GetUpFilesDir()%>",filebrowserWindowWidth:650,filebrowserWindowHeight:290});
			   </script>
	  
				</td>
			</tr>
			<tr class=tdbg> 
			  <td height="25" colspan="2" style="text-align:center"> 
				  <input type="submit" name="Submit" value="发送消息" class='button' onclick="return(checkform())">
				  <input type="reset" name="Submit2" value="重新填写" class='button'>			  </td>
		    </tr>
		  </form>
		</table>
		<script>
		 function checkform()
		 {
		   if (document.myform.title.value==''){
			 alert('站内短信标题不能为空！');
			 document.myform.title.focus();
			 return false;
		  }
		  if ((FCKeditorAPI.GetInstance('message').GetXHTML(true)==""))
			{
			  alert("站内短信内容不能为空！");
			  FCKeditorAPI.GetInstance('message').Focus();
			  return false;
		   } 
		  
       return true
		 }
		</script>
		<br>
		
		<SCRIPT LANGUAGE="JavaScript">
		<!--
		function openset(v,s){
			if (v.value=='打开高级设置'){
				document.getElementById(s).style.display = "";
				v.value="关闭高级设置";
			}
			else{
				v.value="打开高级设置";
				document.getElementById(s).style.display = "none";
			}
		}
		function UType(n){
			if (n==1){
				document.getElementById("ToUserName").style.display = "";
				document.getElementById("ToGroupID").style.display = "none";
			}
			else if(n==2){
				document.getElementById("ToUserName").style.display = "none";
				document.getElementById("ToGroupID").style.display = "";
			}
			else{
				document.getElementById("ToUserName").style.display = "none";
				document.getElementById("ToGroupID").style.display = "none";
			}
		}
		//-->
		</SCRIPT>
		<%
		end sub
		
		sub editsavemsg()
		   dim id:id=ks.chkclng(ks.s("id"))
		   dim title:title=ks.g("title")
		   dim content:content=Request.Form("message")
		   dim sendtime:sendtime=ks.s("sendtime")
		   if not isdate(sendtime) then
		    Response.Write "<script>alert('时间格式不正确!');history.back();</script>"
			Exit Sub
			end if
			dim rs:set rs=server.createobject("adodb.recordset")
			rs.open "select  top 1 * from ks_message where id=" &id,conn,1,3
			if not rs.eof then
			  rs("title")=title
			  rs("content")=content
			  rs("sendtime")=sendtime
			  rs.update
			end if
			rs.close
			set rs=nothing
			response.write "<script>alert('恭喜，修改成功!');location.href='ks.usermessage.asp';</script>"
		   response.end
		end sub
		
		Sub delbyid()
		  If Ks.G("id")="" Then
				Response.Write("<script>alert('参数传递出错!');history.back();</script>")
				Exit Sub
			end if
		    Conn.Execute("delete from ks_message where id in(" & KS.FilterIDs(KS.G("id")) &")")
			Response.Write Response.Write("<script>alert('恭喜，删除操作成功！');location.href='" & Request.ServerVariables("HTTP_REFERER") & "';</script>")
		End Sub
		
		Sub del()
			if KS.G("username")="" then
				Response.Write("<script>alert('请输入要批量删除的用户名!');history.back();</script>")
				Exit Sub
			end if
			sql="delete from KS_Message where sender='"&KS.G("username")&"'"

			Conn.Execute(sql)
			
			Response.Write Response.Write("<script>alert('操作成功！请继续别的操作!');</script>")
		End Sub
		
		sub delall()
			dim selflag,sql
			if request("isread")="yes" then
			selflag=""
			else
			selflag=" and flag=1"
			end if
				select case request("delDate")
				case "all"
				sql="delete from KS_Message where id>0 "&selflag
				case 7
				sql="delete from KS_Message where datediff(" & DataPart_D & ",sendtime," & SqlNowString & ")>7 "&selflag
				case 30
				sql="delete from KS_Message where datediff(" & DataPart_D & ",sendtime," & SqlNowString & ")>30 "&selflag
				case 60
				sql="delete from KS_Message where datediff(" & DataPart_D & ",sendtime," & SqlNowString & ")>60 "&selflag
				case 180
				sql="delete from KS_Message where datediff(" & DataPart_D & ",sendtime," & SqlNowString & ")>180 "&selflag
				end select
				Conn.Execute(sql)

			Call KS.Alert("操作成功！请继续别的操作。","KS.UserMessage.asp")
		end Sub
		
		Sub delchk()
			if request.form("keyword")="" then
				KS.ShowError("请输入关键字！")
				Exit sub
			end if
			if request.form("selaction")=1 then
					conn.Execute("delete from KS_Message where title like '%"&replace(request.form("keyword"),"'","")&"%'")
			elseif request.form("selaction")=2 then
				
					conn.Execute("delete from KS_Message where content like '%"&replace(request.form("keyword"),"'","")&"%'")
			else
				KS.ShowError("未指定相关参数！")
			end if
			Call KS.Alert("操作成功！请继续别的操作。","KS.UserMessage.asp")
		End Sub
		
		Sub SaveMsg()
			Server.ScriptTimeout=99999
			Dim UserType
			UserType = Trim(Request.Form("UserType"))
			Title	 = Trim(Request.Form("title"))
			Message  = Request.Form("message")
			If Title="" or Message="" Then
				KS.Showerror("请填写消息的标题和内容!")
				Exit Sub
			End If
			If Len(Message) > KS.Setting(48) Then
				KS.Showerror("消息内容不能多于" & KS.Setting(48) & "字节")
				Exit Sub
			End If 
 
			Select Case UserType
			Case "0" : SaveMsg_0()	'按所有用户
			Case "1" : SaveMsg_1()	'按指定用户
			Case "2" : SaveMsg_2()	'按指定用户组
			Case Else
				KS.Showerror("请输入收信的用户!") : Exit Sub
			End Select
			Call KS.Alert("操作成功！本次发送"&Numc+1&"个用户。请继续别的操作。","KS.UserMessage.asp")
		End Sub
		'按所有用户发送
		Sub SaveMsg_0()
			Dim Rs,Sql,i
			Sql = "Select UserName From KS_User Order By UserID Desc"
			Set Rs = Conn.Execute(Sql)
			If Not Rs.eof Then
				SQL = Rs.GetRows(-1)
				Numc= Ubound(SQL,2)
				For i=0 To Numc
				   if cbool(KS.SendInfo(SQL(0,i),KS.C("AdminName"),Replace(Title,"'","''"),Replace(Message,"'","''")))=false then
				    KS.Die "<script>alert('用户" & SQL(0,I) & "不存在或是该用户邮箱已满，信件发送失败！');history.back();</script>"
				   end if
				Next
			End If
			Rs.Close : Set Rs = Nothing
		End Sub
		'按指定用户
		Sub SaveMsg_1()
			Dim ToUserName,Rs,Sql,i
			ToUserName = Trim(Request.Form("UserName"))
			If ToUserName = "" Then
				KS.Showerror("请填写目标用户名，注意区分大小写。")
				Exit Sub
			End If
			ToUserName = Replace(ToUserName,"'","")
			ToUserName = Split(ToUserName,",")
			Numc= Ubound(ToUserName)
			For i=0 To Numc
				if cbool(KS.SendInfo(ToUserName(i),KS.C("AdminName"),Title,Replace(Message,"'","''")))=false then
				  KS.Die "<script>alert('用户" & ToUserName(i) & "不存在或是该用户邮箱已满，信件发送失败！');history.back();</script>"
				end if
			Next
		End Sub
		'按指定用户组及条件发送
		Sub SaveMsg_2()
			Dim GroupID,ErrMsg,i
			Dim SearchStr,TempValue,DayStr
			GroupID = Replace(Request.Form("GroupID"),chr(32),"")
			If GroupID="" Then
			    ErrMsg = "请正确选取相应的用户组。"
			ElseIf GroupID<>"" and Not Isnumeric(Replace(GroupID,",","")) Then
				ErrMsg = "请正确选取相应的用户组。"
			Else
				GroupID = KS.R(GroupID)
			End If
			DayStr = "'d'"
			If Instr(GroupID,",")>0 Then
				SearchStr = "GroupID in ("&GroupID&")"
			Else
				SearchStr = "GroupID = "&KS.R(GroupID)
			End If
			'登陆次数
			TempValue = Request.Form("Logins")
			If TempValue<>"" and IsNumeric(TempValue) Then
				SearchStr = GetSearchString(TempValue,SearchStr,Request.Form("LoginsType"),"LoginTimes")
			End If
			'发表文章
			TempValue = Request.Form("UserArticle")
			If TempValue<>"" and IsNumeric(TempValue) Then
				SearchStr = GetSearchString(TempValue,SearchStr,Request.Form("UserArticleType"),"(select count(id) from ks_iteminfo where inputer=ks_user.username)")
			End If
			'最后登陆时间
			TempValue = Request.Form("LoginTime")
			If TempValue<>"" and IsNumeric(TempValue) Then
				SearchStr = GetSearchString(TempValue,SearchStr,Request.Form("LoginTimeType"),"Datediff("&DayStr&",LastLoginTime,"&SqlNowString&")")
			End If
			'注册时间
			TempValue = Request.Form("RegTime")
			If TempValue<>"" and IsNumeric(TempValue) Then
				SearchStr = GetSearchString(TempValue,SearchStr,Request.Form("RegTimeType"),"Datediff("&DayStr&",JoinDate,"&SqlNowString&")")
			End If
			If SearchStr="" Then
				ErrMsg = "请填写发送的条件选项。"
			End If
			If ErrMsg<>"" Then KS.Showerror(ErrMsg) : Exit Sub
			Dim Rs,Sql
			Sql = "Select UserName From KS_User Where "& SearchStr & " Order By UserID Desc"
			
			Set Rs = Conn.Execute(Sql)
			If Not Rs.eof Then
				SQL = Rs.GetRows(-1)
				Numc= Ubound(SQL,2)
				For i=0 To Numc
				    IF Cbool(KS.SendInfo(SQL(0,i),KS.C("AdminName"),Replace(Title,"'","''"),Replace(Message,"'","''")))=false then
					 KS.Die "<script>alert('用户" & SQL(0,I) & "不存在或是该用户邮箱已满，信件发送失败！');history.back();</script>"
					end if
				Next
			End If
			Rs.Close : Set Rs = Nothing
		End Sub
		
		Function GetSearchString(Get_Value,Get_SearchStr,UpType,UpColumn)
			Get_Value = Clng(Get_Value)
			If Get_SearchStr<>"" Then Get_SearchStr = Get_SearchStr & " and " 
			If UpType="1" Then
				Get_SearchStr = Get_SearchStr & UpColumn &" <= "&Get_Value
			Else
				Get_SearchStr = Get_SearchStr & UpColumn &" >= "&Get_Value
			End If
			GetSearchString = Get_SearchStr
		End Function
End Class
%> 

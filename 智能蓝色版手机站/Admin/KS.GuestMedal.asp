<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%dim channelid
Dim KSCls
Set KSCls = New Admin_Ask_Class
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Ask_Class
        Private KS,DataArry,Numc,MedalID,MedalIds
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
         If Not KS.ReturnPowerResult(0, "KSMB10003") Then          '检查是权限
					 Call KS.ReturnErr(1, "")
					 KS.Die ""
		 End If
		%>
		<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
		<head>
		<link href="Include/Admin_Style.CSS" rel="stylesheet" type="text/css">
		<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
		<script src="../KS_Inc/jquery.js" language="JavaScript"></script>
		<script src="../KS_Inc/common.js" language="JavaScript"></script>
		</head>
		<body>
		<div class='topdashed sort' style="padding-left:20px;text-align:left"><a href="KS.GuestMedal.asp">论坛勋章设置</a> | <a href="KS.Guestmedal.asp?action=pub">颁发勋章</a> | <a href="KS.GuestMedal.asp?action=clear" onClick="return(confirm('此操作直接清除所有用户的勋章，且不可恢复！确认操作吗?'))">一键重置用户勋章</a></div>

		<%
		Dim Action,DataArry
		Action = LCase(Request("action"))
		Select Case Trim(Action)
		Case "pub" Call Pub()
		Case "clear" call clearmedal()
		Case "dopub" Call DoPub()
		Case "modify" Call Modify()
		Case "modifysave" Call ModifySave()
		Case "save"
			Call saveScore()
		Case Else
			Call showmain()
		End Select
		End Sub
		
		Sub clearmedal()
		Conn.Execute("Update KS_User Set Medal=''")
		KS.AlertHintScript "恭喜，所有用户的勋章已被清除!"
		End Sub
		Sub Pub()
		%>
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
		<table width="100%" border="0" style="margin-top:3px" align=center cellpadding="3" cellspacing="1" class="ctable">
		 <form action="KS.GuestMedal.asp" method=post name="myform" id="myform">
		   <input type="hidden" value="DoPub" name="action">
			<tr class="sort">
			  <td height="25" colspan="2" align="center">颁发勋章</td>
		    </tr>
			<tr class="tdbg">
				<td height="25" width="120" align="right" class="clefttitle">用户类别：</td>
				<td>
				<Input type="radio" name="UserType" value="1" checked onClick="UType(this.value)">用户名单
				<Input type="radio" name="UserType" value="2" onClick="UType(this.value)">用户组
				<Input type="radio" name="UserType" value="0" onClick="UType(this.value)">所有用户				</td>
			</tr>
			
			<tr class="tdbg" id="ToUserName">
				<td height="25" align="right" class="clefttitle">用 户 名：</td>
				<td> <INPUT class="textbox" TYPE="text" value="" NAME="UserName" size="80"><br>
				请输入用户名：(多个用户名请以英文逗号“,”分隔,注意区分大小写)</td>
			</tr>
			
			<tr class="tdbg" id="ToGroupID" style="display:none;">
				<td height="25" align="right" class="clefttitle">用 户 组：</td>
				<td>
					<table width="100%" border="0" cellspacing="1" cellpadding="3" align=center>
					<tr><td>
					<%=KS.GetUserGroup_CheckBox("GroupID","",5)%>
					
					</td></tr>
					<tr><td height=20><input type="button" class="button" value="打开高级设置" NAME="OPENSET" onClick="openset(this,'UpSetting')"></td></tr>
					<tr><td height=20 ID="UpSetting" style="display:NONE">
						<table width="100%" border="0" cellspacing="1" cellpadding="3" align=center>
						<tr><td height=20 colspan="4">符合条件设置(以下条件将对选择的用户组生效)</td></tr>
						<tr>
							<td width="15%">最后登陆时间：</td>
							<td width="35%">
							<input class="textbox" type="text" name="LoginTime" onKeyUp="CheckNumber(this,'天数')" size=6>天 &nbsp;<INPUT TYPE="radio" NAME="LoginTimeType" checked value="0">多于 <INPUT TYPE="radio" NAME="LoginTimeType" value="1">少于							</td>
							<td width="15%">注册时间：</td>
							<td width="35%">
							<input class="textbox" type="text" name="RegTime" onKeyUp="CheckNumber(this,'天数')" size=6>天 &nbsp;<INPUT TYPE="radio" NAME="RegTimeType" checked value="0">多于 <INPUT TYPE="radio" NAME="RegTimeType" value="1">少于							</td>
						</tr>
						<tr>
							<td>登陆次数：</td>
							<td><input class="textbox" type="text" name="Logins" size=6 onKeyUp="CheckNumber(this,'次数')">次 &nbsp;<INPUT TYPE="radio" NAME="LoginsType" checked value="0">多于 <INPUT TYPE="radio" NAME="LoginsType" value="1">少于							</td>
							<td>发表文章：</td>
							<td><input class="textbox" type="text" name="UserArticle" size=6 onKeyUp="CheckNumber(this,'篇数')">篇 &nbsp;<INPUT TYPE="radio" NAME="UserArticleType" checked value="0">多于 <INPUT TYPE="radio" NAME="UserArticleType" value="1">少于</td>
						</tr></table>
					</td></tr></table>				</td>
			</tr>
			<tr>
			  <td class="clefttitle">选择勋章：</td>
			  <td>
			   <table bordr='0'>
			  <%dim rs:set rs=conn.execute("select * from ks_guestmedal where status=1 order by medalid")
			  do while not rs.eof
			    response.write "<tr><td><img  width='25' src='../" & KS.Setting(66) & "/images/medal/" & rs("ico") & "'></td><td><label><input type='checkbox' name='medalid' value='" & rs("medalid") & "'>" & rs("medalname") & "</label></td></tr>"
			  rs.movenext
			  loop
			  rs.close
			  set rs=nothing
			  %>
			   </table>
			  </td>
			</tr>
			<tr class="tdbg">
			 <td></td><td style="height:40px"><input class="button" type="submit" value="确定颁发">
			 </td>
		    </tr>
		</form>
		</table>
		<div class="attention" style="color:red">
		<strong>说明:</strong>这里颁发勋章会重新覆盖原来用户的勋章，请慎重操作。
		</div>
		<%
		End Sub
		
        Sub DoPub()
			Server.ScriptTimeout=99999
			Dim UserType
			UserType = Trim(Request.Form("UserType"))
			medalid	 = KS.FilterIds(Request.Form("medalid"))
            If medalid="" Then KS.AlertHintScript "对不起，您没有选择勋章!"
			Dim RS:Set RS=Conn.Execute("Select * From KS_GuestMedal Where MedalID in(" & MedalID &") order by medalid")
			If Not RS.Eof Then
			   Do While NOt RS.EOf
			     if MedalIds="" Then
			      MedalIds=rs("medalid") & "|" & rs("medalname") & "|" & rs("ico")
				 else
				  MedalIds=MedalIds & "@@@" & rs("medalid") & "|" & rs("medalname") & "|" & rs("ico") 
				 end if
			   rs.MoveNext
			   Loop
			End If
			RS.Close :Set RS=Nothing
			Select Case UserType
			Case "0" : SaveMsg_0()	'按所有用户
			Case "1" : SaveMsg_1()	'按指定用户
			Case "2" : SaveMsg_2()	'按指定用户组
			Case Else
				KS.Showerror("请输入接收勋章的用户!") : Exit Sub
			End Select
			KS.AlertHintScript "操作成功！本次颁发"&Numc+1&"个用户。"
		End Sub
		'按所有用户发送
		Sub SaveMsg_0()
			Dim Rs,Sql,i
			Sql = "Select UserName From KS_User Order By UserID Desc"
			Set Rs = Conn.Execute(Sql)
			If Not Rs.eof Then
				Numc= conn.execute("select count(1) from ks_user")(0)
				Conn.Execute("update ks_user set medal='" & MedalIds &"'")
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
			Numc= Ubound(Split(ToUserName,","))
			Conn.Execute("update ks_user set medal='" & MedalIds &"' where username in('" & replace(ToUserName,",","','") &"')")
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
					Conn.Execute("update ks_user set medal='" & MedalIds &"' where username='" & SQL(0,i) & "'")
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
			
		'勋章修改
		Sub Modify()
		 Dim RS,id
		 id=KS.ChkClng(KS.S("ID"))
		 If ID=0 Then KS.AlertHintScript "出错啦!"
		 Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select * From KS_GuestMedal Where MedalID=" & ID,conn,1,1
		 If RS.Eof And RS.Bof Then
		   RS.Close : Set RS=Nothing
		   KS.AlertHintscript "勋章不存在!"
		 End If
		%>
		<script type="text/javascript">
		 function CheckForm(){
		  if($("#MedalName").val()==''){
		   alert('请输入勋章名称!');
		   $("#MedalName").focus();
		   return false;
		  }
		  if($("#ICO").val()==''){
		   alert('请输入勋章图片!');
		   $("#ICO").focus();
		   return false;
		  }
		  $("#myform").submit();
		 }
		</script>
		<form name="myform" id="myform" action="KS.GuestMedal.asp" method="post">
		 <input type="hidden" name="action" value="modifySave"/>
		 <input type="hidden" name="id" value="<%=rs("medalid")%>"/>
         <table width="100%" border="0" cellpadding="1" cellspacing="1" class='ctable'>
		   <tr class="tdbg">
		     <td class="clefttitle" width="150"><strong>勋章名称：</strong></td>
			 <td><input class="textbox" type="text" name="MedalName" id='MedalName' value="<%=rs("MedalName")%>" /></td>
		   </tr>
		   <tr class="tdbg">
		     <td class="clefttitle" width="150"><strong>勋章启用：</strong></td>
			 <td><input type="checkbox" name="Status" value="1"<%if rs("Status")="1" then response.write " checked"%> /></td>
		   </tr>
		   <tr class="tdbg">
		     <td class="clefttitle" width="150"><strong>勋章图片：</strong></td>
			 <td><input type="text" class="textbox" name="ICO" id='ICO' value="<%=rs("ico")%>" /> <img align="absmiddle" width="30" src="../<%=KS.Setting(66)%>/images/medal/<%=rs("ico")%>" /></td>
		   </tr>
		   <tr class="tdbg">
		     <td class="clefttitle" width="150"><strong>勋章介绍：</strong></td>
			 <td><textarea  name="descript" style="width:230px;height:40px;overflow:auto"><%=rs("descript")%></textarea></td>
		   </tr> 
		   <tr class="tdbg">
		     <td class="clefttitle" width="150"><strong>领取方式：</strong></td>
			 <td>
			 <label><input type="radio" onClick="$('#score').hide();$('#apply').hide();" name="lqfs" value="0"<%if rs("lqfs")="0" then response.write " checked"%> />管理员手工发放</label><br/>
			 <label><input type="radio" onClick="$('#score').hide();$('#apply').show();" name="lqfs" value="1"<%if rs("lqfs")="1" then response.write " checked"%> />用户申请发放</label><br/>
			 <label><input type="radio" onClick="$('#score').show();$('#apply').hide();" name="lqfs" value="2"<%if rs("lqfs")="2" then response.write " checked"%> />积分购买</label>
			 </td>
		   </tr>
		   
		   
		   
		   <tr class="tdbg" id='apply' <%if rs("lqfs")<>"1" then response.write " style='display:none'"%>>
		     <td class="clefttitle" width="150"><strong>领取权限：</strong> </td>
			 <td>
			  <%
			  dim Expression:Expression=rs("Expression") & "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
			  dim expArr:expArr=split(Expression,",")
			  %>
			   <div style="margin:14px 0px 5px 0px;font-weight:bold">当选择用户申请发放时，请设置以下领取权限(<span style='font-weight:normal;color:#999'>不限制请输入0</span>)：</div>
			   发帖量>= <input class="textbox" type="text" name="fdl" size="4" style="text-align:center" value="<%=expArr(0)%>"> 帖<br/>  
			   精华帖>= <input class="textbox" type="text" name="fdl" size="4" style="text-align:center" value="<%=expArr(1)%>"> 帖<br/>  
			   主题总数>= <input class="textbox" type="text" name="fdl" size="4" style="text-align:center" value="<%=expArr(2)%>"> 帖<br/>  
			   积分>= <input class="textbox" type="text" name="fdl" size="4" style="text-align:center" value="<%=expArr(3)%>"> 分<br/>  
			   威望>= <input class="textbox" type="text" name="fdl" size="4" style="text-align:center" value="<%=expArr(4)%>"> 分<br/>  
			   资金>= <input class="textbox" type="text" name="fdl" size="4" style="text-align:center" value="<%=expArr(5)%>"> 分<br/>  
			   点券>= <input class="textbox" type="text" name="fdl" size="4" style="text-align:center" value="<%=expArr(6)%>"> 分<br/> 
			   <div style="margin:14px 0px 5px 0px;font-weight:bold">论坛等级限制(<span style='font-weight:normal;color:#999'>不限制请留空</span>)：</div>
			   <%
			    Dim rsg:set rsg=conn.execute("select GradeID,UserTitle From KS_AskGrade Where TypeFlag=1 Order By GradeID")
				Do While Not RSG.Eof
				  response.write "<label><input type='checkbox' name='gradeid' value='" & RSG(0) & "'"
				  If KS.FoundInArr(RS("GradeID"),rsg(0),",") Then
				   response.write " checked"
				  End If
				  response.write ">" & rsg(1) & "</label>"
				rsg.moveNext
				LOOP
				RSG.Close
				Set RSG=Nothing
			   %>
			   <br/>
			 </td>
		   </tr> 
		   <tr class="tdbg" id='score'<%if rs("lqfs")<>"2" then response.write " style='display:none'"%>>
		     <td class="clefttitle" width="150"><strong>消费积分：</strong> </td>
			 <td>
			 用户花费<input class="textbox" type="text" name="fdl" size="4" style="text-align:center" value="<%=expArr(7)%>"> 分积分可获得
			 </td>
		   </tr>
		 </table>
		 </form>
		<%
		 RS.CLose
		 Set RS=Nothing
		End Sub
		
		Sub ModifySave()
		 Dim ID:ID=KS.ChkClng(KS.G("ID"))
		 If ID=0 Then KS.AlertHintScript "参数传递出错啦！"
		 Dim Expression:Expression=Replace(KS.S("FDL")," ","")&",0,0,0,0,0,0,0,0,0,0,0,0"
		 Conn.Execute("Update KS_GuestMedal Set MedalName='" & KS.G("MedalName") &"',[status]=" & KS.ChkClng(KS.G("Status")) & ",[ico]='" & KS.G("ICO") & "',[Descript]='" & KS.G("Descript") & "',lqfs=" & KS.ChkClng(KS.G("LQFS")) & ",[Expression]='" & Expression&"',gradeid='" & KS.FilterIds(KS.G("GradeID")) & "' Where MedalID=" & ID)
		Response.Write ("<script>alert('恭喜，勋章详情修改成功!');location.href='KS.GuestMedal.asp';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&OpStr=论坛系统 >> <font color=red>勋章管理</font>';</script>")
		End Sub
		
		Sub showmain()
			Dim i,iCount,lCount
			iCount=2:lCount=1
		%>
		<table id="tablehovered" border="0" align="center" cellpadding="0" cellspacing="0" width="100%">
		<form name="selform" id="selform" method="post" action="?">
		<input type="hidden" name="action" value="save">
		<tr class='sort'>
			<td width="10%" noWrap="noWrap">勋章ID</td>
			<td>勋章名称</td>
			<td noWrap="noWrap">启用</td>
			<td noWrap="noWrap">勋章图片</td>
			<td noWrap="noWrap">领取方式</td>
			<td width="15%" noWrap="noWrap">管理操作</td>
		</tr>
		<%
			Call showScoreList()
			iCount=1:lCount=2
			If IsArray(DataArry) Then
				For i=0 To Ubound(DataArry,2)
					If Not Response.IsClientConnected Then Response.End
		%>
		<tr align="center">
			<td class="splittd"><input type="hidden" name="MedalID" value="<%=DataArry(0,i)%>"><%=DataArry(0,i)%></td>
			<td class="splittd"><input type="text" class="textbox" size="20" name="MedalName<%=DataArry(0,i)%>" value="<%=Server.HTMLEncode(DataArry(1,i))%>" /></td>
			<td class="splittd"><input type="CheckBox" size="10" name="Status<%=DataArry(0,i)%>" value="1" <%if DataArry(2,i)="1" then response.write " checked"%>/></td>
			<td class="splittd"><input type="text" class="textbox" size="10" name="ico<%=DataArry(0,i)%>" value="<%=DataArry(4,i)%>" />
			<img align="absmiddle" width="30" src="../<%=KS.Setting(66)%>/images/medal/<%=DataArry(4,i)%>" />
			</td>
			<td class="splittd">
			 <%if DataArry(5,i)="0" then
			   response.write "管理员发放"
			   else
			   response.write "<font color=green>用户申请发放</font>"
			   end if
			 %>
			</td>

			
			<td class="splittd">
			 <a href="javascript:window.$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=论坛系统 >> <font color=red>勋章详情设置</font>&ButtonSymbol=GoSave';location.href='?action=modify&id=<%=DataArry(0,i)%>'">详情</a> 
			 <a href="?x=c&id=<%=DataArry(0,i)%>" onClick="return(confirm('确定删除吗?'))">删除</a>
			</td>
		</tr>
		<%
				Next
			End If
			DataArry=Null
		%>
		<tr align="center">
			<td class="tablerow<%=lCount%>" colspan="6">
				<input class="button" type="submit" name="submit_button" value="批量保存设置"/>			</td>
		</tr>
		</form>

		<form action="?x=b" method="post" name="myform" id="form">
		    <tr>
			<td height="25" colspan="7">&nbsp;&nbsp;<strong>&gt;&gt;新增勋章</strong><<</td>
		    </tr>
			<tr><td colspan=10 background='images/line.gif'></td></tr>
			<tr valign="middle" class="list"> 
			  <td height="25"></td>
			  <td height="25" align="center"><input name="MedalName" type="text" class="textbox" id="MedalName" size="25"></td>
			  <td align="center"><input name="Status" type="checkbox" value="1" checked id="Status"></td>
			  <td align="center"><input style="text-align:center" name="ico" type="text" value="medal<%=i+1%>.gif" class="textbox" id="ico" size="10"></td>
			  <td height="25" align="center">---</td>
			  <td height="25" align="center"><input name="Submit3" class="button" type="submit" value="OK,提交"></td>
			</tr>
			<tr><td colspan=10 background='images/line.gif'></td></tr>
		</form>
		</table>
		<div class="attention" style="color:#FF0000">
		<strong>说明:</strong>勋章图片请上传到<%=KS.Setting(66)%>/images/Medal目录下,这里填写图片名称即可。
		</div>
		<%
		 Select case request("x")
		   case "b"
		       If KS.G("MedalName")="" Then Response.Write "<script>alert('请输入勋章名称!');history.back();</script>":response.end
			    Dim MedalID:MedalID=KS.ChkClng(Conn.Execute("Select Max(MedalID) From KS_GuestMedal")(0))+1
				conn.execute("Insert into KS_GuestMedal(MedalID,MedalName,Status,ICO,LQFS) values(" & MedalID & ",'" & KS.G("MedalName") & "'," & KS.ChkClng(KS.G("Status")) & ",'" & KS.G("ico") & "',0)")
				
				KS.AlertHintScript "恭喜,勋章添加成功!"
		   case "c"
				conn.execute("Delete from KS_GuestMedal where MedalID="& KS.ChkClng(KS.G("id")))
				KS.AlertHintScript "恭喜,勋章删除成功!"
		End Select
		  
		End Sub
		
		Sub showScoreList()
			Dim Rs,SQL
			SQL="SELECT MedalID,MedalName,Status,Descript,Ico,LQFS,Expression FROM [KS_GuestMedal] order by Medalid"
			Set Rs=Conn.Execute(SQL)
			If Not (Rs.BOF And Rs.EOF) Then
				DataArry=Rs.GetRows(-1)
			Else
				DataArry=Null
			End If
			Rs.close()
			Set Rs=Nothing
		End Sub
		
		Sub saveScore()
			Dim Rs,SQL,i
			Dim MedalID,MedalName,Score,Ico,clubpostnum,Status
			    MedalID=Split(Replace(Request.Form("MedalID")," ",""),",")
                For I=0 To Ubound(MedalID)
				 MedalName=Replace(Request.Form("MedalName"&MedalID(I)),"'","")
				 Score=KS.ChkClng(Request.Form("Score"&MedalID(I)))
				 Ico=Request.Form("Ico"&MedalID(I))
				 Status=KS.ChkClng(Request.Form("Status"&MedalID(I)))
				 If MedalID(I)>0 Then
					Conn.Execute ("UPDATE KS_GuestMedal SET Ico='" & Ico & "',MedalName='"&MedalName&"',Status=" & Status &" WHERE MedalID="&MedalID(I))
				 End If
			   Next
			Call KS.AlertHintScript("恭喜您！批量更新勋章成功!")
		End Sub
End Class
%>
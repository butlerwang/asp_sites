<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="Include/Session.asp"-->
<!--#include file="../KS_Cls/Ubbfunction.asp"-->
<%

Dim KSCls
Set KSCls = New Admin_Ask_Setting
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Ask_Setting
        Private KS,KSCls
		Private maxperpage,totalrec,Pcount,pagelinks,showmode,pagenow,count,AskInstalDir
		Private m_intOrder,m_strOrder,SQLQuery,SQLField,Topiclist
		Private topicid,classid,topicmode,PostNum,ExpiredTime,CommentNum,HeadTitle,TopicUseTable
		Private classarr,cid,child,Catelist,Action
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		  AskInstalDir="../" & KS.Asetting(1)
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub

		Public Sub Kesion()
			Action = LCase(Request("action"))
			If Action="quickanswer" Then 
			  QuickAnswer: Response.End()
			ElseIf action="modifyanswer" Then
			  ModifyAnswer: Response.End()
			End If
		%>
		<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
        <html xmlns="http://www.w3.org/1999/xhtml">
		<head>
		<link href="Include/Admin_Style.CSS" rel="stylesheet" type="text/css">
		<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
		<script src="../KS_Inc/common.js" language="JavaScript"></script>
		<script src="../KS_Inc/jquery.js" language="JavaScript"></script>
		<script language="javascript" src="../KS_Inc/popcalendar.js"></script>
		<script type="text/javascript">
		 function answer(id)
		 {  onscrolls=false;
			new parent.KesionPopup().PopupCenterIframe('<b>快速回答</b>',"KS.AskList.asp?Action=QuickAnswer&id="+id,750,380,'auto');
		 }
		 function modifyanswer(id)
		 {  onscrolls=false;
			new parent.KesionPopup().PopupCenterIframe('<b>查看/修改回答</b>',"KS.AskList.asp?Action=modifyAnswer&id="+id,750,380,'auto');
		 }
		</script>
		</head>
		<body>
		<ul id='mt'> <div id='mtl'>问答列表管理：</div><li><a href="?">所有问题列表</a></li>&nbsp;<li>|&nbsp;<a href="?action=verifyanswer">审核用户的回答</a></li></ul>
		
		<%
		    pagenow=KS.ChkClng(Request("page"))
			If pagenow=0 Then pagenow=1
			Select Case Trim(Action)
			Case "save"
				Call saveAsked()
			Case "asked"
			     If Not KS.ReturnPowerResult(0, "WDXT10001") Then          '检查是权限
					 Call KS.ReturnErr(1, "")
					 response.End
				 End If
				Call showAsked()
			Case "del"
			      If Not KS.ReturnPowerResult(0, "WDXT10001") Then          '检查是权限
					 Call KS.ReturnErr(1, "")
					 response.End
				 End If
				Call delTopic()
			Case "delask"
			    If Not KS.ReturnPowerResult(0, "WDXT10001") Then          '检查是权限
					 Call KS.ReturnErr(1, "")
					 response.End
				 End If
				Call delAsked()
			Case "recommend"
			    If Not KS.ReturnPowerResult(0, "WDXT10001") Then          '检查是权限
					 Call KS.ReturnErr(1, "")
					 response.End
				 End If
			    Call Recommend()
			Case "unrecommend"
			    If Not KS.ReturnPowerResult(0, "WDXT10001") Then          '检查是权限
					 Call KS.ReturnErr(1, "")
					 response.End
				 End If
			    Call UnRecommend()
			Case "verify"
			    If Not KS.ReturnPowerResult(0, "WDXT10001") Then          '检查是权限
					 Call KS.ReturnErr(1, "")
					 response.End
				 End If
			    Call Verify()
			Case "unverify"
			    If Not KS.ReturnPowerResult(0, "WDXT10001") Then          '检查是权限
					 Call KS.ReturnErr(1, "")
					 response.End
				 End If
			    Call unVerify()
			Case "setsatis"
			    If Not KS.ReturnPowerResult(0, "WDXT10001") Then          '检查是权限
					 Call KS.ReturnErr(1, "")
					 response.End
				 End If
			    Call SetSatis()
			Case "doanswersave"  DoAnswerSave
			Case "verifyanswer" 
			     If Not KS.ReturnPowerResult(0, "WDXT10004") Then          '检查是权限
					 Call KS.ReturnErr(1, "")
					 response.End
				 End If
			     Call verifyanswer()
			Case "verifyda" verifyda
			Case "delanswer" delanswer
			Case "domodifyanswersave" DoModifyAnswerSave
			Case Else
			    If Not KS.ReturnPowerResult(0, "WDXT10001") Then          '检查是权限
					 Call KS.ReturnErr(1, "")
					 response.End
				 End If
				Call showmain()
			End Select
	   End Sub
	   
	   Sub verifyanswer()
	      Dim sqlStr,RS,TotalPut,i
		  MaxPerPage=20
		  sqlStr="select * from KS_AskPosts1 order by postsid desc"
		  Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open SQLStr,conn,1,1
		  %>
		 <table  border="0" align="center" style='border-top:1px solid #cccccc' cellpadding="0" cellspacing="0" width="100%">
		<tr class="sort">
			<td width="5%" noWrap="noWrap">选择</td>
			<td>问题</td>
			<td  noWrap="noWrap">回答内容</td>
			<td width="6%" noWrap="noWrap">回答人</td>
			<td>回答时间</td>
			<td width="4%" noWrap="noWrap">状态</td>
			<td noWrap="noWrap">管理操作</td>
		</tr>
		
		<form name="myform" id="myform" method="post" action="?">
		<input type="hidden" name="action" id="action" value="delanswer">
		<input type="hidden" name="v" id="v" value="0">
		<%
			If RS.Eof And RS.Bof Then
			  Response.Write "<tr><td class='splittd' colspan=6 align='center'>对不起, 找不到相关问题回答!</td></tr>"
			Else
					totalPut = RS.RecordCount
					If pagenow < 1 Then	pagenow = 1
					If pagenow >1 and (pagenow - 1) * MaxPerPage < totalPut Then
						RS.Move (pagenow - 1) * MaxPerPage
					Else
						pagenow = 1
					End If
			    i=0
				Do While Not RS.Eof

		%>
		<tr onMouseOut="this.className='list'" onMouseOver="this.className='listmouseover'" onClick="chk_iddiv('<%=rs("postsid")%>')" id='u<%=rs("postsid")%>'>
			<td class="splittd" align="center"><input type="checkbox" name="id" id='c<%=rs("postsid")%>' value="<%=rs("postsid")%>"/></td>
			<td class="splittd" align="left"><a href="<%=AskInstalDir%>q.asp?id=<%=rs("topicid")%>" target="_blank"><%=rs("topic")%></a>
			<td class="splittd" style="text-align:left"><a href="javascript:modifyanswer(<%=rs("postsid")%>);"><%=KS.Gottopic(KS.LoseHtml(Ubbcode(rs("content"),0)),50)%></a></td>
			<td class="splittd" align="center"><%=rs("username")%></td>
			<td class="splittd" align="center"><%=rs("PostTime")%></td>
			<td class="splittd" align="center">
			<a href="?action=verifyda&v=<%=rs("LockTopic")%>&id=<%=rs("postsid")%>"><%
			 if rs("LockTopic")=1 then
			  response.write "<font color=red>未审</font>"
			 else
			  response.write "<font color=green>已审</font>"
			 end if
			%></a></td>
			<td align="center" nowrap="nowrap">
			<a href="?action=verifyda&v=<%=rs("LockTopic")%>&id=<%=rs("postsid")%>">审核</a> | <a href="javascript:modifyanswer(<%=rs("postsid")%>);">修改</a> | <a href="?action=delanswer&id=<%=rs("postsid")%>" onClick="return(confirm('确定删除该回答吗？'))">删除</a>
			</td>
		  <%  i=i+1
		     if i>= MaxPerPage Then Exit Do
		  
		     RS.MoveNext
		    Loop
		  End If
		  %>
		  <tr>
			<td colspan="10">
			&nbsp;<b>选择：</b><a href='javascript:void(0)' onclick='javascript:Select(0)'>全选</a>  <a href='javascript:void(0)' onclick='javascript:Select(1)'>反选</a>  <a href='javascript:void(0)' onclick='javascript:Select(2)'>不选</a>
			
			&nbsp;&nbsp;
		
				<input class="button" type="submit" name="submit_button1" value="批量删除" onClick="$('action').value='del';return confirm('您确定执行该操作吗?');">
				<input type="submit" value="批量审核" class="button" onClick="$('#action').val('verifyda');$('#v').val(1);return(confirm('确定批量审核吗?'));">
				
				<input type="submit" value="批量取消审核" class="button" onClick="$('#action').val('verifyda');$('#v').val(0);return(confirm('确定批量取消审核吗?'));">
			</td>
		</tr>
		</form>
		<tr>
			<td  align="right" colspan="10" id="NextPageText">
			<%
			Call KS.ShowPage(totalput, MaxPerPage, "",pagenow,true,true)
			%>
			</td>
		</tr>
		  </table>
		  <%
		  RS.Close :Set RS=Nothing
		  
	   End Sub
	   
	   Sub verifyda()
	     dim v:v=KS.ChkClng(KS.G("V"))
		 Dim id:id=KS.FilterIds(KS.S("ID"))
		 If Id="" Then KS.AlertHintScript "请选择要操作的答案!"
		 If V=1 Then
		 Conn.Execute("Update KS_AskPosts1 Set LockTopic=0 Where postsid in(" & ID & ")")
		 Else
		 Conn.Execute("Update KS_AskPosts1 Set LockTopic=1 Where postsid in(" & ID & ")")
		 End If
		 KS.AlertHintScript ("恭喜，操作成功!")
	   End Sub
	   Sub delanswer()
	     Dim Id:Id=KS.FilterIds(KS.G("ID"))
		 If Id="" Then KS.AlertHintScript "请选择要删除的答案!"
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "select * From KS_AskPosts1 Where postsid IN(" & ID & ")",conn,1,1
		 Do While Not RS.Eof
		  If RS("satis")=1 Then
		   Conn.Execute("Update KS_AskTopic Set BestUserName='',BestUserID=0,TopicMode=0, PostNum=PostNum-1 Where TopicID=" & RS("TopicID") & " And PostNum>0")
		  Else
		   Conn.Execute("Update KS_AskTopic Set PostNum=PostNum-1 Where TopicID=" & RS("TopicID") & " And PostNum>0")
		  End If 
		 RS.MoveNext
		 Loop
		 RS.Close:Set RS=Nothing
		 Conn.Execute("Delete From KS_AskPosts1 Where postsid IN(" & ID & ")")
		 KS.AlertHintScript ("恭喜，删除回答操作成功!")
	   End Sub

	   Sub showmain()
			Dim i
			maxperpage=20
			showmode=KS.ChkClng(Request("showmode"))
			m_intOrder=KS.ChkClng(Request("order"))
			count=KS.ChkClng(Request("count"))
			classid=KS.ChkClng(Request("classid"))
			 Call GetChildList()
		
		%>
		<div style='line-height:22px;margin:5px 1px;border:1px solid #cccccc;background:#FAFAFE'>
		<table border='0' width='100%'><tr><td width='15%'>
		<%
		If IsArray(classarr) Then
			Dim K,J,N
			N=0
			 For k=0 To Ubound(classarr,2)
			    Response.Write "<tr>"
			    For J=1 To 5
			     Response.Write "<td width='15%'><img src='images/folder/folderopen.gif' align='absmiddle'><a href=""?classid=" & classarr(0,n) & """>" & classarr(1,n) & "(" & classarr(2,n)+classarr(3,n) & ")</a></td>"
				 n=n+1
				 If N>Ubound(classarr,2) Then Exit For
				Next
				Response.Write "</tr>"
			  If N>Ubound(classarr,2) Then Exit For
			Next
	  End If	
		
		%>
		
		</tr></table></div>
		
		<div style="margin-top:5px;height:25px;line-height:25px">
		<b>查看：</b> <a href="KS.AskList.asp"><font color=#999999>全部</font></a> - <a href="?showmode=1"><font color=#999999>待解决</font></a> - <a href="?showmode=2"><font color=#999999>已解决</font></a> - <a href="?showmode=3"><font color=#999999>有悬赏</font></a> - <a href="?showmode=4"><font color=#999999>未审核</font></a> - <a href="?showmode=5"><font color=#999999>已审核</font></a> <b>排序方式:</b>
				  <select name="orders" onChange="location.href='?orders='+this.value">
				  <option value="">--选择排序方式--</option>
				  <option value="TopicID Desc"<%if KS.G("orders")="TopicID Desc" Then response.write " selected"%>>最新提问</option>
				  <option value="LastPostTime Desc,TopicID Desc"<%if KS.G("orders")="LastPostTime Desc,TopicID Desc" Then response.write " selected"%>>最新回答</option>
				  <option value="Hits Desc,TopicID Desc"<%if KS.G("orders")="Hits Desc,TopicID Desc" Then response.write " selected"%>>浏览次数最多</option>
				  <option value="Reward Desc,TopicID Desc"<%if KS.G("orders")="Reward Desc,TopicID Desc" Then response.write " selected"%>>悬赏分最高</option>
				  </select>
		</div>
		<table  border="0" align="center" style='border-top:1px solid #cccccc' cellpadding="0" cellspacing="0" width="100%">
		<tr class="sort">
			<td width="5%" noWrap="noWrap">选择</td>
			<td width="56%">标题</td>
			<td width="12%" noWrap="noWrap">用户名</td>
			<td width="6%" noWrap="noWrap">状态</td>
			<%if KS.G("orders")="LastPostTime Desc,TopicID Desc" Then%>
			<td width="8%" noWrap="noWrap">回答日期</td>
			<%else%>
			<td width="8%" noWrap="noWrap">发布日期</td>
			<%end if%>
			<td width="4%" noWrap="noWrap">浏览</td>
			<td width="9%" noWrap="noWrap">管理操作</td>
		</tr>
		
		<form name="myform" id="myform" method="post" action="?">
		<input type="hidden" name="action" id="action" value="del">
		<%
			Call showTopiclist()
			If Not IsArray(Topiclist) Then
			  Response.Write "<tr><td class='splittd' colspan=6 align='center'>对不起, 找不到相关问题!</td></tr>"
			Else
				For i=0 To Ubound(Topiclist,2)
					If Not Response.IsClientConnected Then Response.End

		%>
		<tr align="center" onMouseOut="this.className='list'" onMouseOver="this.className='listmouseover'" onClick="chk_iddiv('<%=Topiclist(0,i)%>')" id='u<%=Topiclist(0,i)%>'>
			<td class="splittd"><input type="checkbox" name="id" id='c<%=Topiclist(0,i)%>' value="<%=Topiclist(0,i)%>"/></td>
			<td class="splittd" align="left">[<a href="<%=AskInstalDir%>showlist.asp?id=<%=Topiclist(1,i)%>" target="_blank"><%=Topiclist(3,i)%></a>]
			<a href="<%=AskInstalDir%>q.asp?id=<%=Topiclist(0,i)%>" target="_blank"><%=Trim(Topiclist(4,i))%></a>
			<%
			 If Topiclist(5,i)>0 then
			  response.write "<img src=" & AskInstalDir & "images/ask_xs.gif>" & TopicList(5,i) & "分"
			 end if
			 
			 If TopicList(16,i)=1 Then
			  Response.Write " <span style='color:red'>荐</span>"
			 End If
			 If TopicList(11,i)=1 Then
			  Response.Write " <span style='color:green'>未审</font>"
			 End If
			%>
			
			</td>
			<td class="splittd" noWrap="noWrap"><%=Topiclist(2,i)%></td>
			<td class="splittd" noWrap="noWrap"><a target="_blank" href="<%=AskInstalDir%>q.asp?id=<%=Topiclist(0,i)%>"><img src="<%=askInstalDir%>images/ask<%=Topiclist(13,i)%>.gif" border="0"/></a></td>
			<%if KS.G("orders")="LastPostTime Desc,TopicID Desc" Then%>
			<td class="splittd" noWrap="noWrap"><%=formatdatetime(Topiclist(10,i),2)%></td>
			<%else%>
			<td class="splittd" noWrap="noWrap"><%=formatdatetime(Topiclist(9,i),2)%></td>
			<%end if%>
			<td class="splittd" noWrap="noWrap"><%=Topiclist(17,i)%></td>
			<td class="splittd" noWrap="noWrap">
			<%if Topiclist(13,i)="1" then
			   response.write "<span style='color:#999'>回答</span>"
			  else
			   response.write "<a href='javascript:answer(" & Topiclist(0,i) & ");'>回答</a>"
			  end if
			%>
			 | <a href="?action=asked&topicid=<%=Topiclist(0,i)%>">编辑</a> | <a href="?action=del&id=<%=Topiclist(0,i)%>" onClick="return confirm('删除后将不能恢复，您确定要删除吗?')">删除</a></td>
		</tr>
		<%
				Next
			End If
			Topiclist=Null
		%>
		<tr>
			<td colspan="10">
			&nbsp;<b>选择：</b><a href='javascript:void(0)' onclick='javascript:Select(0)'>全选</a>  <a href='javascript:void(0)' onclick='javascript:Select(1)'>反选</a>  <a href='javascript:void(0)' onclick='javascript:Select(2)'>不选</a>
			
			&nbsp;&nbsp;
		
				<input class="button" type="submit" name="submit_button1" value="批量删除" onClick="$('action').value='del';return confirm('您确定执行该操作吗?');">
				<input type="submit" value="审核" class="button" onClick="$('#action').val('verify');return(confirm('确定批量审核吗?'));">
				
				<input type="submit" value="取消审核" class="button" onClick="$('#action').val('unverify');return(confirm('确定批量取消审核吗?'));">
				
				<input type="submit" value="推荐" class="button" onClick="$('#action').val('recommend');return(confirm('将问题设置为推荐将给会员增加相应的积分,确定设置吗?'));">
				
				<input type="submit" value="取消推荐" class="button" onClick="$('#action').val('unrecommend');return(confirm('为保护会员权益,取消推荐将不再扣除原设置推荐所得会员积分,确定设置吗?'));">
			</td>
		</tr>
		</form>
		<tr>
			<td  align="right" colspan="10" id="NextPageText">
			<%
			Call KS.ShowPage(totalrec, MaxPerPage, "",pagenow,true,true)
			%>
			</td>
		</tr>
		<tr> 
		   <td colspan="10">
		    <div>
			<form action="KS.AskList.asp" name="myform" method="get">
			   <div style="border:1px dashed #cccccc;margin:3px;padding:4px">
				  &nbsp;<strong>快速搜索=></strong>
				 &nbsp;关键字:<input type="text" class='textbox' name="keyword">&nbsp;分类:
				 <%
				 Dim SQL,Rs
	Response.Write " <select name=""class"">"
	Response.Write "<option value="""">所有分类</option>"
	SQL = "SELECT classid,depth,ClassName FROM KS_AskClass ORDER BY rootid,orders"
	Set Rs = Conn.Execute(SQL)
	Do While Not Rs.EOF
		Response.Write "<option value=""" & Rs("classid") & """ "
		If Request("editid") <> "" And CLng(Request("editid")) = Rs("classid") Then Response.Write "selected"
		Response.Write ">"
		If Rs("depth") = 1 Then Response.Write "&nbsp;&nbsp;├ "
		If Rs("depth") > 1 Then
			For i = 2 To Rs("depth")
				Response.Write "&nbsp;&nbsp;│"
			Next
			Response.Write "&nbsp;&nbsp;├ "
		End If
		Response.Write Rs("ClassName") & "</option>" & vbCrLf
		Rs.movenext
	Loop
	Rs.Close
	Response.Write "</select>"
	Set Rs = Nothing
%>
				  &nbsp;
				  问题状态:<select name="showmode">
				  <option value="0">全部</option>
				  <option value="1">待解决</option>
				  <option value="2">待解决</option>
				  <option value="3">有悬赏</option>
				  <option value="4">未审核</option>
				  <option value="5">已审核</option>
				  </select>
				  
				  排序方式
				  <select name="orders">
				  <option value="TopicID Desc">最新提问</option>
				  <option value="LastPostTime Desc,TopicID Desc">最新回答</option>
				  <option value="Hits Desc,TopicID Desc">浏览次数最多</option>
				  <option value="Reward Desc,TopicID Desc">悬赏分最高</option>
				  </select>
				  <div style="padding-left:83px">
				  提问时间:从
		      <input name="StartDate" onClick="popUpCalendar(this, this, dateFormat,-1,-1)" type="text" class="textbox" id="StartDate" value="<%=request("StartDate")%>" readonly style="width:12%">
		  <b><a href="javascript:;" onClick="popUpCalendar($('#StartDate')[0], $('#StartDate')[0], dateFormat,-1,-1)"><img src="Images/date.gif" border="0" align="absmiddle" title="选择日期"></a></b>
		      到
		        <input name="EndDate" type="text" onClick="popUpCalendar(this, this, dateFormat,-1,-1)" id="EndDate" class="textbox"  value="<%=request("endDate")%>" readonly style="width:12%">
		       <b><a href="javascript:;" onClick="popUpCalendar($('#EndDate')[0], $('#EndDate')[0], dateFormat,-1,-1)"><img src="Images/date.gif" border="0" align="absmiddle" title="选择日期"></a></b> 
				  提问者:<input type="text" name="askName" class="textbox" size="12">
				  回答者:<input type="text" name="answerName" class="textbox" size="12">
			   
				  <input type="submit" value="开始搜索" class="button" name="s1">
				  </div>
				</div>
			</form>
			</div>

		   </td>
		</tr>
		<tr>
			<td colspan="6" style="border:1px solid #f1f1f1;line-height:21px;padding-left:10px">
			 <font color=red><strong>操作说明:</strong></font><br />
			 1.将问题设置为推荐将给会员增加相应的积分,会员所得积分在"问答参数设置"里设定<br />
			 2.为保护会员权益,取消推荐将不再扣除原设置推荐所得会员积分,一般建议一旦设置为推荐后就不要再取消推荐<br />
			 3.如果您将问题推荐后,然后取消推荐,又重新推荐可能导致多次给会员增加积分
			</td>
		</tr>
		
		
		</table>
		<%
		End Sub
		
		Sub showTopiclist()
			Dim Rs,SQL,Cmd,Param,OrderStr
			SQLField="TopicID,classid,UserName,classname,title,reward,Expired,Closed,PostTable,DateAndTime,LastPostTime,LockTopic,PostNum,TopicMode,Anonymous,IsTop,recommend,Hits"
			Param=" where 1=1"
			Select Case showmode
			 case 1 param=param & " and topicmode=0"
			 case 2 param=param & " and topicmode=1"
			 case 3 param=param & " and reward>0"
			 case 4 param=param & " and locktopic=1"
			 case 5 param=param & " and locktopic=0"
			end select
			If Classid>0 Then param=param & " and classid in(select classid from KS_askclass where ','+parentstr +'' like '%," & classid & ",%')"
			If KS.G("keyword")<>"" Then param=param & " and title like '%" & Trim(KS.G("KeyWord")) & "%'"
			If KS.G("Class")<>"" Then Param=Param & " and classid=" & KS.ChkClng(KS.G("Class"))
			If KS.G("askName")<>"" Then Param=Param &" and username like '%" & Trim(KS.G("askName")) & "%'"
			If KS.G("answerName")<>"" Then Param=Param &" and topicid in(select topicid from KS_AskPosts1 Where UserName like '%" & Trim(KS.G("answerName")) & "%')"
			if Request("StartDate")<>"" and isdate(request("StartDate")) then
			  Param=Param & " and DateAndTime>=#" & request("StartDate") & "#"
			end if
			if Request("endDate")<>"" and isdate(request("endDate")) then
			 Dim enddate:EndDate = DateAdd("d", 1, Request("EndDate"))
			  Param=Param & " and DateAndTime<=#" & enddate & "#"
			end if
			If KS.G("orders")<>"" Then
			 OrderStr=" Order By " & KS.G("orders")
			Else
			 OrderStr=" Order By TopicID Desc"
			End If
			
			If count=0 Then
				totalrec=Conn.Execute("SELECT COUNT(*) FROM KS_AskTopic "&Param&"")(0)
			Else
				totalrec=count
			End If
			Set Rs=KS.InitialObject("ADODB.Recordset")
			SQL="SELECT "& SQLField &" FROM [KS_AskTopic]  "&Param&OrderStr
			Rs.Open SQL,Conn,1,1
			If Not Rs.EOF Then
			   If (pagenow - 1) * MaxPerPage < totalrec Then	Rs.Move (pagenow-1) * maxperpage
				Topiclist=Rs.GetRows(maxperpage)
			Else
				Topiclist=Null
			End If
			Rs.close()
			Set Rs=Nothing
			
			Pcount = CLng(totalrec / maxperpage)
			If Pcount < totalrec / maxperpage Then Pcount = Pcount + 1
			If pagenow>Pcount Then pagenow=1

		End Sub

		
	Sub GetChildList()
		   If Not IsObject(Application(KS.SiteSN&"_askclasslist")) Then KS.LoadCategoryList
		   Set Catelist = Application(KS.SiteSN&"_askclasslist")
		   If Not Catelist Is Nothing Then
			Dim Node:Set Node=Catelist.documentElement.selectSingleNode("row[@classid="&classid&"]")
			If Not Node Is Nothing Then
				child=Node.selectSingleNode("@child").text
				If child>0 Then
					cid=classid
				Else
					cid=CLng(Node.selectSingleNode("@parentid").text)
				End If 
			Else
			  cid=0
			End If
		   Else
		     cid=0
		   End If
		
		  Dim SQLStr:SQLStr = "SELECT classid,classname,AskPendNum,AskDoneNum FROM KS_AskClass WHERE parentid="&KS.ChkClng(cid)&" ORDER BY orders,classid"
		  Dim RS:Set RS=Conn.Execute(SQLStr)
		  If Not RS.Eof Then
		   classarr=RS.GetRows(-1)
		  End If
		  RS.Close:Set RS=Nothing
		End Sub
		%>
		<!--#include file="../ks_cls/ubbfunction.asp"-->
		<%
		Sub showAsked()
			Dim Rs,SQL,XMLDom,Node,i
			Dim PostUserTitle,DelAction
			topicid=KS.ChkClng(Request("topicid"))
			SQL="SELECT TopicID,classid,username,classname,title,Expired,Closed,PostTable,DateAndTime,LastPostTime,ExpiredTime,LockTopic,Reward,Hits,PostNum,CommentNum,TopicMode,Broadcast,Anonymous,supplement FROM KS_AskTopic WHERE topicid="&topicid
			Set Rs = Conn.Execute(SQL)
			If Rs.BOF And Rs.EOF Then
				Set Rs = Nothing
				 KS.AlertHintScript "错误的系统参数"
				Exit Sub
			End If
			Set XMLDom = KS.RsToxml(Rs,"topic","xml")
			Set Rs = Nothing
			Set Node = XMLDom.documentElement.selectSingleNode("topic")
			If Not Node Is Nothing Then
				topicid = CLng(Node.selectSingleNode("@topicid").text)
				classid = CLng(Node.selectSingleNode("@classid").text)
				topicmode = CLng(Node.selectSingleNode("@topicmode").text)
				PostNum = CLng(Node.selectSingleNode("@postnum").text)
				ExpiredTime = CDate(Node.selectSingleNode("@expiredtime").text)
				CommentNum = CLng(Node.selectSingleNode("@commentnum").text)
				HeadTitle = Trim(Node.selectSingleNode("@title").text)
				TopicUseTable = Trim(Node.selectSingleNode("@posttable").text)
			End If
			Set Node = Nothing
			Set XMLDom = Nothing
		%>
		<script type="text/javascript" src="../editor/ckeditor.js" mce_src="../editor/ckeditor.js"></script>
		<script type="text/javascript" src="../ks_Inc/jquery.js"></script>
		<script type="text/javascript">
		 function replaceCk(i){
		  CKEDITOR.replace('content'+i, {width:"99%",height:"100px",toolbar:"Basic",filebrowserBrowseUrl :"../Include/SelectPic.asp?from=ckeditor&Currpath=<%=KS.GetUpFilesDir()%>",filebrowserWindowWidth:650,filebrowserWindowHeight:290});
		  $("#d"+i).fadeOut('fast');
		  $("#button1"+i).hide();
		  $("#button2"+i).show();
		 }
		</script>
		<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1">
		<tr>
			<th>问题：<%=HeadTitle%></th>
		</tr>
		<%
			Call showAskedlist()
			If IsArray(Topiclist) Then
				For i=0 To Ubound(Topiclist,2)
					If Not Response.IsClientConnected Then Response.End
					If CLng(Topiclist(12,i))=0 Then
						PostUserTitle="提问者："
						DelAction="del"
					Else
						PostUserTitle="回答者："
						DelAction="delask"
					End If
		%>
		
		<tr>
			<td class="tdbg">
			  <table border="0" width="100%" <%If TopicList(10,i) = 1 Then Response.Write " style='border:5px solid #ff6600;'"%>>
<tr>
			<td colspan=2  class="clefttitle" height="30">
				<%=PostUserTitle%>:<%=Topiclist(3,i)%>  
				&nbsp;&nbsp;&nbsp;
				时间:<%=TopicList(7,i)%><%If TopicList(10,i) = 1 Then response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color=red size=2><strong>最佳答案</strong></font>"%>
				</td>
		</tr>			  
			   <form action="?i=<%=i%>&action=save&postsid=<%=Topiclist(0,i)%>&topicid=<%=Topiclist(2,i)%>" method="post">
			   
			   <tr>
			    <td width="600">
				 <%if i=0 then%>
				    标题:<input type="text" name="title" class="textbox" value="<%=TopicList(4,i)%>">
					分类:
					
			<%  dim ii
				Response.Write " <select name=""classid"">"
				Response.Write "<option value=""0"">做为一级分类</option>"
				SQL = "SELECT classid,depth,ClassName FROM KS_AskClass ORDER BY rootid,orders"
				Set Rs = Conn.Execute(SQL)
				Do While Not Rs.EOF
					Response.Write "<option value=""" & Rs("classid") & """ "
					If  CLng(classid) = Rs("classid") Then Response.Write "selected"
					Response.Write ">"
					If Rs("depth") = 1 Then Response.Write "&nbsp;&nbsp;├ "
					If Rs("depth") > 1 Then
						For ii = 2 To Rs("depth")
							Response.Write "&nbsp;&nbsp;│"
						Next
						Response.Write "&nbsp;&nbsp;├ "
					End If
					Response.Write Rs("ClassName") & "</option>" & vbCrLf
					Rs.movenext
				Loop
				Rs.Close
				Response.Write "</select>"
				Set Rs = Nothing
			%>
					
				 <%end if%>
				 
				 审核
				 <input type="radio" name="LockTopic" value="0"<%If TopicList(11,i) = 0 Then Response.Write " checked=""checked"""%> /> 确定审核&nbsp;&nbsp;
				<input type="radio" name="LockTopic" value="1"<%If TopicList(11,i) = 1 Then Response.Write " checked=""checked"""%> /> 取消审核
				
				<br />
				
				    <div id="d<%=i%>">
				   <%=ubbcode(Topiclist(5,i),i)%>
				   </div>
			        <textarea name="content<%=i%>" id="content<%=i%>" style="display:None;width:600px;height:80px"><%=server.Htmlencode(Topiclist(5,i))%></textarea>
					
				
			    </td>
				<td width="200" align="center">
				<input type="button" value=" 编 辑" class="button" onClick="replaceCk(<%=i%>)" id="button1<%=i%>">
			<input type="submit" value=" 保 存 " class="button"  style="display:none" id="button2<%=i%>"> 
			<%If TopicList(10,i) <> 1 Then%>
			<input type="button" value=" 删 除 " class="button" onClick="if (confirm('删除后将不能恢复，您确定要删除吗?')){location.href='KS.AskList.asp?action=<%=DelAction%>&postsid=<%=Topiclist(0,i)%>&topicid=<%=Topiclist(2,i)%>'}">
			<%end if%>
			<%If topicmode=0 and i<>0 then%>
			<br /><br/><input type="button" value=" 采纳为最佳答案 " class="button" onClick="if (confirm('您确定采纳该答案为最佳答案吗?')){location.href='KS.AskList.asp?action=SetSatis&postsid=<%=Topiclist(0,i)%>&topicid=<%=Topiclist(2,i)%>'}">
			<%end if%>
			    </td>
			  </tr>
			  </form>
			  </table>
			</td>
		</tr>
		<%
				Next
			End If
			Topiclist=Null
		%>
		<tr>
			<td class="tablerow1" align="right" id="NextPageText">
			<%
			Call KS.ShowPage(totalrec, MaxPerPage, "", pagenow,true,true)
			%>
			</td>
		</tr>
		</table>
		<%
		End Sub
		
		Sub showAskedlist()
			Dim Rs,SQL
			maxperpage=10
			
			SQLField="postsid,classid,TopicID,UserName,topic,content,addText,PostTime,DoneTime,star,satis,LockTopic,PostsMode,VoteNum,Plus,Minus,PostIP,Report"
			If count=0 Then
				totalrec=Conn.Execute("SELECT COUNT(*) FROM ["&TopicUseTable&"] WHERE topicid="&topicid&" "&SQLQuery&"")(0)
			Else
				totalrec=count
			End If
			Set Rs=Server.CreateObject("ADODB.Recordset")
			SQL="SELECT "& SQLField &" FROM ["&TopicUseTable&"]  WHERE topicid="&topicid&" "&SQLQuery&" ORDER BY postsMode ASC,Satis desc,postsid"
			Rs.Open SQL,Conn,1,1
			If Not Rs.EOF Then
			   If (pagenow - 1) * MaxPerPage < totalrec Then Rs.Move (pagenow-1) * maxperpage
				Topiclist=Rs.GetRows(maxperpage)
			Else
				Topiclist=Null
			End If
			
			Rs.close()
			Set Rs=Nothing
		
			Pcount = CLng(totalrec / maxperpage)
			If Pcount < totalrec / maxperpage Then Pcount = Pcount + 1
			If pagenow>Pcount Then pagenow=1
			pagelinks="KS.AskList.asp.asp?action=asked&topicid="&topicid&"&count="&totalrec&"&"
		End Sub
		
		
		Sub saveAsked()
			Dim Rs,SQL,postsid
			Dim TextContent,satis,LockTopic,strTitle,star
			postsid=KS.ChkClng(Request("postsid"))
			topicid=KS.ChkClng(Request("topicid"))
			If Trim(Request.Form("content"&request("i")))="" Then
				Call KS.AlertHintScript("内容不能为空!")
				Exit Sub
			End If
			SQL="SELECT top 1 TopicID,classid,title,Username,Expired,Closed,PostTable,LockTopic,TopicMode,supplement FROM KS_AskTopic WHERE topicid="&topicid
			Set Rs = Conn.Execute(SQL)
			If Rs.BOF And Rs.EOF Then
				Set Rs = Nothing
				KS.AlertHintScript "错误的系统参数"
				Exit Sub
			End If
			topicid=Rs("TopicID")
			strTitle=Rs("title")
			TopicUseTable=Trim(Rs("PostTable"))
			TopicMode=Rs("TopicMode")
			Set Rs = Nothing
			TextContent=Request.Form("content"&request("i"))
			LockTopic=KS.ChkClng(Request.Form("LockTopic"))
			Conn.Execute ("UPDATE ["&TopicUseTable&"] SET content='"&TextContent&"',LockTopic="&LockTopic&" WHERE postsid="&postsid&" And topicid="&topicid)
			If KS.G("I")="0" Then
			 dim className:className=LFCls.GetSingleFieldValue("select top 1 classname from [KS_AskClass] Where ClassID=" & KS.ChkClng(KS.G("ClassID")))
			Conn.Execute ("UPDATE [KS_AskTopic] SET className='" & className&"',ClassID=" & KS.ChkClng(KS.G("ClassID")) & ",LockTopic="&LockTopic&" WHERE topicid="&topicid)
			Conn.Execute ("UPDATE [KS_AskAnswer] SET className='" & className&"',ClassID=" & KS.ChkClng(KS.G("ClassID")) & " WHERE topicid="&topicid)
			Conn.Execute ("UPDATE ["&TopicUseTable&"] SET ClassID=" & KS.ChkClng(KS.G("ClassID")) & " WHERE topicid="&topicid)
			End If
			
			If strTitle<>Request.Form("title") and trim(Request.Form("title"))<>"" Then
				Conn.Execute ("UPDATE ["&TopicUseTable&"] SET topic='"&Trim(Request.Form("title"))&"' WHERE topicid="&topicid)
				Conn.Execute ("UPDATE [KS_AskTopic] SET title='"&Trim(Request.Form("title"))&"' WHERE topicid="&topicid)
				Conn.Execute ("UPDATE [KS_AskAnswer] SET title='"&Trim(Request.Form("title"))&"' WHERE topicid="&topicid)
			End If
			Call KS.AlertHintScript("恭喜您！编辑/审核问题成功。")
		End Sub
		
		'推荐问题
		Sub Recommend()
			Dim TopicIDlist,SQL,RS,ScoreToQuestionUser,ScoreToAnswerUser
			TopicIDlist=KS.FilterIds(Request("id"))
			If TopicIDlist="" Then 	Call KS.AlertHintScript("你没有选择问题!"):Response.End
			ScoreToQuestionUser=KS.ChkClng(KS.ASetting(33))
			ScoreToAnswerUser=KS.ChkClng(KS.ASetting(34))
			SQL="SELECT * FROM KS_AskTopic Where recommend=0 and TopicID in(" & TopicIDList & ")"
			Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open SQL,Conn,1,3
			Do While Not RS.Eof
			  
			  RS("Recommend")=1
			  RS.Update
			  
			   '给提问者加积分
			  If ScoreToQuestionUser>0 Then
				 Call KS.ScoreInOrOut(RS("UserName"),1,ScoreToQuestionUser,"系统","问吧问题[" & rs("title") & "]被管理员推荐!",0,0)
			  End If
			   '给最佳回答者加积分
			  If ScoreToAnswerUser>0 Then
			     Dim rsb:set rsb=Conn.Execute("select username From KS_AskAnswer Where TopicID=" & RS("TopicID") & " and AnswerMode=1")
				 if not rsb.eof then
				 Call KS.ScoreInOrOut(rsb(0),1,ScoreToAnswerUser,"系统","问吧问题[" & rs("title") & "]最佳答案被管理员推荐!",0,0)
				 end if
				 rsb.close:set rsb=nothing
			  
			  End If
			  
			  RS.MoveNext
			Loop
			RS.Close
			Set RS=Nothing
			Response.Redirect Request.ServerVariables("HTTP_REFERER")
		End Sub
		
		'取消推荐问题
		Sub UnRecommend()
			Dim TopicIDlist,SQL,RS
			TopicIDlist=KS.FilterIds(Request("id"))
			If TopicIDlist="" Then 	Call KS.AlertHintScript("你没有选择问题!"):Response.End
			SQL="SELECT * FROM KS_AskTopic Where recommend=1 and TopicID in(" & TopicIDList & ")"
			Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open SQL,Conn,1,3
			Do While Not RS.Eof
			  RS("Recommend")=0
			  RS.Update
			  RS.MoveNext
			Loop
			RS.Close
			Set RS=Nothing
			Response.Redirect Request.ServerVariables("HTTP_REFERER")
		End Sub

        '批量审核
		Sub Verify()
			Dim TopicIDlist,SQL,RS
			TopicIDlist=KS.FilterIds(Request("id"))
			If TopicIDlist="" Then 	Call KS.AlertHintScript("你没有选择问题!"):Response.End
			Conn.Execute("Update KS_AskTopic Set LockTopic=0 Where TopicID in(" & TopicIDList & ")")
			Conn.Execute("Update KS_AskPosts1 Set LockTopic=0 Where PostsMode=0 and TopicID in(" & TopicIDList & ")")
			Response.Redirect Request.ServerVariables("HTTP_REFERER")
		End Sub
        '取消审核
		Sub UnVerify()
			Dim TopicIDlist,SQL,RS
			TopicIDlist=KS.FilterIds(Request("id"))
			If TopicIDlist="" Then 	Call KS.AlertHintScript("你没有选择问题!"):Response.End
			Conn.Execute("Update KS_AskTopic Set LockTopic=1 Where TopicID in(" & TopicIDList & ")")
			Conn.Execute("Update KS_AskPosts1 Set LockTopic=1 Where PostsMode=0 and TopicID in(" & TopicIDList & ")")
			Response.Redirect Request.ServerVariables("HTTP_REFERER")
		End Sub
		
		'设为最佳答案
		Sub SetSatis()
		   Dim Rs,SQL,i,SQLArry,postsid,ClassID
			Dim TopicID,userName,k,TopicUseTable,BestUserName,BestUserId,UserReward
			TopicID=KS.ChkClng(Request("topicid"))
            Postsid=KS.ChkClng(Request("postsid"))
			SQL="SELECT TopicID,userName,PostTable,TopicMode,classid,Reward FROM KS_AskTopic WHERE topicid="&TopicID
			Set Rs = Conn.Execute(SQL)
			If Rs.BOF And Rs.EOF Then
				Set Rs=Nothing
				Call KS.AlertHintScript("错误的系统参数!")
				Response.End
			End If
			TopicUseTable=Rs(2)
			UserName=Rs(1)
			ClassID=RS(4)
			UserReward=RS(5)
			Set Rs=Nothing
			
			Set Rs = Server.CreateObject("ADODB.Recordset")
			SQL = "SELECT postsid,TopicID,username,topic FROM ["&TopicUseTable&"] WHERE topicid="&topicid&" and PostsMode=1 and LockTopic=0 and postsid="& Postsid
			Rs.Open SQL,Conn,1,1
			If Rs.BOF And Rs.EOF Then
				Set Rs = Nothing
				Response.Write "<script>alert('友情提示!\n\n请选择正确的问题ID!');history.back();</script>"
				Response.End
			Else
				Do While Not Rs.EOF
					Conn.Execute ("UPDATE ["&TopicUseTable&"] SET satis=1,DoneTime="& SqlNowString &" WHERE postsid="& Rs(0))
					BestUserName=RS(2)
					if UserReward>0 then
				     Call KS.ScoreInOrOut(Rs(2),1,UserReward,"系统","问吧回答问题[" & rs("topic") & "]被采纳悬赏!",0,0)
					end if
					If KS.ChkClng(KS.ASetting(31))>0 Then
				    Call KS.ScoreInOrOut(Rs(2),1,KS.ChkClng(KS.ASetting(31)),"系统","您的对问题[" & rs("topic") & "]的回答被设为最佳答案!",0,0)
					End If

					Conn.Execute ("UPDATE KS_AskAnswer SET AnswerMode=1 WHERE topicid="&topicid&" and username='"& Rs(2) & "'")
					Rs.movenext
				Loop
			End If
			Rs.Close
			If BestUserName<>"" Then
			  RS.Open "select top 1 userid from ks_user where username='" & BestUserName &"'",conn,1,1
			  if not rs.eof then
			   BestUserId=rs(0)
			  end if
			  rs.close
			End If
			Set Rs = Nothing
		
			Conn.Execute ("UPDATE KS_AskTopic SET BestUserName='" & BestUserName & "',BestUserId=" & KS.ChkClng(BestUserId)&",LastPostTime="& SqlNowString &",TopicMode=1 WHERE topicid="&topicid&" and username='"& UserName &"' and Closed=0 and LockTopic=0")
			Conn.Execute ("UPDATE KS_AskAnswer SET TopicMode=1 WHERE topicid="&topicid)
			
			'Conn.Execute ("UPDATE KS_User SET Score=Score+" & KS.ChkClng(KS.ASetting(32)) & " WHERE username='"& UserName & "'")
			Conn.Execute ("UPDATE KS_AskClass SET AskPendNum=AskPendNum-1,AskDoneNum=AskDoneNum+1 WHERE classid="& classid)
			Call KS.Alert("恭喜您！设置最佳答案成功!","KS.AskList.asp?action=asked&topicid=" & topicid)
		End Sub
		
		
		Sub delTopic()
			Dim Rs,SQL,i,SQLArry
			Dim TopicIDlist,userName,k
			Dim MinusPoints,ClassNumStr,parentArr
			TopicIDlist=KS.FilterIds(Request("id"))
			If TopicIDlist="" Then 	Call KS.AlertHintScript("你没有选择问题!"):Response.End

			SQL="SELECT TopicID,userName,PostTable,TopicMode,classid,title FROM KS_AskTopic WHERE topicid in("&TopicIDlist&")"
			Set Rs = Conn.Execute(SQL)
			If Rs.BOF And Rs.EOF Then
				Set Rs=Nothing
				Call KS.AlertHintScript("错误的系统参数!")
				Response.End
			End If
			SQLArry=Rs.GetRows(-1)
			Set Rs=Nothing
		
			If IsArray(SQLArry) Then
				For i=0 To Ubound(SQLArry,2)
					topicid=CLng(SQLArry(0,i))
					userName=SQLArry(1,i)
					TopicUseTable=Trim(SQLArry(2,i))
					TopicMode=CLng(SQLArry(3,i))
					parentArr=split(conn.execute("select parentstr from KS_askclass where classid=" & SQLArry(4,i))(0),",")
					Select Case TopicMode
						Case 1
							MinusPoints=KS.ChkCLng(KS.ASetting(39))
							ClassNumStr="AskDoneNum=AskDoneNum-1 Where AskDoneNum>0"
						Case Else
							MinusPoints=KS.ChkClng(KS.ASetting(40))
							ClassNumStr="AskPendNum=AskPendNum-1 Where AskPendNum>0"
					End Select
					Conn.Execute("DELETE FROM KS_UploadFiles WHERE channelid=1032 and infoid in(select postsid from "&TopicUseTable&" WHERE topicid="&topicid & ")")
					Conn.Execute("DELETE FROM KS_AskTopic WHERE topicid="&topicid)
					Conn.Execute("DELETE FROM KS_AskAnswer WHERE topicid="&topicid)
					Conn.Execute("DELETE FROM "&TopicUseTable&" WHERE topicid="&topicid)
					For K=0 To Ubound(parentarr)-1
					Conn.Execute("Update KS_AskClass Set " & ClassNumStr & " and classid=" & parentarr(k))
					Next
					
					If TopicMode=0 Then
					 If KS.ChkClng(KS.ASetting(39))<>0 Then
					  Call KS.ScoreInOrOut(UserName,2,KS.ChkClng(KS.ASetting(39)),"系统","问吧的问题[" & SQLArry(5,i) & "]被删除!",0,0)
					 End If
					Else
					  If KS.ChkClng(KS.ASetting(40))<>0 Then
					  Call KS.ScoreInOrOut(UserName,2,KS.ChkClng(KS.ASetting(40)),"系统","问吧的问题[" & SQLArry(5,i) & "]被删除!",0,0)
					 End If
					End If
					
				Next
				SQLArry=Null
			End If
			if instr(lcase(REQUEST.SERVERVARIABLES("HTTP_REFERER")),"index.asp")=0 then
			Call KS.AlertHintScript("恭喜您！数据删除成功!")
			else
			Call KS.Alert("恭喜您！数据删除成功!","KS.AskList.asp")
			end if
		End Sub
		
		Sub delAsked()
			Dim Rs,SQL,postsid
			Dim SQLArry,userName,PostNum,Title
			Dim MinusPoints,MinusExperience
			Dim satis,PostsMode
			postsid=KS.ChkClng(Request("postsid"))
			topicid=KS.ChkClng(Request("topicid"))
			SQL="SELECT TopicID,username,PostTable,TopicMode,PostNum,Title FROM KS_AskTopic WHERE topicid="&topicid
			Set Rs = Conn.Execute(SQL)
			If Rs.BOF And Rs.EOF Then
				Set Rs=Nothing
				Call KS.AlertHistory("错误的系统参数!",-1)
				Response.End
			End If
			SQLArry=Rs.GetRows(1)
			Set Rs=Nothing
			If IsArray(SQLArry) Then
				topicid=CLng(SQLArry(0,0))
				userName=SQLArry(1,0)
				TopicUseTable=Trim(SQLArry(2,0))
				TopicMode=CLng(SQLArry(3,0))
				PostNum=CLng(SQLArry(4,0))
				Title=SQLArry(5,0)
			Else
				Call KS.AlertHintScript("错误的系统参数!")
				Response.End
			End If
			SQLArry=Null
			If PostNum>0 Then
				SQL="SELECT postsid,username,satis,PostsMode FROM "&TopicUseTable&" WHERE postsid="&postsid
				Set Rs = Conn.Execute(SQL)
				If Rs.BOF And Rs.EOF Then
					Set Rs=Nothing
					Call KS.AlertHintScript("错误的系统参数!")
				    Response.End
				End If
				SQLArry=Rs.GetRows(1)
				Set Rs=Nothing
				If IsArray(SQLArry) Then
					postsid=CLng(SQLArry(0,0))
					username=SQLArry(1,0)
					satis=CLng(SQLArry(2,0))
					PostsMode=CLng(SQLArry(3,0))
					If satis=0 Then
						MinusPoints=KS.ChkCLng(KS.ASetting(37))
					Else
						MinusPoints=KS.ChkClng(KS.ASetting(38))
					End If
					If PostsMode>0 Then
						Conn.Execute("DELETE FROM KS_AskAnswer WHERE topicid="&topicid&" And username='"&username&"' And AnswerNum<2")
						Conn.Execute("DELETE FROM "&TopicUseTable&" WHERE postsid="&postsid)
						'Conn.Execute ("UPDATE KS_User SET score=score-"&MinusPoints&" WHERE username='"&username & "'")
						Conn.Execute ("UPDATE KS_AskAnswer SET AnswerNum=AnswerNum-1 WHERE topicid="&topicid&" And username='"&username & "'")
						
						if MinusPoints<>0 then
						  Call KS.ScoreInOrOut(UserName,2,MinusPoints,"系统","问吧对问题[" & Title & "]的回答被删除!",0,0)

						end if
						if satis=1 then
						 Conn.Execute("update KS_AskTopic Set TopicMode=0 WHERE topicid="&topicid)
						end if
						
					End If
				End If
				SQLArry=Null
			End If
			Call KS.Alert("恭喜您！数据删除成功!","KS.AskList.asp?action=asked&topicid=" & topicid)
		End Sub
		
		Sub modifyAnswer()
		%>
		<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
        <html xmlns="http://www.w3.org/1999/xhtml">
		<head>
		<link href="Include/Admin_Style.CSS" rel="stylesheet" type="text/css">
		<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
		<script src="../KS_Inc/common.js" language="JavaScript"></script>
		<script src="../KS_Inc/jquery.js" language="JavaScript"></script>
        <script type="text/javascript" src="../editor/ckeditor.js" ></script>	
		<body>
		 <%
		  Dim ID:ID=KS.ChkClng(Request("ID"))
		  If ID=0 Then KS.Die "error!"
		  Dim RS,PostTable,Title,Content,LockTopic
		  Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open "select top 1 * from KS_AskPosts1 Where postsid=" & ID,conn,1,1
		  If RS.Eof And RS.Bof Then
		    rs.close:set rs=nothing
			ks.die "error!"
		  End If
		  Title=RS("Topic") : LockTopic= RS("LockTopic")
		  Content=RS("Content")
		  RS.Close
		 %>
	 <table width="100%" class="ctable">
		  <tr>
		    <td class="clefttitle" width="120"><strong>问题：</strong></td><td><%=title%></td>
		  </tr>

		  <iframe src="about:blank" name="hidifame" width="0" height="0"></iframe>
		  <form name="myform" method="post" action="KS.AskList.asp" target="hidifame">
		  <input type="hidden" name="action" value="DoModifyAnswerSave">
		  <input type="hidden" name="ID" value="<%=ID%>">
		  <tr>
		    <td class="clefttitle"><strong>回答内容：</strong></td><td><textarea  ID='Content' name='Content' cols=90 rows=6 style='display:none'><%=Content%></textarea><script type="text/javascript">CKEDITOR.replace('Content', {width:"98%",height:"160px",toolbar:"Basic",filebrowserBrowseUrl :"Include/SelectPic.asp?from=ckeditor&Currpath=<%=KS.GetUpFilesDir%>",filebrowserWindowWidth:650,filebrowserWindowHeight:290});</script>
			 设置为最佳答案：
			 <input type="radio" <%IF LockTopic="0" THEN response.write " checked"%> value="0" name="LockTopic">已审核
			 <input type="radio"<%IF LockTopic="1" THEN response.write " checked"%> value="1" name="LockTopic" checked>未审核
			</td>
		  </tr>
		  <tr>
		    <td  colspan=2 style="text-align:center"><input type="submit"  value="保存回答" class="button"/>
			<input type="button" onClick="parent.closeWindow()" value="关闭取消" class="button"/>
			</td>
		  </tr>
		  </form>
		 </table>

		</body>
		</html>
		<%
		End Sub
		
		Sub DoModifyAnswerSave()
		  Dim ID:ID=KS.ChkClng(KS.G("ID"))
		  If ID=0 Then KS.AlertHintScript "出错啦!"
		  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open "Select top 1 * From KS_AskPosts1 Where PostsID=" & ID,conn,1,3
		  If RS.EOf And RS.Bof Then
		    RS.Close:Set RS=Nothing
			 KS.AlertHintScript "出错啦!"
		  End If
		    RS("Content")=Request.Form("Content")
			RS("LockTopic")=KS.ChkClng(Request.Form("LockTopic"))
			RS.Update
		  RS.Close:SET RS=Nothing
	      Response.Write "<script>alert('恭喜，答案修改成功!');top.MainFrame.location.reload();top.closeWindow();</script>" 
		End Sub
		
		Sub QuickAnswer()
		%>
		<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
        <html xmlns="http://www.w3.org/1999/xhtml">
		<head>
		<link href="Include/Admin_Style.CSS" rel="stylesheet" type="text/css">
		<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
		<script src="../KS_Inc/common.js" language="JavaScript"></script>
		<script src="../KS_Inc/jquery.js" language="JavaScript"></script>
        <script type="text/javascript" src="../editor/ckeditor.js" ></script>	
		<body>
		 <%
		  Dim ID:ID=KS.ChkClng(Request("ID"))
		  If ID=0 Then KS.Die "error!"
		  Dim RS,PostTable,Title,Content
		  Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open "select top 1 * from KS_AskTopic Where TopicID=" & ID,conn,1,1
		  If RS.Eof And RS.Bof Then
		    rs.close:set rs=nothing
			ks.die "error!"
		  End If
		  PostTable=rs("PostTable")
		  Title=RS("Title")
		  RS.CLose
		  RS.Open "select top 1 * From " & PostTable & " Where TopicID=" & ID & " And PostsMode=0 order by postsid",conn,1,1
		  If RS.Eof And RS.Bof Then
		    RS.Close :Set RS=Nothing
			ks.die "error!"
		  End If
		  Content=RS("Content")
		  RS.Close
		 %>
	 <table width="100%" class="ctable">
		  <tr>
		    <td class="clefttitle" width="120"><strong>标题：</strong></td><td><%=title%></td>
		  </tr>
		  <tr>
		    <td class="clefttitle"><strong>内容：</strong></td><td><%=content%></td>
		  </tr>
		  <iframe src="about:blank" name="hidifame" width="0" height="0"></iframe>
		  <form name="myform" method="post" action="KS.AskList.asp" target="hidifame">
		  <input type="hidden" name="action" value="DoAnswerSave">
		  <input type="hidden" name="ID" value="<%=ID%>">
		  <tr>
		    <td class="clefttitle"><strong>回答：</strong></td><td><textarea  ID='Content' name='Content' cols=90 rows=6 style='display:none'></textarea><script type="text/javascript">CKEDITOR.replace('Content', {width:"98%",height:"160px",toolbar:"Basic",filebrowserBrowseUrl :"Include/SelectPic.asp?from=ckeditor&Currpath=<%=KS.GetUpFilesDir%>",filebrowserWindowWidth:650,filebrowserWindowHeight:290});</script>
			 设置为最佳答案：
			 <input type="radio" value="1" name="bestanswer" checked>是
			 <input type="radio" value="0" name="bestanswer">否
			</td>
		  </tr>
		  <tr>
		    <td  colspan=2 style="text-align:center"><input type="submit"  value="保存回答" class="button"/>
			<input type="button" onClick="parent.closeWindow()" value="关闭取消" class="button"/>
			</td>
		  </tr>
		  </form>
		 </table>

		</body>
		</html>
		<%
		End Sub
		
		Sub DoAnswerSave()
		 Dim TopicID:TopicID=KS.ChkClng(Request("id"))
		 Dim bestanswer:bestanswer=KS.ChkClng(Request("bestanswer"))
		 Dim Content:Content=Request.Form("content")
		 If KS.IsNUL(Content) Then KS.Die "<script>alert('您没有输入回答内容!');</script>"
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open "select top 1 * from KS_AskTopic Where TopicID=" & TopicID,conn,1,1
		  If RS.Eof And RS.Bof Then
		    rs.close:set rs=nothing
			KS.Die "<script>alert('不存在!');</script>"
		  End If
		 Dim PostTable,SQL,ClassID,ClassName,PostUsername,AskTopic,PostsId
		 PostTable=RS("PostTable")
         ClassID=RS("ClassiD"):ClassName=rs("ClassName")
		 PostUsername=RS("UserName")
		 AskTopic=RS("Title")
		 RS.Close
		 
		 SQL = "SELECT top 1 * FROM KS_AskAnswer WHERE TopicID="& TopicID &" And UserName='"& KS.C("AdminName") & "'"
		 Rs.Open SQL,Conn,1,3
		 If Rs.BOF And Rs.EOF Then
		            Rs.Addnew
					Rs("TopicID") = TopicID
					Rs("classid") = ClassID
					Rs("classname") = ClassName
					Rs("Username") = KS.C("AdminName")
					Rs("PostUsername") = PostUsername
					Rs("title") = AskTopic
					Rs("AnswerTime") = Now()
					Rs("PostTable") = PostTable
					Rs("AnswerNum") = 1
					If bestanswer=1 Then
					Rs("AnswerMode") = 1
					Rs("TopicMode") = 1
					else
					Rs("AnswerMode") = 0
					Rs("TopicMode") = 0
					end if
					Rs.Update
		Else
					Rs("AnswerTime") = Now()
					Rs("AnswerNum") = Rs("AnswerNum") + 1
					Rs.Update
		End If
		
		Rs.Close:Set Rs = Nothing
		Set Rs = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT top 1 * FROM " & PostTable & " WHERE (postsid is null)"
		Rs.Open SQL,Conn,1,3
			Rs.Addnew
					Rs("classid") = classid
					Rs("TopicID") = TopicID
					Rs("UserName") = KS.C("AdminName")
					Rs("topic") = AskTopic
					Rs("content") = Content
					Rs("addText") = ""
					Rs("PostTime") = Now()
					Rs("DoneTime") = Now()
					Rs("length") = KS.strLength(Content)
					If bestanswer=1 Then
					Rs("star") = 3
					Rs("satis") = 1
					else
					Rs("star") = 0
					Rs("satis") = 0
					end if
					Rs("LockTopic") = 0
					Rs("PostsMode") = 1
					Rs("VoteNum") = 0
					Rs("Plus") = 0
					Rs("Minus") = 0
					Rs("PostIP") = KS.GetIP()
					Rs("Report") = 0
					Rs("IsZJ")=1
			Rs.Update
			Rs.MoveLast
			PostsId=rs("postsid")
		Rs.Close:Set Rs = Nothing
		If bestanswer=1 Then
			Conn.Execute ("UPDATE KS_AskTopic SET BestUserName='" & KS.C("UserName") &"',BestUserId=" & KS.ChkClng(KS.C("UserId"))&",LastPostTime="& SqlNowString &",TopicMode=1 WHERE topicid="&topicid)
			Conn.Execute ("UPDATE KS_AskClass SET AskPendNum=AskPendNum-1,AskDoneNum=AskDoneNum+1 WHERE classid="& classid)
		end if
	  Response.Write "<script>alert('恭喜，答案提交成功!');top.MainFrame.location.reload();top.closeWindow();</script>" 

	End Sub
End Class
%>
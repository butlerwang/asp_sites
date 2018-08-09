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
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
         If Not KS.ReturnPowerResult(0, "WDXT10002") Then          '检查是权限
					 Call KS.ReturnErr(1, "")
					 KS.Die ""
		 End If
%>
<html>
<head>
<link href="Include/Admin_Style.CSS" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script src="../KS_Inc/common.js" language="JavaScript"></script>
<script src="../KS_Inc/jQuery.js" language="JavaScript"></script>
</head>
<body>
<%
    Response.Write "<ul id='menu_top'>"
	Response.Write "<li onclick=""parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape('问答系统 >> <font color=red>添加问答分类</font>')+'&ButtonSymbol=Go';location.href='?action=add';"" class='parent'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/addfolder.gif' border='0' align='absmiddle'>添加分类</span></li>"
	Response.Write "<li onclick='location.href=""?action=orders""' class='parent' onclick='MoveClassInfo();'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/set.gif' border='0' align='absmiddle'>一级分类排序</span></li>"
	Response.Write "<li onclick='location.href=""?action=total""' class='parent' onclick='MoveClassInfo();'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/unite.gif' border='0' align='absmiddle'>更新分类统计</span></li>"
	Response.Write "<li class='parent' onclick=""location.href='?';"""
	if KS.G("Action")="" Then Response.Write " disabled"
	Response.Write"><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/move.gif' border='0' align='absmiddle'>回上一级</span></li>"
	Response.Write "</ul>"

	Dim Action:Action = LCase(Request("action"))
	Select Case Trim(Action)
	Case "savenew"
		Call savenew()
	Case "savedit"
		Call savedit()
	Case "add"
		Call addCategory()
	Case "edit"
		Call editCategory()
	Case "del"
		Call delCategory()
	Case "orders"
		Call ClassOrders()
	Case "updatorders"
		Call UpdateOrders()
	Case "restore"
		Call Restoration()
	Case "total"
	    Call ClassTotal()
	Case Else
		Call showmain()
	End Select
End Sub

Sub showmain()
	Dim Rs,SQL,i
	Dim tdstyle
	Response.Write " <table border='0' cellpadding='0' cellspacing='0'  width='100%' align='center'>"
	Response.Write " <tr class='sort'>"
	Response.Write " <td width=""35%"">问答分类名称 </td>"
	Response.Write " <td width=""43%"">管理选项</td>"
	Response.Write "</tr>" & vbNewLine
	SQL = "SELECT * FROM KS_AskClass ORDER BY rootid,orders"
	Set Rs=Server.CreateObject("ADODB.Recordset")
	Rs.Open SQL, Conn, 1, 1
	If Rs.BOF And Rs.EOF Then
		Response.Write " <tr> <td align=""center"" colspan=""2"" class=""tablerow1"">您还没有添加任何分类！</td></tr>"
	End If
	i = 0
	Do While Not Rs.EOF
		Response.Write " <tr>"
		Response.Write " <td class='splittd'>"
		Response.Write " "
		If Rs("depth") = 1 Then Response.Write "&nbsp;&nbsp;<font color=""#666666"">├</font>"
		If Rs("depth") > 1 Then
			For i = 2 To Rs("depth")
				Response.Write "&nbsp;&nbsp;<font color=""#666666"">│</font>"
			Next
			Response.Write "&nbsp;&nbsp;<font color=""#666666"">├</font> "
		End If
		If Rs("parentid") = 0 Then Response.Write ("<img src='Images/Folder/domain.gif' align='absmiddle'/><b>")
		Response.Write Rs("ClassName")
		If Rs("child") > 0 Then Response.Write "(" & Rs("child") & ")"
		Response.Write " </td>" & vbNewLine
		Response.Write " <td class='splittd' align=""center"">"
		Response.Write "<a onclick=""parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape('问答系统 >> <font color=red>添加问答分类</font>')+'&ButtonSymbol=Go';"" href=""?action=add&editid="
		Response.Write Rs("classid")
		Response.Write """>添加分类</a>"
		Response.Write " | <a onclick=""parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape('问答系统 >> <font color=red>问答分类设置</font>')+'&ButtonSymbol=GoSave';"" href=""?action=edit&editid="
		Response.Write Rs("classid")
		Response.Write """>分类设置</a>"
		Response.Write " |"
		Response.Write " "
		If Rs("child") < 1 Then
			Response.Write " <a href=""?action=del&ChannelID="&ChannelID&"&editid="
			Response.Write Rs("classid")
			Response.Write """ onclick=""{if(confirm('删除将包括该分类的所有信息，确定删除吗?')){return true;}return false;}"">删除分类</a>"
		Else
			Response.Write " <a href=""#"" onclick=""{if(confirm('该分类含有下属分类，必须先删除其下属分类方能删除本分类！')){return true;}return false;}"">"
			Response.Write " 删除分类</a>"
		End If
		Response.Write " </td>" & vbNewLine
		Response.Write "</tr>" & vbNewLine
		Rs.movenext
		i = i + 1
	Loop
	Rs.Close
	Set Rs = Nothing
	Response.Write "</table>"
End Sub

Sub addCategory()
	Dim NewClassID
	Dim Rs,SQL,i
	SQL = "SELECT MAX(ClassID) FROM KS_AskClass"
	Set Rs = Conn.Execute(SQL)
	If Rs.BOF And Rs.EOF Then
		NewClassID = 1
	Else
		NewClassID = Rs(0) + 1
	End If
	If IsNull(NewClassID) Then NewClassID = 1
	Rs.Close
%>
<script language="javascript">
function CheckForm(){ 
 if ($('#ClassName').val()=='')
 {
   alert('请输入分类名称!');
   $('#ClassName').focus();
   return false;
 }
 $("#myform").submit();
}
</script>
<div style="text-align:center;height:30px;line-height:30px;font-weight:bold">添加问答分类</div>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="ctable">
	<form name="myform" id="myform" method="POST" action="?action=savenew">
	<input type="hidden" name="NewClassID" value="<%=NewClassID%>">
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
		<td class="clefttitle" align="right"><strong>所属分类：</strong></td>
		<td>
<%
	Response.Write " <select name=""class"">"
	Response.Write "<option value=""0"">做为一级分类</option>"
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
		</td>
	</tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
		<td width="20%" class="clefttitle" align="right"><strong>分类名称：</strong><br/>
		<font color="red">添加多个分类请用回车分开</font></td>
		<td width="80%">
		<textarea name="ClassName" id="ClassName" cols="50" rows="5"></textarea>
		</td>
	</tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
		<td class="clefttitle" align="right"><strong>分类说明：</strong></td>
		<td>
		<textarea name="Readme" cols="50" rows="5"></textarea></td>
	</tr>

	</form>
</table>
<%
End Sub

Sub editCategory()
	Dim RsObj
	Dim Rs,SQL,i
	Set Rs = Conn.Execute("SELECT * FROM KS_AskClass WHERE classid = " & KS.ChkClng(Request("editid")))
	If Rs.BOF And Rs.EOF Then
		FoundErr = True
		ErrMsg = "数据库出现错误,没有此站点分类!"
		Rs.Close
		Set Rs = Nothing
		Exit Sub
	End If
%>
<script language="javascript">
function CheckForm(){ 
 if ($('#ClassName').val()=='')
 {
   alert('请输入分类名称!');
   $('#ClassName').focus();
   return false;
 }
 $("#myform").submit();
}
</script>
<div style="text-align:center;height:30px;line-height:30px;font-weight:bold">编辑问答分类</div>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="ctable">
	<form name="myform" id="myform" method="POST" action="?action=savedit">
	<input type="hidden" name="editid" value="<%=Request("editid")%>">
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
		<td class="clefttitle" align="right"><strong>所属分类：</strong></td>
		<td>
<%
	Response.Write " <select name=""class"">"
	Response.Write "<option value=""0"">做为一级分类</option>"
	SQL = "SELECT classid,depth,ClassName FROM KS_AskClass ORDER BY rootid,orders"
	Set RsObj = Conn.Execute(SQL)
	Do While Not RsObj.EOF
		Response.Write "<option value=""" & RsObj("classid") & """ "
		If CLng(Rs("parentid")) = RsObj("classid") Then Response.Write "selected"
		Response.Write ">"
		If RsObj("depth") = 1 Then Response.Write "&nbsp;&nbsp;├ "
		If RsObj("depth") > 1 Then
			For i = 2 To RsObj("depth")
				Response.Write "&nbsp;&nbsp;│"
			Next
			Response.Write "&nbsp;&nbsp;├ "
		End If
		Response.Write RsObj("ClassName") & "</option>" & vbCrLf
		RsObj.movenext
	Loop
	RsObj.Close
	Response.Write "</select>"
	Set RsObj = Nothing
%>
		</td>
	</tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
		<td width="20%" class="clefttitle" align="right"><strong>分类名称：</strong></td>
		<td width="80%">
		<input type="text" name="ClassName" id="ClassName" size="35" value="<% = Rs("ClassName")%>">
		</td>
	</tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
		<td class="clefttitle" align="right"><strong>分类说明：</strong></td>
		<td >
		<textarea name="Readme" cols="50" rows="5"><%=Server.HTMLEncode(Rs("readme")&"")%></textarea></td>
	</tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
		<td class="clefttitle" align="right"><strong>分类统计：</strong></td>
		<td>
		未解决：<input type="text" name="AskPendNum" size="10" value="<%=Rs("AskPendNum")%>">
		已解决：<input type="text" name="AskDoneNum" size="10" value="<%=Rs("AskDoneNum")%>">
		<span style="display:none">
		投票：<input type="text" name="AskVoteNum" size="10" value="<%=Rs("AskVoteNum")%>">
		分享：<input type="text" name="AskshareNum" size="10" value="<%=Rs("AskshareNum")%>">
		</span>
		</td>
	</tr>
	</form>
</table>
<%
Set Rs = Nothing
End Sub

Sub savenew()
	If Trim(Request.Form("classname")) = "" Then
		Call KS.AlertHistory("请输入分类名称!",-1)
		Exit Sub
	End If
	If Not IsNumeric(Request.Form("class")) Then
		Call KS.AlertHistory("请选择所属分类!",-1)
		Exit Sub
	End If
	'If Trim(Request.Form("Readme")) = "" Then
	'	Call KS.AlertHistory("请输入分类说明!",-1)
	'	Exit Sub
	'End If
	Dim Rs,SQL,i
	Dim newclassid,rootid,ParentID,depth,orders
	Dim maxrootid,Parentstr,neworders
	Dim m_strClassname,m_arrClassname,strClassname

	m_strClassname = Replace(Trim(Request("classname")), vbCrLf, "$$$")
	m_arrClassname = Split(m_strClassname, "$$$")

	If Request("class") <> "0" Then
		SQL = "SELECT rootid,classid,depth,orders,Parentstr FROM KS_AskClass WHERE classid=" & KS.ChkClng(Request("class"))
		Set Rs = Conn.Execute (SQL)
		rootid = Rs(0)
		ParentID = Rs(1)
		depth = Rs(2)
		orders = Rs(3)
		If depth > 3 Then
			Call KS.AlertHistory("本系统限制3级分类",-1)
			Exit Sub
		End If
		Parentstr = Rs(4)
		Set Rs = Nothing
	Else
		SQL = "SELECT MAX(rootid) FROM KS_AskClass"
		Set Rs = Conn.Execute (SQL)
		maxrootid = KS.ChkClng(Rs(0)) + 1
		If maxrootid =0 Then maxrootid = 1
		Set Rs = Nothing
	End If

	SQL = "SELECT classid FROM KS_AskClass WHERE classid=" & KS.ChkClng(Request("newclassid"))
	Set Rs = Conn.Execute (SQL)
	If Not (Rs.EOF And Rs.BOF) Then
		Call KS.AlertHistory("您不能指定和别的分类一样的序号!",-1)
		Exit Sub
	Else
		newclassid = KS.ChkClng(Request("newclassid"))
	End If
	Set Rs = Nothing
	
	Set Rs=Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM KS_AskClass"
	Rs.Open SQL, Conn, 1, 3
	For i = 0 To UBound(m_arrClassname)
		strClassname = KS.R(Trim(m_arrClassname(i)))
		If strClassname <> "" Then
			Rs.addnew
			If Request("class") <> "0" Then
				Rs("depth") = depth + 1
				Rs("rootid") = rootid
				Rs("parentid") = Request.Form("class")
				'If Parentstr = "0" Then
				'	Rs("Parentstr") = Request.Form("class")
				'Else
				'	Rs("Parentstr") = parentstr & "," & KS.ChkClng(Request.Form("class"))
				'End If
			Else
				Rs("depth") = 0
				Rs("rootid") = maxrootid
				Rs("parentid") = 0
				Rs("ParentStr") = 0
			End If
            Rs("parentstr")=parentstr & newclassid & ","
			Rs("child") = 0
			Rs("classid") = newclassid
			Rs("orders") = newclassid
			Rs("classname") = strClassname
			Rs("readme") = Trim(Request.Form("readme"))
			Rs("Askmaster") = ""
			Rs("c_setting") = ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,"
			Rs("AskPendNum") = 0
			Rs("AskDoneNum") = 0
			Rs("AskVoteNum") = 0
			Rs("AskshareNum") = 0
			Rs.Update
			Rs.MoveNext
			newclassid = newclassid + 1
			maxrootid = maxrootid + 1
		End If
	Next
	Rs.Close
	Set Rs = Nothing

	CheckAndFixClass 0,1
	Call KS.Confirm("恭喜您！添加新的分类成功,继续添加吗?","?action=add","?")
End Sub

Sub savedit()
	If CLng(Request.Form("editid")) = CLng(Request.Form("class")) Then
		Call KS.AlertHistory("所属分类不能指定自己",-1)
		Exit Sub
	End If
	If Trim(Request.Form("classname")) = "" Then
		Call KS.AlertHistory("请输入分类名称!",-1)
		Exit Sub
	End If
	
	Dim newclassid,maxrootid,readme
	Dim parentid,depth,child,ParentStr,rootid,iparentid,iParentStr
	Dim trs,mrs
	Dim Rs,SQL,nParentStr
	Set Rs=Server.CreateObject("ADODB.Recordset")
	Rs.Open "SELECT ParentStr FROM KS_AskClass Where ClassID=" & KS.ChkClng(KS.G("Class")),conn,1,1
	If Not RS.Eof Then
	 nParentStr=Rs(0)
	End If
	Rs.Close
	SQL = "SELECT * FROM KS_AskClass WHERE classid="& KS.ChkClng(Request("editid"))
	Rs.Open SQL,Conn,1,3
	newclassid = Rs("classid")
	parentid = Rs("parentid")
	iparentid = Rs("parentid")
	ParentStr = Rs("ParentStr")
	depth = Rs("depth")
	child = Rs("child")
	rootid = Rs("rootid")
	
	'判断所指定的分类是否其下属分类
	If ParentID=0 Then
		If CLng(Request("class"))<>0 Then
		Set trs=Conn.Execute("SELECT rootid FROM KS_AskClass WHERE classid="&KS.ChkClng(Request("class")))
		If rootid=trs(0) Then
			Call KS.AlertHistory("您不能指定该问答的下属分类作为所属分类",-1)
			Exit Sub
		End If

		End If
	Else
		Set trs=Conn.Execute("SELECT classid FROM KS_AskClass WHERE ParentStr like '%"&ParentStr&","&newclassid&"%' And classid="&KS.ChkClng(Request("class")))
		If Not (trs.EOF And trs.BOF) Then
			Call KS.AlertHistory("您不能指定该问答的下属分类作为所属分类",-1)
			Exit Sub
		End If
	End If
	If parentid = 0 Then
		parentid = Rs("classid")
		iparentid=0
	End If
	Rs("parentstr")=nParentStr & rs("classid") & ","
	Rs("classname") = Trim(Request.Form("classname"))
	Rs("parentid") = KS.ChkClng(Request.Form("class"))
	Rs("readme") =Trim( Request("readme"))
	Rs("AskPendNum") = KS.ChkClng(Request.Form("AskPendNum"))
	Rs("AskDoneNum") = KS.ChkClng(Request.Form("AskDoneNum"))
	Rs("AskVoteNum") = KS.ChkClng(Request.Form("AskVoteNum"))
	Rs("AskshareNum") = KS.ChkClng(Request.Form("AskshareNum"))
	Rs.Update 
	Rs.Close
	Set Rs=nothing
	
	Set mrs=Conn.Execute("SELECT MAX(rootid) FROM KS_AskClass")
	Maxrootid=mrs(0)+1
	mrs.close:Set mrs=nothing
	CheckAndFixClass 0,1
	Call KS.Alert("恭喜您！分类修改成功!","?")
End Sub

Sub delCategory()
	Dim Rs,SQL,i
	Dim ChildStr,nChildStr
	Dim Rss,Rsc
	On Error Resume Next
	Set Rs = Conn.Execute("SELECT ParentStr,child,depth,parentid FROM KS_AskClass WHERE classid=" & KS.ChkClng(Request("editid")))
	If Not (Rs.EOF And Rs.BOF) Then
		If Rs(1) > 0 Then
			Call KS.AlertHistory("该分类含有下属分类，请删除其下属分类后再进行删除本分类的操作!",-1)
			Exit Sub
		End If

		If Rs(2) > 0 Then
			Conn.Execute ("UPDATE KS_AskClass Set child=child-1 WHERE classid in (" & Rs(0) & ")")
		End If
		For i = 0 To Ubound(AllPostTable)
			SQL = "DELETE FROM " & AllPostTable(i) & " WHERE classid=" & KS.ChkClng(Request("editid"))
			Conn.Execute(SQL)
		Next
		Conn.Execute("DELETE FROM KS_AskAnswer WHERE classid=" & KS.ChkClng(Request("editid")))
		Conn.Execute("DELETE FROM KS_AskTopic WHERE classid=" & KS.ChkClng(Request("editid")))
		Conn.Execute("DELETE FROM KS_AskClass WHERE classid=" & KS.ChkClng(Request("editid")))
		
	End If
	Set Rs = Nothing
	Conn.Execute("UPDATE KS_AskClass Set child=0 WHERE child<0")
	CheckAndFixClass 0,1
	UpdateClassTotal
	Call KS.Alert("恭喜您！分类删除成功。","?")
End Sub

Sub Restoration()
	CheckAndFixClass 0,1
	Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub

Sub CheckAndFixClass(ParentID,orders)
	Dim Rs,Child,ParentStr
	If ParentID=0 Then
		Conn.Execute("UPDATE KS_AskClass Set Depth=0 WHERE ParentID=0")
	End If
	Set Rs=Conn.Execute("SELECT classid,rootid,ParentStr,Depth FROM KS_AskClass WHERE ParentID="&ParentID&" ORDER BY rootid,orders")
	Do while Not Rs.EOF
		Conn.Execute "UPDATE KS_AskClass Set Depth="&Rs(3)+1&",rootid="&Rs(1)&" WHERE ParentID="&Rs(0)&"",Child
		Conn.Execute("UPDATE KS_AskClass Set Child="&Child&",orders="&orders&" WHERE classid="&Rs(0)&"")
		orders=orders+1
		CheckAndFixClass Rs(0),orders
		Rs.MoveNext
	Loop
	Set Rs=Nothing
	Application(KS.SiteSN&"_askclasslist")=empty
End Sub


Sub ClassTotal()
 UpdateClassTotal()
 Call KS.AlertHistory("恭喜,分类问题数统计成功!",-1)
End Sub

Sub UpdateClassTotal()
 Dim Rs:Set Rs=Server.CreateObject("ADODB.RECORDSET")
 Rs.Open "Select * From KS_AskClass Order By Rootid,orders",conn,1,3
 do while not rs.Eof 
   Rs("AskPendNum")=Conn.Execute("select count(topicid) From KS_AskTopic WHERE classid in (SELECT classid FROM KS_AskClass WHERE ','+parentstr+'' like '%,"&rs("classid")&",%') and topicmode=0")(0)
   Rs("AskDoneNum")=Conn.Execute("select count(topicid) From KS_AskTopic WHERE classid in (SELECT classid FROM KS_AskClass WHERE ','+parentstr+'' like '%,"&rs("classid")&",%') and topicmode<>0")(0)
   Rs.Update
  Rs.MoveNext
 Loop
 Rs.Close
 Set RS=Nothing
 Application(KS.SiteSN&"_askclasslist")=empty
End Sub

Sub ClassOrders()
%>
<br>
<table border="0" cellspacing="1" cellpadding="3" align="center"  class="Ctable">
	<tr> 
	<th class="sort" colspan=2>问答一级分类重新排序修改(请在相应分类的排序表单内输入相应的排列序号)</th>
	</tr>
	<tr>
<%
	Dim Rs,SQL,i
	Set Rs=Server.CreateObject("ADODB.Recordset")
	SQL="SELECT * FROM KS_AskClass WHERE ParentID=0 ORDER BY rootid"
	Rs.Open SQL,Conn,1,1
	If Rs.Eof And Rs.Bof Then
		Response.Write "还没有相应的问答分类。"
	Else
		Do While Not Rs.Eof
		Response.Write "<form action=""?action=updatorders"" method=""post""><tr class='tdbg'>"
		Response.Write "<td align=""right"" class=""clefttitle"">" & rs("ClassName") & "</td><td><input type=""text"" name=""OrderID"" size=""4"" value="""&rs("rootid")&"""><input type=""hidden"" name=""cID"" value="""&rs("rootid")&""">&nbsp;&nbsp;<input type=""submit"" name=""Submit"" value=""修改"" class=""button""></td></tr></form>"
		Rs.Movenext
		Loop
%>
</table>
<%
	End If
	Rs.Close
	Set Rs=Nothing
%>
	</td>
	</tr>
</table>
<%
End Sub

Sub UpdateOrders()
	Dim cID,OrderID,Rs
	cID = Replace(Request.Form("cID"),"'","")
	OrderID = Replace(Request.Form("OrderID"),"'","")
	Set Rs = Conn.Execute("SELECT classid FROM KS_AskClass WHERE rootid="&orderid)
	If Rs.EOF And Rs.BOF Then
		Conn.Execute("UPDATE KS_AskClass SET rootid="&OrderID&" WHERE rootid="&cID)
		Call KS.AlertHintScript("设置成功!")
	Else
		Call KS.AlertHistory("请不要和其他分类设置相同的序号",-1)
		Response.End
	End If
End Sub
End Class
%>
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim Action, Flag, i, RsObj,selfont,ChannelID,SQL,RS
Dim KS:Set KS= New PublicCls
ChannelID = KS.ChkClng(KS.G("ChannelID"))
If ChannelID = 0 Then ChannelID = 3
If Not KS.ReturnPowerResult(0, "KMST20002") Then Call KS.ReturnErr(1, "")   '权限检查
With KS
.echo "<html>"
.echo "<head>"
.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
.echo "<title>管理主页面</title>"
.echo "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
.echo "<script src=""../KS_Inc/Jquery.js"" language=""JavaScript""></script>"
.echo "<body>"
.echo "<ul id='menu_top'>"

.echo "<li class='parent' onclick='addserver()'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>添加服务器</span></li>"
.echo "<li class='parent' onclick=""location.href='KS.DownServer.asp?action=serverorders&ChannelID=" & ChannelID & "';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/reb.gif' border='0' align='absmiddle'>路径排序</span></li>"
.echo "<li class='parent' onclick=""location.href='KS.DownServer.asp?ChannelID=" & ChannelID & "';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/back.gif' border='0' align='absmiddle'>返回首页</span></li>"
			.echo "<li></li><div><strong>按模型设置:</strong><select id='channelid' name='channelid' onchange=""if (this.value!=0){location.href='?channelid='+this.value;}"">"
			.echo " <option value='0'>---请选择模型---</option>"
	
			If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
			Dim ModelXML,Node
			Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
			For Each Node In ModelXML.documentElement.SelectNodes("channel[@ks21=1][@ks6=3]")
			  If trim(ChannelID)=trim(Node.SelectSinglenode("@ks0").text) Then
			    .echo "<option value='" &Node.SelectSingleNode("@ks0").text &"' selected>" & Node.SelectSingleNode("@ks1").text & "</option>"

			  Else
			   .echo "<option value='" &Node.SelectSingleNode("@ks0").text &"'>" & Node.SelectSingleNode("@ks1").text & "</option>"
			  End If
			next
			.echo "</select></div>"

.echo "</ul>"
End With
%>
<script language="javascript">

function CheckForm()
{
 if ($('input[name=DownloadName]').val()=='')
 {
  alert('请输入下载服务器名称!');
  $('input[name=DownloadName]').focus();
  return false;
 }
 if ($('#DownloadPath').val()==''&&$('#servers').val()!=0)
 {
  alert('请输入服务器路径!');
  $("#DownloadPath").focus();
  return false;
  }
 $('form[name=myform]').submit();
}
function addserver(editid)
{
 location.href='?action=add&editid='+editid+'&channelid=<%=channelid%>';
 $(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ChannelID=<%=ChannelID%>&OpStr=下载服务器管理 >> <font color=red>添加下载服务器</font>&ButtonSymbol=GO';
}
function editserver(editid)
{
 location.href='?action=edit&editid='+editid+'&channelid=<%=channelid%>';
 $(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ChannelID=<%=ChannelID%>&OpStr=下载服务器管理 >> <font color=red>编辑下载服务器</font>&ButtonSymbol=GOSave';
}
function formatbt()
{
  var arr = showModalDialog("include/btformat.htm?",null, "dialogWidth:250pt;dialogHeight:166pt;toolbar=no;location=no;directories=no;status=no;menubar=NO;scrollbars=no;resizable=no;help=0; status:0");
  if (arr != null){
     $('#selfont').val(arr);
     myfont.innerHTML="<span style='background-color: #FFFFff;font-size:14px' "+arr+">设置标题样式 ABCdef</span>";
  }
}
function Cancelform()
{
  $('#selfont').val('');
  myfont.innerHTML="<span style='background-color: #FFFFff;font-size:14px;color:#000000'>设置标题样式 ABCdef</span>";
}
function setunion(val)
{
 val=parseInt(val);
 if (val==0)
 {
   $('#unionarea').hide();
 }
 else
 {
 $('#unionarea').show();}
}
function setdisabled(val)
{
  val=parseInt(val);
  if (val==0)
  {
   $('#s1').hide();
   $('#s2').hide();
   $('#s3').hide();
   $('#s4').hide();
  }
  else
  {
   $('#s1').show();
   $('#s2').show();
   $('#s3').show();
   $('#s4').show();
  }
}
//-->
</script>
<%
Action = LCase(KS.G("action"))

Select Case Request("action")
	Case "add"
		Call sAdd
	Case "edit"
		Call sEdit
	Case "savenew"
		Call savenew
	Case "savedit"
		Call saveedit
	Case "del"
		Call DelDownPath
	Case "serverorders"
		Call serverorders
	Case "updateorders"
		Call updateorders
	Case "lock"
		Call isLock
	Case "free"
		Call FreeLock
	Case Else
		Call ShowMain
End Select
'================================================
'过程名：ShowMain
'作  用：服务器管理首页
'================================================
Sub ShowMain()
	Dim DownloadName
	Response.Write "<script>$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ChannelID=" & ChannelID & "&OpStr=" & KS.C_S(ChannelID,1)&" >> <font color=red>下载服务器管理</font>&ButtonSymbol=Disabled';</script>"

	Response.Write " <table width=""100%"" cellspacing=""0"" cellpadding=""0"" align=center>" & vbcrlf
	Response.Write " <tr class=""sort"">" & vbcrlf
	Response.Write " <td width=""35%"">服务器分类</td>" & vbcrlf
	Response.Write " <td width=""45%"">操 作</td>" & vbcrlf
	Response.Write " <td width=""10%"" noWrap>日下载数</td>" & vbcrlf
	Response.Write " <td width=""10%"" noWrap>总共下载数</td>" & vbcrlf
	Response.Write "</tr>" & vbcrlf
	SQL = "SELECT * FROM KS_DownSer WHERE ChannelID=" & ChannelID & " ORDER BY rootid,orders"
	Set Rs = CreateObject("ADODB.Recordset")
	Rs.Open SQL, Conn, 1, 1
	Do While Not Rs.EOF
		selfont = Rs("selfont") & ""
		Response.Write " <tr height='23' class='list' onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">" & vbcrlf
		Response.Write " <td width='35%'>" & vbcrlf
		If Rs("depth") = 1 Then Response.Write "&nbsp;&nbsp;<font color=""#666666"">├</font>"
		If Rs("depth") > 1 Then
			For i = 2 To Rs("depth")
				Response.Write "&nbsp;&nbsp;<font color=""#666666"">│</font>"
			Next
			Response.Write "&nbsp;&nbsp;<font color=""#666666"">├</font> "
		End If
		If Rs("parentid") = 0 Then Response.Write ("<b>[" & Rs("rootid") & "] ")
		If Len(selfont) < 10 Then
			DownloadName = Rs("DownloadName")
		Else
			DownloadName = "<span " & selfont & ">" & Rs("DownloadName") & "</span>"
		End If
		Response.Write DownloadName
		If Rs("isLock") = 1 Then
			Response.Write " <img src='images/locks.gif' border=0 align=absMiddle>"
		End If
		If Rs("child") > 0 Then Response.Write "(" & Rs("child") & ")"
		If Rs("parentid") = 0 Then Response.Write ("</b>")
		Response.Write " </td>" & vbcrlf
		Response.Write " <td align=right>"
		If Rs("depth") = 0 Then
			Response.Write "<a href=""javascript:addserver(" &Rs("downid") &");"">添加下载服务器路径</a>" & vbcrlf
		Else
		    If Rs("isLock") <> 1 Then
			Response.Write "<a href=""KS.DownServer.asp?action=lock&editid="&Rs("downid")&"&ChannelID=" & ChannelID & """>锁定</a>"
			Else
			Response.Write "<a href=""KS.DownServer.asp?action=free&editid="&Rs("downid")&"&ChannelID=" & ChannelID & """>解除</a>"
			End If
		End If
		Response.Write " | <a href=""javascript:editserver(" & Rs("downid") & ");"">服务器设置</a>" & vbcrlf
		Response.Write " |" & vbcrlf
		Response.Write " "
		If Rs("child") = 0 Then
			Response.Write " <a href=""KS.DownServer.asp?action=del&editid="
			Response.Write Rs("downid")
			Response.Write "&amp;ChannelID=" & ChannelID & """ onclick=""{if(confirm('删除将包括该服务器的所有信息，确定删除吗?')){return true;}return false;}"">删除" & vbcrlf
			Response.Write " "
		Else
			Response.Write "<a href=""#"" onclick=""{if(confirm('该服务器含有下载路径，必须先删除其下载路径方能删除本服务器！')){return true;}return false;}"">" & vbcrlf
			Response.Write " 删除</a>" & vbcrlf
			Response.Write " "
		End If
		Response.Write " </td>" & vbcrlf
		Response.Write " <td align=""center"">"
		If Rs("depth") > 0 Then
			Response.Write Rs("DayDownHits")
		End If
		Response.Write " </td>" & vbcrlf
		Response.Write " <td align=""center"">"
		If Rs("depth") > 0 Then
			Response.Write Rs("AllDownHits")
		End If
		Response.Write " </td>" & vbcrlf
		Response.Write "</tr>" & vbcrlf
		Response.Write ("<tr><td colspan=6 background='images/line.gif'></td></tr>")
		Rs.MoveNext
		i = i + 1
	Loop
	Rs.Close
	Set Rs = Nothing
	Response.Write "</table>" & vbcrlf
End Sub
'================================================
'过程名：sAdd
'作  用：添加服务器
'================================================
Sub sAdd()
	Dim ServerNum
	On Error Resume Next
	Set Rs = CreateObject("ADODB.Recordset")
	SQL = "SELECT MAX(downid) FROM KS_DownSer"
	Rs.Open SQL, Conn, 1, 1
	If Rs.BOF And Rs.EOF Then
		ServerNum = 1
	Else
		ServerNum = Rs(0) + 1
	End If
	If IsNull(ServerNum) Then ServerNum = 1
	Rs.Close
	Response.Write " <table width=""100%"" border=""0"" cellspacing=""1"" cellpadding=""1"" align=center>" & vbcrlf
	Response.Write "<form name=""myform"" action =""KS.DownServer.asp?action=savenew"" method=""post"">" & vbcrlf
	Response.Write "<input type=""hidden"" name=""newdownid"" value="""&ServerNum&""">" & vbcrlf
	Response.Write "<input type=""hidden"" name=ChannelID value="""& ChannelID&""">" & vbcrlf
	Response.Write " <tr class='sort'>" & vbcrlf
	Response.Write " <td colspan=2>添加新的服务器</td>" & vbcrlf
	Response.Write "</tr>" & vbcrlf
	Response.Write " <tr class='tdbg'>" & vbcrlf
	Response.Write " <td class=""clefttitle"" align=""right""><b>所属类别：</b></td>" & vbcrlf
	Response.Write " <td>" & vbcrlf
	Response.Write " <select name=""servers"" id=""servers"" onchange=""setdisabled(this.value);"">" & vbcrlf
	Response.Write " <option value=""0"">做为服务器分类</option>" & vbcrlf
	
	SQL = "SELECT * FROM KS_DownSer WHERE ChannelID=" & ChannelID & " And depth = 0 ORDER BY rootid"
	Rs.Open SQL, Conn, 1, 1
	Do While Not Rs.EOF
		Response.Write "<option value=""" & Rs("downid") & """ "
		If Len(Request("editid")) <> 0 And CLng(Request("editid")) = Rs("downid") Then Response.Write "selected"
		Response.Write ">"
		Response.Write Rs("DownloadName") & "</option>" & vbCrLf
		Rs.MoveNext
	Loop
	Rs.Close
	
	Response.Write "</select>"
	Response.Write "</td></tr>" & vbcrlf
	Response.Write " <tr class='tdbg'>" & vbcrlf
	Response.Write " <td width=""30%"" class=""clefttitle"" align=""right""><b>服务器名称：</b></td>" & vbcrlf
	Response.Write " <td width=""70%"">"
	Response.Write " <input type=""text"" name=""DownloadName"" size=""60"">" & vbcrlf
	Response.Write "</td>" & vbcrlf
	Response.Write "</tr>" & vbcrlf
	Response.Write " <tr class='tdbg' id='s1'>" & vbcrlf
	Response.Write " <td  class=""clefttitle"" align=""right""><b>服务器名称样式：</b></td>" & vbcrlf
	Response.Write " <td>样式:<input type=""hidden"" name=""selfont"" id=""selfont"" size=""1"" value="""">&nbsp;"
	Response.Write " <span style=""background-color: #fFfFff"" id=""myfont"" onclick=""javascript:formatbt(this);""  style='cursor:pointer; font-size:14px' >设置标题样式 ABCdef</span> " & vbcrlf
	Response.Write "<input type=""checkbox"" name=""cancel"" onclick=""Cancelform()""> 取消格式"
	Response.Write "</td>" & vbcrlf
	Response.Write "</tr>" & vbcrlf
	'-------
	Response.Write " <tr class='tdbg' id='s2'>" & vbcrlf
	Response.Write " <td  class=""clefttitle"" align=""right""><b>服务器路径：</b></td>" & vbcrlf
	Response.Write " <td>" & vbcrlf
	Response.Write " <input type=""text"" name=""DownloadPath"" id=""DownloadPath"" size=""60"">" & vbcrlf
	Response.Write "</td>" & vbcrlf
	Response.Write "</tr>" & vbcrlf

	Response.Write " <tr class='tdbg' id='s3'>" & vbcrlf
	Response.Write " <td  class=""clefttitle"" align=""right""><b>是否直接显示下载地址：</b></td>" & vbcrlf
	Response.Write " <td>"
	Response.Write " <input type=radio name=isDisp value=""0"" checked> 否&nbsp;&nbsp;"
	Response.Write " <input type=radio name=isDisp value=""1""> 是"
	Response.Write " </td>" & vbcrlf
	Response.Write " </tr>" & vbcrlf
	Response.Write " <tr class='tdbg' id='s4'>" & vbcrlf
	Response.Write " <td height='25' class=""clefttitle"" align=""right""><b>是否外部连接：</b></td>" & vbcrlf
	Response.Write " <td>"
	Response.Write " <input type=radio onclick='setunion(this.value)' name=IsOuter value=""0"" checked> 否&nbsp;&nbsp;"
	Response.Write " <input type=radio onclick='setunion(this.value)' name=IsOuter value=""2""> WEB迅雷专用下载地址&nbsp;&nbsp;"
	Response.Write " <input type=radio onclick='setunion(this.value)' name=IsOuter value=""3""> FLASHGET(快车)专用下载地址"
	Response.Write "<div id='unionarea' style='display:none'>"
	Response.Write "联盟ID:<input type='text' name='unionid' size=12>"
	Response.Write "<font color=""red"">如果还没有联盟ID”，"
	Response.Write "请先注册<a href=""http://union.xunlei.com/"" target=""_blank""><font color=""blue"">迅雷联盟</font></a>|<a href=""http://union.flashget.com/"" target=""_blank""><font color=""blue"">快车联盟</font></a>"
	Response.Write "</div>"

	Response.Write "</td>" & vbcrlf
	Response.Write " </tr>" & vbcrlf
	Response.Write "</table>" & vbcrlf
	Response.Write "</form>" & vbcrlf
	Set Rs = Nothing
	Response.Write "<script>setdisabled($('#servers').val());</script>"
End Sub
'================================================
'过程名：sEdit
'作  用：编辑服务器
'================================================
Sub sEdit()
	Dim Rs_e
	On Error Resume Next
	Set Rs = CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM KS_DownSer WHERE downid=" & CLng(Request("editid"))
	Set Rs_e = Conn.Execute(SQL)
	Response.Write " <table width=""100%"" border=""0"" cellspacing=""1"" cellpadding=""1"" align=centers>" & vbcrlf
	Response.Write "<form name=""myform"" action =""KS.DownServer.asp?action=savedit"" method=""post"">" & vbcrlf
	Response.Write "<input type=""hidden"" name=editid value="""& Request("editid") &""">" & vbcrlf
	Response.Write "<input type=""hidden"" name=ChannelID value=""" & ChannelID & """>" & vbcrlf
	Response.Write " <tr class='sort'>" & vbcrlf
	Response.Write " <td colspan=2>编辑服务器："
	Response.Write Rs_e("DownloadName")
	Response.Write "</td>" & vbcrlf
	Response.Write " </tr>" & vbcrlf
	Response.Write " <tr class='tdbg'>" & vbcrlf
	Response.Write " <td height=30 class=""clefttitle"" align=""right""><b>所属类别：</b></td>" & vbcrlf
	Response.Write " <td>" & vbcrlf
	Response.Write " <select name=""servers"" id=""servers"" onchange='setdisabled(this.value)'>" & vbcrlf
	Response.Write " <option value=""0"">做为主服务器分类</option>" & vbcrlf
	Response.Write " "
	SQL = "SELECT * FROM KS_DownSer WHERE ChannelID=" & ChannelID & " ORDER BY rootid,orders"
	Set Rs = Conn.Execute(SQL)
	Do While Not Rs.EOF
		Response.Write "<option value=""" & Rs("downid") & """ "
		If Rs_e("parentid") = Rs("downid") Then Response.Write "selected"
		Response.Write ">"
		If Rs("depth") = 1 Then Response.Write "&nbsp;&nbsp;├ "
		If Rs("depth") > 1 Then
			For i = 2 To Rs("depth")
				Response.Write "&nbsp;&nbsp;│"
			Next
			Response.Write "&nbsp;&nbsp;├ "
		End If
		Response.Write Rs("DownloadName") & "</option>" & vbCrLf
		Rs.MoveNext
	Loop
	Rs.Close: Set Rs = Nothing
	Response.Write " </select> </td>" & vbcrlf
	Response.Write " </tr>" & vbcrlf
	Response.Write " <tr class=""tdbg"">" & vbcrlf
	Response.Write " <td width=""30%"" height=30 class=""clefttitle"" align=""right""><b>服务器名称：</b></td>" & vbcrlf
	Response.Write " <td width=""70%"">" & vbcrlf
	Response.Write " <input type=""text"" name=""DownloadName"" size=""60"" value="""
	Response.Write Rs_e("DownloadName")
	Response.Write """>" & vbcrlf
	Response.Write " </td>" & vbcrlf
	Response.Write " </tr>" & vbcrlf
	Response.Write " <tr class='tdbg' id='s1'>" & vbcrlf
	Response.Write " <td class=""clefttitle"" align=""right""><b>服务器名称样式：</b></td>" & vbcrlf
	Response.Write " <td>样式:<input type=""hidden"" name=""selfont"" id=""selfont"" size=""1"" value="""& Server.HTMLEncode(Rs_e("selfont") & "") &""">&nbsp;"
	Response.Write " <span style=""background-color: #fFfFff;"" id=""myfont"" onclick=""javascript:formatbt(this);""  style='cursor:pointer; font-size:14px'><span "& Rs_e("selfont") &">设置标题样式 ABCdef</span></span> " & vbcrlf
	Response.Write "<input type=""checkbox"" name=""cancel"" onclick=""Cancelform()""> 取消格式"
	Response.Write "</td>" & vbcrlf
	Response.Write "</tr>" & vbcrlf
	'-------
	Response.Write " <tr class='tdbg' id='s2'>" & vbcrlf
	Response.Write " <td class=""clefttitle"" align=""right""><b>服务器路径：</b><BR>" & vbcrlf
	Response.Write " 可以使用HTML代码</td>" & vbcrlf
	Response.Write " <td>" & vbcrlf
	Response.Write " <input type=""text"" name=""DownloadPath"" id=""DownloadPath"" size=""60"" value="""
	Response.Write Rs_e("DownloadPath")
	Response.Write """>" & vbcrlf
	Response.Write " </td>" & vbcrlf
	Response.Write " </tr>" & vbcrlf

	Response.Write " <tr class='tdbg' id='s3'>" & vbcrlf
	Response.Write " <td class=""clefttitle"" align=""right""><b>是否直接显示下载地址：</b></td>" & vbcrlf
	Response.Write " <td>"
	Response.Write " <input type=radio name=isDisp value=""0"""
	If Rs_e("IsDisp") = 0 Then Response.Write "  checked"
	Response.Write "> 否&nbsp;&nbsp;"
	Response.Write " <input type=radio name=isDisp value=""1"""
	If Rs_e("IsDisp") = 1 Then Response.Write "  checked"
	Response.Write "> 是"
	Response.Write " </td>" & vbcrlf
	Response.Write " </tr>" & vbcrlf
	Response.Write " <tr class='tdbg' id='s4'>" & vbcrlf
	Response.Write " <td heihgt='25' class=""clefttitle"" align=""right""><b>是否外部连接：</b></td>" & vbcrlf
	Response.Write " <td>"
	Response.Write " <input onclick='setunion(this.value)' type=radio name=IsOuter value=""0"""
	If Rs_e("IsOuter") = 0 Then Response.Write "  checked"
	Response.Write "> 否&nbsp;&nbsp;"
	Response.Write " <input onclick='setunion(this.value)' type=radio name=IsOuter value=""2"""
	If Rs_e("IsOuter") = 2 Then Response.Write "  checked"
	Response.Write "> WEB迅雷专用下载地址&nbsp;&nbsp;"
	Response.Write " <input onclick='setunion(this.value)' type=radio name=""IsOuter"" value=""3"""
	If Rs_e("IsOuter") = 3 Then Response.Write "  checked"
	Response.Write "> FLASHGET(快车)专用下载地址&nbsp;&nbsp;"
	Response.Write " <br>"
	If Rs_e("IsOuter") = 0 Then
	Response.Write "<div id='unionarea' style='display:none'>"
	Else
	Response.Write "<div id='unionarea'>"
	End If
	Response.Write "联盟ID:<input type='text' name='unionid' value='" & Rs_e("UnionID") &"' size=12>"
	Response.Write "<font color=""red"">如果还没有联盟ID”，"
	Response.Write "请先注册<a href=""http://union.xunlei.com/"" target=""_blank""><font color=""blue"">迅雷联盟</font></a>|<a href=""http://union.flashget.com/"" target=""_blank""><font color=""blue"">快车联盟</font></a>"
	Response.Write "</div>"
	Response.Write "</td>" & vbcrlf
	Response.Write " </tr>" & vbcrlf
	Response.Write " </table>" & vbcrlf
	Response.Write "</form>" & vbcrlf
	Response.Write "<script>setdisabled(" & Rs_e("parentid") &");</script>"
	Set Rs_e = Nothing
End Sub
'================================================
'过程名：savenew
'作  用：保存新的服务器
'================================================
Sub savenew()
	Dim downid,rootid,ParentID
	Dim depth,orders,Maxrootid
	Dim strParent,neworders
	Dim DownloadPath,Server_Url
	
	On Error Resume Next
	'保存添加服务器信息
	If KS.G("DownloadName") = "" Then
		Call KS.AlertHistory("请输入服务器名称!",-1)
		Exit Sub
	End If
	If KS.G("servers") = "" Then
		Call KS.AlertHistory("请选择服务器!",-1)
		Exit Sub
	End If
	If KS.G("DownloadPath") = "" and KS.G("servers")<>0 Then
		Call KS.AlertHistory("服务器路径不能为空!",-1)
		Exit Sub
	End If
	Server_Url = Replace(Request.Form("DownloadPath"), "\", "/")
	If Right(Server_Url, 1) <> "/" Then
		DownloadPath = Server_Url
	Else
		DownloadPath = Server_Url
	End If
	Set Rs = CreateObject("adodb.recordset")
	If Request.Form("servers") <> "0" Then
		SQL = "SELECT rootid,downid,depth,orders,strparent FROM KS_DownSer WHERE downid=" & Request("servers")
		Rs.Open SQL, Conn, 1, 1
		rootid = Rs(0)
		ParentID = Rs(1)
		depth = Rs(2)
		orders = Rs(3)
		If depth + 1 > 2 Then
			Call KS.AlertHistory("本系统限制最多只能有2级子服务器",-1)
			Exit Sub
		End If
		strParent = Rs(4)
		Rs.Close
		neworders = orders
		SQL = "SELECT MAX(orders) FROM KS_DownSer WHERE ParentID=" & Request("servers")
		Rs.Open SQL, Conn, 1, 1
		If Not (Rs.EOF And Rs.BOF) Then
			neworders = Rs(0)
		End If
		If IsNull(neworders) Then neworders = orders
		Rs.Close
		Conn.Execute ("UPDATE KS_DownSer SET orders=orders+1 WHERE orders>" & CInt(neworders) & "")
	Else
		SQL = "SELECT MAX(rootid) FROM KS_DownSer"
		Rs.Open SQL, Conn, 1, 1
		If Rs.BOF And Rs.EOF Then
			Maxrootid = 1
		Else
			Maxrootid = Rs(0) + 1
		End If
		If IsNull(Maxrootid) Then Maxrootid = 1
		Rs.Close
	End If
	If Maxrootid = 0 Then Maxrootid = 1
	
	SQL = "SELECT downid FROM KS_DownSer WHERE downid=" & Request("newdownid")
	Rs.Open SQL, Conn, 1, 1
	If Not (Rs.EOF And Rs.BOF) Then
		Call KS.AlertHistory("您不能指定和别的服务器一样的序号!",-1)
		Exit Sub
	Else
		downid = CLng(Request("newdownid"))
	End If
	Rs.Close
	
	SQL = "SELECT * FROM KS_DownSer"
	Rs.Open SQL, Conn, 1, 3
	Rs.AddNew
	If Request("servers") <> "0" Then
		Rs("depth") = depth + 1
		Rs("rootid") = rootid
		Rs("orders") = neworders + 1
		Rs("parentid") = KS.G("servers")
		If strParent = "0" Then
			Rs("strparent") = KS.G("servers")
		Else
			Rs("strparent") = strParent & "," & KS.G("servers")
		End If
	Else
		Rs("depth") = 0
		Rs("rootid") = Maxrootid
		Rs("orders") = 0
		Rs("parentid") = 0
		Rs("strparent") = 0
	End If
	Rs("child") = 0
	Rs("downid") = KS.G("newdownid")
	Rs("DownloadName") = Replace(KS.G("DownloadName"), "|", "")
	Rs("DownloadPath") = Replace(DownloadPath, "|", "")
	Rs("isDisp") = KS.G("isDisp")
	Rs("ChannelID") = KS.G("ChannelID")
	Rs("isLock") = 0
	Rs("IsOuter") = KS.ChkClng(KS.G("IsOuter"))
	Rs("UnionID") = KS.G("UnionID")
	Rs("selfont") = KS.G("selfont")
	Rs("AllDownHits") = 0
	Rs("DayDownHits") = 0
	Rs("HitsTime") = Now()
	Rs.Update
	Rs.Close
	If Request("servers") <> "0" Then
		If depth > 0 Then Conn.Execute ("update KS_DownSer set child=child+1 where downid in (" & strParent & ")")
		Conn.Execute ("update KS_DownSer set child=child+1 where downid=" & Request("servers"))
	End If
	call KS.Alert("服务器添加成功！","KS.DownServer.asp?ChannelID=" & ChannelID)
	Set Rs = Nothing
End Sub
'================================================
'过程名：saveedit
'作  用：保存编辑
'================================================
Sub saveedit()
	Dim newdownid,Maxrootid,ParentID
	Dim depth,Child,strParent,rootid
	Dim iparentid,istrparent
	Dim trs,brs,mrs,k
	Dim nstrparent,mstrparent,ParentSql
	Dim boardcount,DownloadPath,Server_Url
	
	On Error Resume Next
	If CLng(Request("editid")) = CLng(Request("servers")) Then
		Call KS.AlertHistory("所属服务器不能指定自己",-1)
		Exit Sub
	End If
	Server_Url = Replace(Request.Form("DownloadPath"), "\", "/")
	If Right(Server_Url, 1) <> "/" Then
		DownloadPath = Server_Url
	Else
		DownloadPath = Server_Url
	End If
	Set Rs = CreateObject("adodb.recordset")
	SQL = "SELECT * FROM KS_DownSer WHERE downid=" & CLng(Request("editid"))
	Rs.Open SQL, Conn, 1, 3
	newdownid = Rs("downid")
	ParentID = Rs("parentid")
	iparentid = Rs("parentid")
	strParent = Rs("strparent")
	depth = Rs("depth")
	Child = Rs("child")
	rootid = Rs("rootid")
	If ParentID = 0 Then
		If CLng(Request("servers")) <> 0 Then
			Set trs = Conn.Execute("select rootid from KS_DownSer where downid=" & Request("servers"))
			If rootid = trs(0) Then
				Call KS.AlertHistory("您不能指定该服务器的下属服务器作为所属服务器",-1)
				Exit Sub
			End If
		End If
	Else
		Set trs = Conn.Execute("select downid from KS_DownSer where strparent like '%" & strParent & "%' and downid=" & Request("servers"))
		If Not (trs.EOF And trs.BOF) Then
			Call KS.AlertHistory("您不能指定该服务器的下属服务器作为所属服务器",-1)
			Exit Sub
		End If
	End If
	If ParentID = 0 Then
		ParentID = Rs("downid")
		iparentid = 0
	End If
	Rs("DownloadName") = Replace(KS.G("DownloadName"), "|", "")
	Rs("DownloadPath") = Replace(DownloadPath, "|", "")
	Rs("isDisp")       = KS.G("isDisp")
	Rs("UserGroup") = Request.Form("UserGroup")
	Rs("ChannelID") = Request.Form("ChannelID")
	Rs("DownPoint") = KS.ChkClng(KS.G("DownPoint"))
	Rs("isLock") = 0
	Rs("IsOuter") = KS.ChkClng(KS.G("IsOuter"))
	Rs("UnionID") = KS.G("UnionID")
	Rs("selfont") = Trim(Request.Form("selfont"))
	Rs.Update
	Rs.Close:Set Rs = Nothing
	Set mrs = Conn.Execute("select max(rootid) from KS_DownSer")
	Maxrootid = mrs(0) + 1
	If KS.ChkClng(ParentID) <> KS.ChkClng(Request("servers")) And Not (iparentid = 0 And CInt(Request("servers")) = 0) Then
		If iparentid > 0 And CInt(Request("servers")) = 0 Then
			Conn.Execute ("update KS_DownSer set depth=0,orders=0,rootid=" & Maxrootid & ",parentid=0,strparent='0' where downid=" & newdownid)
			strParent = strParent & ","
			Set Rs = Conn.Execute("select count(*) from KS_DownSer where strparent like '%" & strParent & "%'")
			boardcount = Rs(0)
			If IsNull(boardcount) Then
				boardcount = 1
			Else
				boardcount = boardcount + 1
			End If
			Conn.Execute ("update KS_DownSer set child=child-" & boardcount & " where downid=" & iparentid)
			For i = 1 To depth
				Set Rs = Conn.Execute("select parentid from KS_DownSer where downid=" & iparentid)
				If Not (Rs.EOF And Rs.BOF) Then
					iparentid = Rs(0)
					Conn.Execute ("update KS_DownSer set child=child-" & boardcount & " where downid=" & iparentid)
				End If
			Next
			If Child > 0 Then
				i = 0
				Set Rs = Conn.Execute("select * from KS_DownSer where strparent like '%" & strParent & "%'")
				Do While Not Rs.EOF
					i = i + 1
					mstrparent = Replace(Rs("strparent"), strParent, "")
					Conn.Execute ("update KS_DownSer set depth=depth-" & depth & ",rootid=" & Maxrootid & ",strparent='" & mstrparent & "' where downid=" & Rs("downid"))
					Rs.MoveNext
				Loop
			End If
		ElseIf iparentid > 0 And CInt(Request("servers")) > 0 Then
			Set trs = Conn.Execute("select * from KS_DownSer where downid=" & Request("servers"))
			strParent = strParent & ","
			Set Rs = Conn.Execute("select count(*) from KS_DownSer where strparent like '%" & strParent & "%'")
			boardcount = Rs(0)
			If IsNull(boardcount) Then boardcount = 1
			Conn.Execute ("update KS_DownSer set orders=orders + " & boardcount & " + 1 where rootid=" & trs("rootid") & " and orders>" & trs("orders") & "")
			Conn.Execute ("update KS_DownSer set depth=" & trs("depth") & "+1,orders=" & trs("orders") & "+1,rootid=" & trs("rootid") & ",ParentID=" & Request("servers") & ",strparent='" & trs("strparent") & "," & trs("downid") & "' where downid=" & newdownid)
			i = 1
			SQL = "select * from KS_DownSer where strparent like '%" & strParent & "%' order by orders"
			Set Rs = Conn.Execute(SQL)
			Do While Not Rs.EOF
				i = i + 1
				istrparent = trs("strparent") & "," & trs("downid") & "," & Replace(Rs("strparent"), strParent, "")
				Conn.Execute ("update KS_DownSer set depth=depth+" & trs("depth") & "-" & depth & "+1,orders=" & trs("orders") & "+" & i & ",rootid=" & trs("rootid") & ",strparent='" & istrparent & "' where downid=" & Rs("downid"))
				Rs.MoveNext
			Loop
			ParentID = Request("servers")
			If rootid = trs("rootid") Then
				Conn.Execute ("update KS_DownSer set child=child+" & i & " where (not ParentID=0) and downid=" & ParentID)
				For k = 1 To trs("depth")
					Set Rs = Conn.Execute("select parentid from KS_DownSer where (not ParentID=0) and downid=" & ParentID)
					If Not (Rs.EOF And Rs.BOF) Then
						ParentID = Rs(0)
						Conn.Execute ("update KS_DownSer set child=child+" & i & " where (not ParentID=0) and  downid=" & ParentID)
					End If
				Next
				Conn.Execute ("update KS_DownSer set child=child-" & i & " where (not ParentID=0) and downid=" & iparentid)
				For k = 1 To depth
					Set Rs = Conn.Execute("select parentid from KS_DownSer where (not ParentID=0) and downid=" & iparentid)
					If Not (Rs.EOF And Rs.BOF) Then
						iparentid = Rs(0)

						Conn.Execute ("update KS_DownSer set child=child-" & i & " where (not ParentID=0) and  downid=" & iparentid)
					End If
				Next
			Else

				Conn.Execute ("update KS_DownSer set child=child+" & i & " where downid=" & ParentID)
				For k = 1 To trs("depth")
					Set Rs = Conn.Execute("select parentid from KS_DownSer where downid=" & ParentID)
					If Not (Rs.EOF And Rs.BOF) Then
						ParentID = Rs(0)
						Conn.Execute ("update KS_DownSer set child=child+" & i & " where downid=" & ParentID)
					End If
				Next
				Conn.Execute ("update KS_DownSer set child=child-" & i & " where downid=" & iparentid)
				For k = 1 To depth
					Set Rs = Conn.Execute("select parentid from KS_DownSer where downid=" & iparentid)
					If Not (Rs.EOF And Rs.BOF) Then
						iparentid = Rs(0)
						Conn.Execute ("update KS_DownSer set child=child-" & i & " where downid=" & iparentid)
					End If
				Next
			End If
		Else
			Set trs = Conn.Execute("select * from KS_DownSer where downid=" & Request("servers"))
			Set Rs = Conn.Execute("select count(*) from KS_DownSer where rootid=" & rootid)
			boardcount = Rs(0)
			ParentID = Request("servers")
			Conn.Execute ("update KS_DownSer set child=child+" & boardcount & " where downid=" & ParentID)
			For k = 1 To trs("depth")
				Set Rs = Conn.Execute("select parentid from KS_DownSer where downid=" & ParentID)
				If Not (Rs.EOF And Rs.BOF) Then
					ParentID = Rs(0)
					Conn.Execute ("update KS_DownSer set child=child+" & boardcount & " where downid=" & ParentID)
				End If

			Next
			Conn.Execute ("update KS_DownSer set orders=orders + " & boardcount & " + 1 where rootid=" & trs("rootid") & " and orders>" & trs("orders") & "")
			i = 0
			SQL = "select * from KS_DownSer where rootid=" & rootid & " order by orders"
			Set Rs = Conn.Execute(SQL)
			Do While Not Rs.EOF
				i = i + 1
				If Rs("parentid") = 0 Then
					If trs("strparent") = "0" Then
						strParent = trs("downid")
					Else
						strParent = trs("strparent") & "," & trs("downid")
					End If
					Conn.Execute ("update KS_DownSer set depth=depth+" & trs("depth") & "+1,orders=" & trs("orders") & "+" & i & ",rootid=" & trs("rootid") & ",strparent='" & strParent & "',parentid=" & Request("servers") & " where downid=" & Rs("downid"))
				Else
					If trs("strparent") = "0" Then
						strParent = trs("downid") & "," & Rs("strparent")
					Else
						strParent = trs("strparent") & "," & trs("downid") & "," & Rs("strparent")
					End If
					Conn.Execute ("update KS_DownSer set depth=depth+" & trs("depth") & "+1,orders=" & trs("orders") & "+" & i & ",rootid=" & trs("rootid") & ",strparent='" & strParent & "' where downid=" & Rs("downid"))
				End If
				Rs.MoveNext
			Loop
		End If
	End If
	Call KS.Alert("服务器修改成功！","KS.DownServer.asp?ChannelID=" & ChannelID)
	Set Rs = Nothing
	Set mrs = Nothing
	Set trs = Nothing
End Sub
'================================================
'过程名：DelDownPath
'作  用：删除服务器
'================================================
Sub DelDownPath()
	Dim rsUsage
	
	On Error Resume Next
	Set Rs = Conn.Execute("select strparent,child,depth,rootid from KS_DownSer where downid=" & Request("editid"))
	If Not (Rs.EOF And Rs.BOF) Then
		If Rs(1) > 0 Then
			Call KS.AlertHistory("该服务器含有下载路径，请删除其下载路径后再进行删除本服务器的操作",-1)
			Exit Sub
		End If
		If Rs("depth") = 0 Then
		     Dim RSC:Set RSC=Server.CreateObject("ADODB.Recordset")
			 RSC.Open "Select DownUrls From "& KS.C_S(ChannelID,2),conn,1,1
			 Do While Not RSC.Eof
			   Dim N,DArr
			   DArr=Split(RSC(0),"|||")
			   For N=0 To Ubound(DArr)
			    If Split(Darr(N),"|")(0)=Rs("rootid") Then
				 Call KS.AlertHistory("该下载服务器正在使用中，不能删除!",-1)
				 Exit Sub
			    End If
			   Next
			   RSC.MoveNext
			 Loop
			 RSC.Close:Set RSC=Nothing
	
		End If
		If Rs(2) > 0 Then
			Conn.Execute ("UPDATE KS_DownSer SET child=child-1 WHERE downid in (" & Rs(0) & ")")
		End If
		SQL = "DELETE FROM KS_DownSer WHERE downid=" & Request("editid")
		Conn.Execute (SQL)
	End If
	Set Rs = Nothing
	Call KS.Alert("服务器删除成功！","KS.DownServer.asp?ChannelID=" & ChannelID)
End Sub
'================================================
'过程名：isLock
'作  用：锁定服务器
'================================================
Sub isLock()

	conn.Execute ("update KS_DownSer set isLock = 1 where downid in (" & Request("editid") & ")")
	Response.Redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub
'================================================
'过程名：FreeLock
'作  用：解除服务器锁定
'================================================
Sub FreeLock()
	conn.Execute ("update KS_DownSer set isLock = 0 where downid in (" & Request("editid") & ")")
	Response.Redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub
'================================================
'过程名：serverorders
'作  用：服务器排序
'================================================
Sub serverorders()
	Dim trs
	Dim uporders
	Dim doorders
	
	Response.Write " <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=center>" & vbcrlf
	Response.Write " <tr class='sort'>" & vbcrlf
	Response.Write " <td colspan=2>服务器路径重新排序修改(请在相应服务器的排序表单内输入相应的排列序号)" & vbcrlf
	Response.Write " </td>" & vbcrlf
	Response.Write " </tr>" & vbcrlf
	Set Rs = CreateObject("Adodb.recordset")
	SQL = "SELECT * FROM KS_DownSer WHERE ChannelID=" & ChannelID & " ORDER BY RootID,orders"
	Rs.Open SQL, Conn, 1, 1
	If Rs.BOF And Rs.EOF Then
		Response.Write "还没有相应的服务器。"
	Else
		Do While Not Rs.EOF
			Response.Write "<form action=KS.DownServer.asp?action=updateorders method=post><tr height='23' class='list' onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'""><td width=""50%"">"
			If Rs("depth") = 1 Then Response.Write "&nbsp;&nbsp;<font color=""#666666"">├</font>"
			If Rs("depth") > 1 Then
				For i = 2 To Rs("depth")
					Response.Write "&nbsp;&nbsp;<font color=""#666666"">│</font>"
				Next
				Response.Write "&nbsp;&nbsp;<font color=""#666666"">├</font> "
			End If
			If Rs("parentid") = 0 Then Response.Write ("<b>")
			Response.Write Rs("DownloadName")
			If Rs("child") > 0 Then Response.Write "(" & Rs("child") & ")"
			Response.Write "</td><td width=""50%"">"
			If Rs("ParentID") > 0 Then
				Set trs = Conn.Execute("SELECT COUNT(*) FROM KS_DownSer WHERE ParentID=" & Rs("ParentID") & " and orders<" & Rs("orders") & "")
				uporders = trs(0)
				If IsNull(uporders) Then uporders = 0
				Set trs = Conn.Execute("SELECT COUNT(*) FROM KS_DownSer WHERE ParentID=" & Rs("ParentID") & " and orders>" & Rs("orders") & "")
				doorders = trs(0)
				If IsNull(doorders) Then doorders = 0
				If uporders > 0 Then
					Response.Write "<select name=uporders size=1><option value=0>↑</option>"
					For i = 1 To uporders
						Response.Write "<option value=" & i & ">↑" & i & "</option>"
					Next
					Response.Write "</select>"
				End If
				If doorders > 0 Then
					If uporders > 0 Then Response.Write "&nbsp;"
					Response.Write "<select name=doorders size=1><option value=0>↓</option>"
					For i = 1 To doorders
						Response.Write "<option value=" & i & ">↓" & i & "</option>"
					Next
					Response.Write "</select>"
				End If
				If doorders > 0 Or uporders > 0 Then
					Response.Write vbcrlf & "<input type=""hidden"" name=ChannelID value="""
					Response.Write ChannelID
					Response.Write """>" & vbcrlf
					Response.Write "<input type=hidden name=""editID"" value=""" & Rs("downid") & """>&nbsp;<input type=submit name=Submit class=button value='修 改'>"
				End If
			End If
			Response.Write "</td></tr>"
			Response.Write ("<tr><td colspan=6 background='images/line.gif'></td></tr>")
			Response.Write "</form>"
			uporders = 0
			doorders = 0
			Rs.MoveNext
		Loop
	End If
	Rs.Close
	Set Rs = Nothing
	Response.Write "</table>" & vbcrlf
End Sub
'================================================
'过程名：updateorders
'作  用：更新服务器排序
'================================================
Sub updateorders()
	Dim ParentID
	Dim orders
	Dim strParent
	Dim Child
	Dim uporders
	Dim doorders
	Dim oldorders
	Dim trs
	Dim ii
	If Not IsNumeric(Request("editID")) Then
		Call KS.AlertHistory("非法的参数！",-1)
		Exit Sub
	End If
	If Request("uporders") <> "" And Not CInt(Request("uporders")) = 0 Then
		If Not IsNumeric(Request("uporders")) Then
			Call KS.AlertHistory("非法的参数！",-1)
			Exit Sub
		ElseIf CInt(Request("uporders")) = 0 Then
			Call KS.AlertHistory("请选择要提升的数字！",-1)
			Exit Sub
		End If
		Set Rs = Conn.Execute("SELECT ParentID,orders,strparent,child FROM KS_DownSer where downid=" & Request("editID"))
		ParentID = Rs(0)
		orders = Rs(1)
		strParent = Rs(2) & "," & Request("editID")
		Child = Rs(3)
		i = 0
		If Child > 0 Then
			Set Rs = Conn.Execute("SELECT COUNT(*) FROM KS_DownSer WHERE strparent like '%" & strParent & "%'")
			oldorders = Rs(0)
		Else
			oldorders = 0
		End If
		Set Rs = Conn.Execute("SELECT downid,orders,child,strparent FROM KS_DownSer WHERE ParentID=" & ParentID & " and orders<" & orders & " order by orders desc")
		Do While Not Rs.EOF
			i = i + 1
			If CInt(Request("uporders")) >= i Then
				If Rs(2) > 0 Then
					ii = 0
					Set trs = Conn.Execute("select downid,orders from KS_DownSer where strparent like '%" & Rs(3) & "," & Rs(0) & "%' order by orders")
					If Not (trs.EOF And trs.BOF) Then
						Do While Not trs.EOF
							ii = ii + 1
							Conn.Execute ("update KS_DownSer set orders=" & orders & "+" & oldorders & "+" & ii & " where downid=" & trs(0))
							trs.MoveNext
						Loop
					End If
				End If
				Conn.Execute ("update KS_DownSer set orders=" & orders & "+" & oldorders & " where downid=" & Rs(0))
				If CInt(Request("uporders")) = i Then uporders = Rs(1)
			End If
			orders = Rs(1)
			Rs.MoveNext
		Loop
		Conn.Execute ("update KS_DownSer set orders=" & uporders & " where downid=" & Request("editID"))
		If Child > 0 Then
			i = uporders
			Set Rs = Conn.Execute("select downid from KS_DownSer where strparent like '%" & strParent & "%' order by orders")
			Do While Not Rs.EOF
				i = i + 1
				Conn.Execute ("update KS_DownSer set orders=" & i & " where downid=" & Rs(0))
				Rs.MoveNext
			Loop
		End If
		Set Rs = Nothing
		Set trs = Nothing
	ElseIf Request("doorders") <> "" Then
		If Not IsNumeric(Request("doorders")) Then
			Call KS.AlertHistory("非法的参数！",-1)
			Exit Sub
		ElseIf CInt(Request("doorders")) = 0 Then
			Call KS.AlertHistory("请选择要下降的数字！",-1)
			Exit Sub
		End If
		Set Rs = Conn.Execute("select ParentID,orders,strparent,child from KS_DownSer where downid=" & Request("editID"))
		ParentID = Rs(0)
		orders = Rs(1)
		strParent = Rs(2) & "," & Request("editID")
		Child = Rs(3)
		i = 0
		If Child > 0 Then
			Set Rs = Conn.Execute("select count(*) from KS_DownSer where strparent like '%" & strParent & "%'")
			oldorders = Rs(0)
		Else
			oldorders = 0
		End If
		Set Rs = Conn.Execute("select downid,orders,child,strparent from KS_DownSer where ParentID=" & ParentID & " and orders>" & orders & " order by orders")
		Do While Not Rs.EOF
			i = i + 1
			If CInt(Request("doorders")) >= i Then
				If Rs(2) > 0 Then
					ii = 0
					Set trs = Conn.Execute("select downid,orders from KS_DownSer where strparent like '%" & Rs(3) & "," & Rs(0) & "%' order by orders")
					If Not (trs.EOF And trs.BOF) Then
						Do While Not trs.EOF
							ii = ii + 1
							Conn.Execute ("update KS_DownSer set orders=" & orders & "+" & ii & " where downid=" & trs(0))
							trs.MoveNext
						Loop
					End If
				End If
				Conn.Execute ("update KS_DownSer set orders=" & orders & " where downid=" & Rs(0))
				If CInt(Request("doorders")) = i Then doorders = Rs(1)
			End If
			orders = Rs(1)
			Rs.MoveNext
		Loop
		Conn.Execute ("UPDATE KS_DownSer SET orders=" & doorders & " WHERE downid=" & Request("editID"))
		If Child > 0 Then
			i = doorders
			Set Rs = Conn.Execute("SELECT downid from KS_DownSer WHERE strparent like '%" & strParent & "%' ORDER BY orders")
			Do While Not Rs.EOF
				i = i + 1
				Conn.Execute ("UPDATE KS_DownSer SET orders=" & i & " WHERE downid=" & Rs(0))
				Rs.MoveNext
			Loop
		End If
		Set Rs = Nothing
		Set trs = Nothing
	End If
	Response.Redirect "KS.DownServer.asp?action=serverorders&ChannelID=" & Request("ChannelID")
End Sub
%> 

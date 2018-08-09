<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.CollectCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Collect_ItemHistory
KSCls.Kesion()
Set KSCls = Nothing

Class Collect_ItemHistory
        Private KS
		Private KMCObj
		Private ConnItem
		Private i, totalPut, CurrentPage, SqlStr
		Private Rs, Sql, SqlItem, RSObj, Action, FoundErr, ErrMsg
		Private HistoryID, ItemID, ChannelID, ClassID, SpecialID, ArticleID, Title, CollecDate, NewsUrl, Result
		Private Arr_History, Arr_ArticleID, i_Arr, Del, Flag
		Private MaxPerPage
		Private Sub Class_Initialize()
		  MaxPerPage = 20
		  Set KS=New PublicCls
		  Set KMCObj=New CollectPublicCls
		  Set ConnItem = KS.ConnItem()
		End Sub
        Private Sub Class_Terminate()
		 Call KS.CloseConnItem()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KMCObj=Nothing
		End Sub
		Sub Kesion()
		If Not KS.ReturnPowerResult(0, "KMCL10003") Then
		  Response.Write "<script src='../../ks_inc/jquery.js'></script>"
		  Response.Write ("<script>$(parent.document).find('#BottomFrame')[0].src='javascript:history.back();';</script>")
		  Call KS.ReturnErr(1, "")
		End If
		ChannelID=KS.ChkClng(KS.G("ChannelID"))
		'response.write "channelid=" & channelid
		If ChannelID=0 Then 
		 Call KS.AlertHistory("参数传递出错!",-1)
		 response.end
		End IF
		If Request("page") <> "" Then
			  CurrentPage = CInt(Request("page"))
		Else
			  CurrentPage = 1
		End If
		FoundErr = False
		Action = Trim(Request("Action"))
		If FoundErr <> True Then
		   Call Main
		Else
		   Call KS.AlertHistory(ErrMsg,-1)
		End If
		End Sub
		Sub Main()
		    Dim HistoryID:HistoryID = Trim(KS.G("HistoryID"))
			Dim Action:Action=KS.G("Action")
			Dim Page:Page = KS.G("Page")
		    If Action = "del" Then
			  HistoryID = Replace(HistoryID, " ", "")
			  If KS.IsNul(HistoryID) Then KS.AlertHintScript "对不起,没有选择记录!"
			  ConnItem.Execute ("Delete From KS_History Where HistoryID In(" & HistoryID & ")")
			 Response.Write "<script>location.href='Collect_ItemHistory.asp?ChannelID="& ChannelID & "&Page=" & Page & "';</script>"
			ElseIf Action="DelSucceed" Then
			  ConnItem.Execute ("Delete From KS_History  Where  Result=True")
			  Response.Write "<script>location.href='Collect_ItemHistory.asp?ChannelID="& ChannelID & "&Page=" & Page & "';</script>"
			ElseIf Action="DelFailure" Then
			  ConnItem.Execute ("Delete From KS_History  Where  Result=False")
			  Response.Write "<script>location.href='Collect_ItemHistory.asp?ChannelID="& ChannelID & "&Page=" & Page & "';</script>"
			ElseIf Action = "delall" Then
			  ConnItem.Execute ("Delete From KS_History")
			 Response.Write "<script>location.href='Collect_ItemHistory.asp?ChannelID="& ChannelID & "&Page=" & Page & "';</script>"
			End If
			
		 Response.Write "<html>"
		 Response.Write "<head>"
		Response.Write "<title>采集系统</title>"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		Response.Write "<link rel=""stylesheet"" type=""text/css"" href=""../Include/Admin_Style.css"">"
		Response.Write "<script language=""JavaScript"">"
		Response.Write "var Page='" & CurrentPage & "';"
		Response.Write "</script>"
		Response.Write "<script language=""JavaScript"" src=""../../KS_Inc/common.js""></script>"
		Response.Write "<script language=""JavaScript"" src=""../../KS_Inc/jquery.js""></script>"
		Response.Write "<script language=""JavaScript"" src=""../Include/ContextMenu.js""></script>"
		Response.Write "<script language=""JavaScript"" src=""../Include/SelectElement.js""></script>"
		%>
		<script>
		var DocElementArrInitialFlag=false;
		var DocElementArr = new Array();
		var DocMenuArr=new Array();
		var SelectedFile='',SelectedFolder='';
		$(document).ready(function()
		{   if (DocElementArrInitialFlag) return;
			InitialDocElementArr('FolderID','HistoryID');
			InitialContextMenu();
			DocElementArrInitialFlag=true;
		});
		function InitialContextMenu()
		{	
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.SelectAllElement();",'全 选(A)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.DelRecords();",'删除选中(D)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.DelFailure();",'删除失败(S)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.DelSucceed();",'删除成功(F)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.DelAllRecords();",'删除全部(Y)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.location.reload();",'刷 新(Z)','disabled');
		}
		function DocDisabledContextMenu()
		{
			DisabledContextMenu('FolderID','HistoryID','删除选中(D)','','','','')
		}
		function DelRecords()
		{
		 GetSelectStatus('FolderID','HistoryID');
		 if (SelectedFile!='')
		  {
		   if (confirm('真的要删除选中的记录吗?'))
			location="Collect_ItemHistory.asp?ChannelID=<%=ChannelID%>&Action=del&Page="+Page+"&HistoryID="+SelectedFile;
		  }
		 else
		  alert('请选择要删除的记录!');
		  SelectedFile='';
		}
		function DelSucceed()
		{
		 if (confirm('真的要清除所有成功记录吗?'))
			location="Collect_ItemHistory.asp?ChannelID=<%=ChannelID%>&Action=DelSucceed&Page="+Page;
		}
		function DelFailure()
		{
		 if (confirm('真的要清除所有记录吗?'))
			location="Collect_ItemHistory.asp?ChannelID=<%=ChannelID%>&Action=DelFailure&Page="+Page;
		}
		function DelAllRecords()
		{
		 if (confirm('真的要清除所有记录吗?'))
			location="Collect_ItemHistory.asp?ChannelID=<%=ChannelID%>&Action=delall&Page="+Page;
		}
		function GetKeyDown()
		{ 
		if (event.ctrlKey)
		  switch  (event.keyCode)
		  {  case 90 : location.reload(); break;
			 case 65 : SelectAllElement();break;
			 case 68 : DelRecords('');break;
			 case 83 : DelSucceed('');break;
			 case 70 : DelFailure('');break;
			 case 89 : event.keyCode=0;event.returnValue=false;DelAllRecords('');break;
		   }	
		else	
		 if (event.keyCode==46) DelRecords('');
		}
		function CheckAll(form)
			{
			  for (var i=0;i<form.elements.length;i++)
				{
				var e = form.elements[i];
				if (e.Name != "chkAll")
				   e.checked = form.chkAll.checked;
				}
			}
		</script>
		<%
		Response.Write "</head>"
		Response.Write "<body scroll=no topmargin=""0"" leftmargin=""0"" onclick=""SelectElement();"" onkeydown=""GetKeyDown();"" onselectstart=""return false;"">"
		Response.Write "<ul id='menu_top'>"
		Response.Write "<li class='parent' onclick='location.href=""Collect_ItemModify.asp?channelid=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/a.gif' border='0' align='absmiddle'>新建项目</span></li>"
		Response.Write "<li class='parent' onclick='location.href=""Collect_ItemFilters.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/move.gif' border='0' align='absmiddle'>过滤设置</span></li>"
		Response.Write "<li class='parent' onclick='location.href=""Collect_IntoDatabase.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/save.gif' border='0' align='absmiddle'>审核入库</span></li>"
		Response.Write "<li class='parent' onclick='location.href=""Collect_ItemHistory.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/Recycl.gif' border='0' align='absmiddle'>历史记录</span></li>"
		Response.Write "<li class='parent' onclick='location.href=""Collect_Field.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/addjs.gif' border='0' align='absmiddle'>自定义字段</span></li>"
		Response.Write "<li class='parent' onclick='location.href=""Collect_main.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/back.gif' border='0' align='absmiddle'>回上一级</span></li><li></li>"

			Response.Write "<div id='go'><select OnChange=""location.href=this.value"" style='width:120px' name='id'>"
			Response.Write "<option value='Collect_ItemHistory.asp?channelid=" & channelid & "'>快速查找历史记录</option>"
			Response.Write "<option value='Collect_ItemHistory.asp?channelid=" & channelid & "'>查看全部记录</option>"
			Response.Write "<option value='Collect_ItemHistory.asp?channelid=" & channelid & "&Action=Succeed'>查看成功记录</option>"
			Response.Write "<option value='Collect_ItemHistory.asp?channelid=" & channelid & "&Action=Failure'>查看失败记录</option>"
			
			Response.Write " </select>"
			Response.Write "</div>"
			Response.Write ("</ul>")
            

									
		Set RSObj = Server.CreateObject("adodb.recordset")
		'SqlItem = "select * From KS_History Where ChannelID=" & ChannelID
		SqlItem = "select * From KS_History"
		If Action = "Succeed" Then
		   SqlItem = SqlItem & "  where Result=True"
		   Flag = "成 功 记 录"
		ElseIf Action = "Failure" Then
		   SqlItem = SqlItem & " where Result=False"
		   Flag = "失 败 记 录"
		Else
		   Flag = "所 有 记 录"
		End If
		Response.Write ("<div style=""height:94%; overflow: auto; width:100%"" align=""center"">")
		Response.Write "  <table border=""0"" cellspacing=""1"" width=""100%"" cellpadding=""0"">"
		Response.Write "     <tr style=""padding: 0px 2px;"">"
		Response.Write "      <td width=""435"" height=""22"" align=""center"" class=sort>标题</td>"
		Response.Write "      <td width=""214"" align=""center"" class=sort>项目名称</td>"
		Response.Write "      <td width=""123"" height=""22"" align=""center"" class=sort>所属系统</td>"
		Response.Write "      <td width=""120"" height=""22"" align=""center"" class=sort>(频道)栏目</td>"
		Response.Write "      <td width=""126"" align=""center"" class=sort>来源</td>"
		 Response.Write "     <td width=""87"" align=""center"" class=sort>结果</td>"
		 Response.Write "    </tr>"
		
		If Request("page") <> "" Then
			CurrentPage = CInt(Request("Page"))
		Else
			CurrentPage = 1
		End If
		SqlItem = SqlItem & " order by HistoryID DESC"
		RSObj.Open SqlItem, ConnItem, 1, 1
		 If Not RSObj.EOF Then
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
		
			
		Response.Write "</table>"
		Response.Write "</div>"
		Response.Write "</body>"
		Response.Write "</html>"
		End Sub
		Sub showContent()
		   Response.Write "<form name='myform' method='Post' action='?Page=" & CurrentPage & "&channelid=" & channelid & "'>"
		 Do While Not RSObj.EOF
			 Response.Write ("<tr>")
			 Response.Write (" <td class='splittd' width=""435"" height=""18"">")
				Response.Write "<input type='checkbox' name='HistoryID' value='" &RSObj("HistoryID") & "'><span  HistoryID='" & RSObj("HistoryID") & "'><img src='../Images/folder/TheSmallWordNews1.gif'  align='absmiddle'>"
				  Response.Write "  <span style='cursor:default;'>" & KS.GotTopic(RSObj("Title"), 40) & "</span></span>"
			  Response.Write ("</td> ")
			  Response.Write ("<td class='splittd' width=""214"" align=""center"">" & KMCObj.Collect_ShowItem_Name(RSObj("ItemID"), ConnItem) & "</td>")
			  Response.Write ("<td class='splittd' width=""123"" align=""center"">" & KS.C_S(ChannelID,1) & "</td>")
			  Response.Write ("<td class='splittd' width=""120"" align=""center"">" & KMCObj.Collect_ShowClass_Name(RSObj("ChannelID"), RSObj("ClassID")) & "</td>")
			  Response.Write ("<td class='splittd' width=""126"" align=""center""><a href=""" & RSObj("NewsUrl") & """ target=""_blank"" title=""" & RSObj("NewsUrl") & """>点击访问</a></td>")
			  Response.Write (" <td width=""87"" align=""center"">")
			  If RSObj("Result") = True Then
				   Response.Write "<font color=red>成功</font>"
				ElseIf RSObj("Result") = False Then
				   Response.Write "<font color=red>失败</font>"
				Else
				   Response.Write "<font color=red>异常</font>"
				End If
			  Response.Write (" </td></tr> ")
				   i = i + 1
				   If i >=MaxPerPage Then
					  Exit Do
				   End If
				RSObj.MoveNext
		   Loop
		RSObj.Close
		Set RSObj = Nothing
		   Response.Write "<tr><td colspan=7 height='25'><input name='chkAll' type='checkbox' id='chkAll' onclick=CheckAll(this.form) value='checkbox'>批量选择 <input type='submit' value='批量删除' class='button' onclick=""this.form.action='Collect_ItemHistory.asp?ChannelID=" & ChannelID & "&Action=del&Page=" & currentPage & "'"">&nbsp;<input type='button' onclick=""DelAllRecords();"" value='删除全部记录' class='button'>&nbsp;<input type='button' onclick=""DelSucceed();"" value='删除所有成功记录' class='button'>&nbsp;<input type='button' onclick=""DelFailure();"" value='删除所有失败记录' class='button'></td></tr>"
		   Response.Write "</form>"
			Response.Write ("<tr><td height=""22"" colspan=""6"" align=""right"">")
		 Call KS.ShowPageParamter(totalPut, MaxPerPage, "Collect_ItemHistory.asp", True, "条", CurrentPage, "ChannelID=" & ChannelID &"&Action=" & Action)
		   Response.Write ("</td></tr>")
		End Sub
End Class
%> 

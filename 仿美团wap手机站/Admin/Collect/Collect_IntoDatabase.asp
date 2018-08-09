<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.CollectCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.BaseFunCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Collect_IntoDatabase
KSCls.Kesion()
Set KSCls = Nothing

const RepeatInto=1   '重复记录入库

Class Collect_IntoDatabase
        Private KS
		Private KMCObj
		Private ConnItem,ChannelID
		Private i,Arr_Field
		Private totalPut
		Private CurrentPage
		Private SqlStr
		Private RSObj,SuccNum,ErrNum
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
		'On Error Resume Next
		ChannelID=KS.ChkClng(KS.G("ChannelID"))
		If ChannelID=0 Then  ChannelID=1
		Response.Write "<script src=""../../ks_inc/jquery.js""></script>"
			If Not KS.ReturnPowerResult(0, "M010008") Then                  '权限检查
				Call KS.ReturnErr(1, "")   
				Response.End()
			End if
		
		  '自定义字段
		Dim RsItem:Set RsItem = Server.CreateObject("adodb.recordset")
		RsItem.Open "Select FieldName From KS_FieldItem Where ChannelID=" & ChannelID &" Order by OrderID", ConnItem, 1, 1
		If Not RsItem.EOF Then
			  Arr_Field = RsItem.GetRows()
		End If
		RsItem.Close:Set RsItem = Nothing

		
		If Request("page") <> "" Then
			  CurrentPage = CInt(Request("page"))
		Else
			  CurrentPage = 1
		End If
		Dim Rs, Sql, SqlItem, RSObj, Action, FoundErr, ErrMsg
		Dim ID, ItemID, ClassID, SpecialID, ArticleID, Title, CollecDate, NewsUrl, Result
		Dim Arr_History, Arr_ArticleID, i_Arr, Del, Flag
		Dim HistoryNum, i_His
		FoundErr = False
		Del = Trim(Request("Del"))
		Action = Trim(Request("Action"))
		If Action="View" Then
		 Call RecordView()
		 Exit Sub
		ElseIF Action="Into" Then
		 Call IntoDataBase()
		ElseIf Action <> "" Then
		  Call ExecuteAction
		End If
		If FoundErr <> True Then
		   Call Main
		Else
		   Call KS.Alert("出错!","")
		End If
		End Sub
		
		Sub Main()
		Dim SqlItem
		Response.Write "<html>"
		Response.Write "<head>"
		Response.Write "<title>采集系统</title>"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		Response.Write "<link rel=""stylesheet"" type=""text/css"" href=""../Include/Admin_Style.css"">"
		Response.Write "<script language=""JavaScript"">" & vbCrLf
		Response.Write "var Page='" & CurrentPage & "';" & vbCrLf
		Response.Write "</script>" & vbCrLf
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
		{     if (DocElementArrInitialFlag) return;
			InitialDocElementArr('FolderID','NewsID');
			InitialContextMenu();
			DocElementArrInitialFlag=true;
		});
		function InitialContextMenu()
		{	
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.ViewRecords();",'预 览(V)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.IntoDataBase();",'入 库(I)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.AllIntoDataBase();",'全部入库(W)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.VerificRecords();",'审核选中(S)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.VerificAllRecords();",'审核全部(U)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.SelectAllElement();",'全 选(A)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.DelRecords();",'删 除(D)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.DelAllRecords();",'删除全部(Y)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.location.reload();",'刷 新(Z)','disabled');
		}
		function DocDisabledContextMenu()
		{
			DisabledContextMenu('FolderID','NewsID','入 库(I),审核选中(S),删 除(D)','','','','')
		}
		function IntoDataBase()
		{
			GetSelectStatus('FolderID','NewsID');
		 if (SelectedFile!='')
		  {
		   if (confirm('真的要将选中的记录转移到主数据库吗?'))
			location.href='?action=Into&channelid=<%=channelid%>&inflag=1&ID='+SelectedFile;
		  }
		  else
		  alert('请选择要转移到主数据库的记录!');
		}
		function AllIntoDataBase()
		{
			if (confirm('真的要将所有记录转移到主数据库吗?'))
			 {location.href='?action=Into&channelid=<%=channelid%>&inflag=0';}
		}
		function ViewRecords(NewsID)
		{
		 if (NewsID!='')
		  {
			 window.open('?Action=View&ChannelID=<%=ChannelID%>&id='+NewsID,'new','');
		   }
		 else
		  {GetSelectStatus('FolderID','NewsID');
		   if (SelectedFile!='')
			{
			 if (SelectedFile.indexOf(',')==-1) 
			 {    
			window.open('Collect_View.asp?id='+SelectFile,'new','');
			 }
		   else
			alert('一次仅能查看一条记录信息!');
			}
			  SelectedFile='';
		  }
		}
		function DelRecords()
		{
		 GetSelectStatus('FolderID','NewsID');
		 if (SelectedFile!='')
		  {
		   if (confirm('真的要删除选中的记录吗?'))
			location="Collect_IntoDataBase.asp?ChannelID=<%=ChannelID%>&Action=del&ID="+SelectedFile+"&Page="+Page;
		  }
		 else
		  alert('请选择要删除的记录!');
		  SelectedFile='';
		}
		function DelAllRecords()
		{
		 if (confirm('真的要清除所有记录吗?'))
			location="Collect_IntoDataBase.asp?channelid=<%=channelid%>&Action=delall&Page="+Page;
		}
		function VerificAllRecords()
		{
		 if (confirm('真的要审核所有记录吗?'))
			location="Collect_IntoDataBase.asp?channelid=<%=channelid%>&Action=verificall&Page="+Page;
		}
		function VerificRecords()
		{  GetSelectStatus('FolderID','NewsID');
		if (SelectedFile!='')
		  {
			location="Collect_IntoDataBase.asp?channelid=<%=channelid%>&ID="+SelectedFile+"&Action=verific&Page="+Page;
		  }
		  else
		   alert('你没有选择要审核的记录!');
		}
		function GetKeyDown()
		{ 
		if (event.ctrlKey)
		  switch  (event.keyCode)
		  {  case 90 : location.reload(); break;
			 case 65 : SelectAllElement();break;
			 case 83 : event.keyCode=0;event.returnValue=false;VerificRecords('');break;
			 case 73 : event.keyCode=0;event.returnValue=false;IntoDataBase('');break;
			 case 87 : event.keyCode=0;event.returnValue=false;AllIntoDataBase('');break;
			 case 85 : event.keyCode=0;event.returnValue=false;VerificAllRecords('');break;
			 case 86 : event.keyCode=0;event.returnValue=false;ViewRecords('');break;
			 case 89 : event.keyCode=0;event.returnValue=false;DelAllRecords('');break;
			 case 68 : DelRecords('');break;
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
		Response.Write "<li class='parent' onclick='location.href=""Collect_main.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/back.gif' border='0' align='absmiddle'>回上一级</span></li>"
		Response.Write "</ul>"

		Response.Write ("<div style=""height:94%; overflow: auto; width:100%"" align=""center"">")
		If Request("Key")<>"" Then
		 Response.Write "<div style='width:300px;float:left;padding-left:4px;text-align:left;height:25px;line-height:25px;'>查看关键词含有“<font color=red>" & KS.S("Key") & "</font>”的文档:</div>"
		End If
        Response.Write "<div style='text-align:right'>请按模型入库<select id='channelid' name='channelid' onchange=""if (this.value!=0){location.href='?channelid='+this.value;}"">"
			Response.Write " <option value='0'>---请选择模型---</option>"
	
			If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
			Dim ModelXML,Node
			Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
			For Each Node In ModelXML.documentElement.SelectNodes("channel[@ks21=1][@ks6=1||@ks6=2||@ks6=5]")
			  If trim(ChannelID)=trim(Node.SelectSinglenode("@ks0").text) Then
			   Response.Write "<option value='" &Node.SelectSingleNode("@ks0").text &"' selected>" & Node.SelectSingleNode("@ks1").text & "</option>"

			  Else
			   Response.Write "<option value='" &Node.SelectSingleNode("@ks0").text &"'>" & Node.SelectSingleNode("@ks1").text & "</option>"
			  End If
			next
			Response.Write "</select></div>"		
		
		Response.Write "<table border=""0"" cellspacing=""0"" width=""100%"" cellpadding=""0"">"
		Response.Write "    <tr style=""padding: 0px 2px;"">"
		Response.Write "      <td width=""299"" height=""22"" align=""center"" class=sort>标题</td>"
		Response.Write "      <td width=""183"" align=""center"" class=sort>文章来源</td>"
		Response.Write "      <td width=""194"" height=""22"" align=""center"" class=sort>栏目</td>"
		Response.Write "      <td width=""211"" align=""center"" class=sort>时间</td>"
		Response.Write "      <td width=""152"" height=""22"" align=""center"" class=sort>管理操作</td>"
		Response.Write "    </tr>"
		Dim Param:Param=" Where 1=1"
		If Request("Key")<>"" Then
		  Param=Param& " and Title like '%" & KS.S("Key") & "%'"
		End If 
		Set RSObj = Server.CreateObject("adodb.recordset")
		SqlItem = "select * from " & KS.C_S(ChannelID,2) & Param
		If Request("page") <> "" Then
			CurrentPage = CInt(Request("Page"))
		Else
			CurrentPage = 1
		End If
		SqlItem = SqlItem & " order by ID DESC"
		
		RSObj.Open SqlItem, ConnItem, 1, 1
		
		If Not RSObj.EOF Then
					totalPut = RSObj.RecordCount
		
							If CurrentPage < 1 Then	CurrentPage = 1
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
			Else
		       Response.Write "<tr><td class='splittd' colspan=7 height='22' align='center'>没有记录!</td></tr>"
			End If
		Response.Write "</table>"
		Response.Write "</div>"
		Response.Write "</body>"
		Response.Write "</html>"
		End Sub
		
		Sub ExecuteAction()
		Dim Action, ID, FoundErr, ErrMsg, SqlItem
		
		Action = Trim(KS.G("Action"))
		ID = KS.FilterIds(KS.G("ID"))

		
		If Action = "verific" Then
		   If ID = "" Then
			  FoundErr = True
			  ErrMsg = ErrMsg & "<br><li>请选择要审核的记录</li>"
		   Else
			  SqlItem = "update " & KS.C_S(ChannelID,2) & " set Verific=1 Where ID in(" & ID & ")"
		   End If
		ElseIf Action = "verificall" Then
			ID = 1
		   SqlItem = "update " & KS.C_S(ChannelID,2) & " set Verific=1"
		ElseIf Action = "del" Then
		   SqlItem = "Delete From " & KS.C_S(ChannelID,2) & " Where ID in(" & ID & ")"
		ElseIf Action = "delall" Then
			ID = 1
		   SqlItem = "Delete From " & KS.C_S(ChannelID,2)
		End If
		If FoundErr <> True And ID <> "" Then
		   ConnItem.Execute (SqlItem)
		End If
		End Sub
		Sub showContent()
		   i = 0
		   Response.Write "<form name='myform' method='Post' action='?Page=" & CurrentPage & "&channelid=" & channelid & "'>"
		 Do While Not RSObj.EOF
			Response.Write ("<tr>")
			 Response.Write (" <td class='splittd' width=""435"" height=""18"">          ")
				Response.Write "<input type='checkbox' name='ID' value='" &RSObj("ID") & "'><span ondblclick='ViewRecords(this.NewsID);' NewsID='" & RSObj("ID") & "'><img src='../Images/folder/TheSmallWordNews1.gif'  align='absmiddle'>"
				  Response.Write "  <span style='cursor:default;'>" & KS.GotTopic(RSObj("Title"), 38) & "</span></span>"
			  Response.Write ("</td> ")
			  
		 Response.Write "     <td class='splittd' width=""142"" align=""center""> " 
		 if channelid=5 then
		 response.write RSObj("ProducerName") &"&nbsp;"
		 else
		 Response.Write KS.GotTopic(RSObj("Origin"), 15) 
		 End If
		 Response.Write " </td>"
		 Response.Write "     <td class='splittd' width=""110"" align=""center"">" & KMCObj.Collect_ShowClass_Name(Channelid, RSObj("TID")) & "</td>"
		 Response.Write "     <td class='splittd' align=""center""> "
				   Response.Write RSObj("adddate")
		Response.Write "     </td>"
		 Response.Write "     <td class='splittd' width=""132"" align=""center""><a href=""Collect_IntoDataBase.asp?ChannelID=" & ChannelID & "&Action=Into&page=" & currentpage & "&inflag=1&id=" & rsobj("id") & """>入库</a> <a href=""Collect_IntoDataBase.asp?ChannelID=" & ChannelID & "&Action=del&page=" & currentpage & "&id=" & rsobj("id") & """>删除</a></td>"
		Response.Write "    </tr>"
				   i = i + 1
				   If i > MaxPerPage Then
					  Exit Do
				   End If
				RSObj.MoveNext
		   Loop
			   
		RSObj.Close
		Set RSObj = Nothing
		   Response.Write "<tr><td colspan=7 height='25'><input name='chkAll' type='checkbox' id='chkAll' onclick=CheckAll(this.form) value='checkbox'>批量选择 <input type='submit' value='批量删除' class='button' onclick=""this.form.action='Collect_IntoDataBase.asp?ChannelID=" & ChannelID & "&Action=del&page=" & currentpage&"'"">&nbsp;<input type='submit' onclick=""this.form.action='?Action=Into&page=" & currentpage& "&channelid=" & channelid & "&inflag=1';"" value='批量入库' class='button'>&nbsp;<input type='button' onclick=""AllIntoDataBase();"" value='全部入库' class='button'>&nbsp;<input type='button' onclick=""DelAllRecords();"" value='全部删除' class='button'></td></tr>"
		   Response.Write "</form>"
		Response.Write ("<tr> ")
		Response.Write ("      <td height=""22"" align=""right"" colspan=""5"">")
		 Call KS.ShowPageParamter(totalPut, MaxPerPage, "Collect_IntoDataBase.asp", True, KS.C_S(ChannelID,4), CurrentPage, "ChannelID=" & ChannelID)
		Response.Write ("</td></tr>")
		%>
		<form name="sform" action="?" method="post">
		<tr>
		<td colspan="6">
		 <div style="border:1px dashed #cccccc;margin:3px;padding:4px">
		  <strong>快速搜索=></strong>  关键字:<input type="text" name="key"/> <input class="button" type="submit" value="开始搜索"/>
		 </div>
		</td>
		</tr>
		</form>
		<%
		End Sub
		
		'预览
		Sub RecordView()
		Dim SqlItem, RsItem
		'on error resume next
		If KS.G("id") = "" Then Response.End
		Set RsItem = Server.CreateObject("adodb.recordset")
		SqlItem = "select * from " & KS.C_S(ChannelID,2) & " where ID=" & KS.ChkClng(request("id"))
		RsItem.Open SqlItem, ConnItem, 1, 1
		If Not RsItem.EOF Then
		Response.Write "<center><strong>"
		Response.Write RsItem("title") & "</strong></center><br><div align='center'>" 
		If ChannelID<>5 Then
		 Response.Write "作者:" & RsItem("author") & "  来源:" & RsItem("origin")
		End If 
		 Response.Write " 录入:" & RsItem("Inputer") & " 更新时间:" & RsItem("Adddate") & "</div> <hr>"
		If KS.C_S(ChannelID,6)="1" Then
		 Response.Write RsItem("ArticleContent") 
		ElseIf KS.C_S(ChannelID,6)="2" Then
		 Response.Write "<div align='center'><img src='" & split(RsItem("PicURLS"),"|")(1) & "'><br/>" & split(RsItem("PicURLS"),"|")(0) & "</div>"
		ElseIf KS.C_S(ChannelID,6)="5" Then
		 Response.Write "<div>"
		 if Not KS.IsNul(rsItem("bigPhoto")) Then
		  response.write "大图:<a href='" & rsItem("bigphoto") & "' target='_blank'><img src='" & rsItem("bigphoto") & "' width='130'></a><br/>"
		 End If
		 response.write "产品规格:" & rsItem("ProSpecificat") & "<br/>"
		 response.write "生 产 商:" & rsItem("ProducerName") & "<br/>"
		 response.write "产品商标:" & rsItem("TrademarkName") & "<br/>"
		 response.write "计量单位:" & rsItem("Unit") & "<br/>"
		 response.write "市场价格:" & rsItem("price_market") & "<br/>"
		 response.write "商城价格:" & rsItem("price") & "<br/>"
		 response.write "会员价格:" & rsItem("price_member") & "<br/>"
		 response.write "介绍:" & RsItem("prointro") & "</div>"
		End If
		Response.Write "<br><div align='center'><input type='button' value='关闭窗口' onclick='window.close()'></div>"
		Else
		  Response.Write "参数传递出错!"
		End If
		RsItem.Close
		Set RsItem = Nothing
		End Sub

		
		'审核入库
		Sub IntoDataBase()
		Dim ID, SqlStr, Page, FRS
		ErrNum = 0
		SuccNum = 0
		ID = KS.FilterIds(Replace(KS.G("ID")," ",""))
		Page = KS.G("Page")
		Set FRS = Server.CreateObject("ADODB.RECORDSET")
		 If KS.G("Inflag") = 1 Then
			If ID="" Then
			 KS.AlertHintScript "请选择入库的记录!"
			End If
		   SqlStr = "Select * From " & KS.C_S(ChannelID,2) & " Where ID in (" & ID & ")  Order BY ID"
		  Else
		   SqlStr = "Select * From " & KS.C_S(ChannelID,2) & " Order BY ID"
		 End If
		FRS.Open SqlStr, ConnItem, 1, 3
		If Not FRS.EOF Then
		   Do While Not FRS.EOF
			Call InsertIntoBase(FRS)  '调用插入主数据库函数
			FRS.Delete
			FRS.MoveNext
		   Loop
		End If
		FRS.Close
		Set FRS = Nothing
		Response.Write ("<script>alert('提示:本次共操作 " & SuccNum + ErrNum & " " & KS.C_S(ChannelID,4) & KS.C_S(ChannelID,3) & "\n其中成功入库 " & SuccNum & " " & KS.C_S(ChannelID,4) &",重复而不允许入库 " & ErrNum & " " & KS.C_S(ChannelID,4) &";');location.href='Collect_IntoDataBase.asp?ChannelID=" & KS.G("ChannelID") & "&Page=" & Page & "';</script>")
		End Sub
		'执行插入主数据库(文章表)
		Sub InsertIntoBase(FRS)
		 ' on error resume next
		  Dim Rs, SqlStr, Tid,Intro,Images
		  Tid = FRS("Tid")
		  
		  Set Rs = Server.CreateObject("ADODB.RECORDSET")
		  If RepeatInto<>"1" Then
		  SqlStr = "Select * From " & KS.C_S(ChannelID,2) & " Where Title='" & FRS("Title") & "' And Tid='" & Tid & "'"
		  Else
		  SqlStr = "Select * From " & KS.C_S(ChannelID,2) & " where 1=0"
		  End If
		  Rs.Open SqlStr, conn, 1, 3
		  If Rs.EOF Then
		   Rs.AddNew
		   Rs("Tid") = Tid
		   Rs("Keywords") = FRS("Keywords")
		   Rs("Title") = FRS("Title")
		   
		  Select Case  KS.C_S(ChannelID,6)
		   Case 1
            Rs("ShowComment") = FRS("ShowComment")
		    Rs("TitleType") = FRS("TitleType")
		    Rs("TitleFontColor") = FRS("TitleFontColor")
		    Rs("TitleFontType") = FRS("TitleFontType")
		    Rs("ArticleContent") = FRS("ArticleContent")
		    Rs("Intro") = FRS("Intro")
			Rs("PicNews") = FRS("PicNews")
			Rs("Changes") = FRS("Changes")
			Intro=FRS("Intro")
		    Rs("Author") = FRS("Author")
		    Rs("Origin") = FRS("Origin")
			Images=FRS("ArticleContent")
		   Case 2
		    Intro=FRS("PictureContent")
		    Rs("PictureContent")=Intro
			Rs("PicUrls")=FRS("PicUrls")
		    Rs("Author") = FRS("Author")
		    Rs("Origin") = FRS("Origin")
			Rs("ShowStyle")=1
			Rs("PageNum")=12
			Images=FRS("PicUrls")&FRS("PictureContent")
		  Case 5
		  	Rs("ProID")=FRs("ProID")
			Rs("ProIntro")= Frs("ProIntro")
			Rs("BigPhoto")=FRs("BigPhoto")
			Rs("Unit")=FRs("Unit")
			Rs("Price_Member")=FRs("Price_Member")
			Rs("Price_Market")=FRs("Price_Market")
			Rs("Price_Original")=FRs("Price_Original")
			Rs("Price")=FRs("Price")
			Rs("ProModel")=FRs("ProModel")
			Rs("ProSpecificat")=FRs("ProSpecificat")
			Rs("ProducerName")=FRs("ProducerName")
			Rs("TrademarkName")=FRs("TrademarkName")
			Images=FRS("ProIntro")&FRS("BigPhoto")
          End Select
		   
		   Rs("Rank") = FRS("Rank")        '阅读星级
		   Rs("Hits") = FRS("Hits")
		   Rs("AddDate") = FRS("AddDate")   '更新时间
		   Rs("JSID") = FRS("JSID")
		   Rs("TemplateID") = FRS("TemplateID") '模板
		   Rs("Fname") = FRS("Fname")
		   rs("PhotoUrl")=FRS("PhotoUrl")
		   Rs("Inputer") = FRS("Inputer")
		   Rs("Recommend") = FRS("Recommend")
		   Rs("Rolls") = FRS("Rolls")
		   Rs("strip") = FRS("strip")
		   Rs("Popular") = FRS("Popular")
		   Rs("Verific") = FRS("Verific")     '审核与否
		   Rs("Slide") = FRS("Slide")
		   Rs("Comment") = FRS("Comment")
		   Images=Images&FRS("PhotoUrl")
		   If IsArray(Arr_Field) Then
		    For I=0 To Ubound(Arr_Field,2)
			 RS(Arr_Field(0,I))=FRS(Arr_Field(0,I))
			 Images=Images&FRS(Arr_Field(0,I))
			Next
		   End If
		   Rs.Update
		   SuccNum = SuccNum + 1
		   
		   RS.MoveLast
		   Call KS.FileAssociation(ChannelID,RS("ID"),Images ,0)
		   '向主表插入记录
		   Call LFCls.AddItemInfo(ChannelID,RS("ID"),RS("Title"),RS("Tid"),Intro,RS("KeyWords"),Rs("PhotoUrl"),RS("AddDate"),KS.C("AdminName"),rs("Hits"),rs("HitsByDay"),rs("HitsByWeek"),rs("HitsByMonth"),rs("Recommend"),rs("Rolls"),rs("Strip"),rs("Popular"),rs("Slide"),rs("IsTop"),rs("Comment"),rs("Verific"),RS("Fname"))

		  Else
		   ErrNum = ErrNum + 1
		  End If
		   Rs.Close
		   Set Rs = Nothing
		End Sub
End Class
%> 

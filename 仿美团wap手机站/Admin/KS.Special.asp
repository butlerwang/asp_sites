<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Special_Main
KSCls.Kesion()
Set KSCls = Nothing

Class Special_Main
        Private KS,KSCls
		Private SpecialID, i, totalPut, CurrentPage, SqlStr, SpecialRS
		Private FolderSql, FolderRS, ArticleTid, SpecialName
		Private CreateDate, TempStr,IcoUrl
		Private ChannelID,ClassID
		Private KeyWord, SearchType, StartDate, EndDate
		  '搜索参数集合
		Dim SearchParam
		  
		Private MaxPerPage
		Private Sub Class_Initialize()
		  MaxPerPage = 20
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub
		Public Sub Kesion()
		
		If Not KS.ReturnPowerResult(0, "M010003") Then                  '权限检查
				Call KS.ReturnErr(1, "")   
				Response.End()
		End iF

		KeyWord     = KS.G("KeyWord")
		SearchType  = KS.G("SearchType")
		StartDate   = KS.G("StartDate")
		EndDate     = KS.G("EndDate")
		SearchParam = "KeyWord=" & KeyWord & "&SearchType=" & SearchType & "&StartDate=" & StartDate & "&EndDate=" & EndDate
		ClassID    = KS.G("ClassID"):If ClassID = "" Then ClassID = "0"
		SpecialID   = KS.G("SpecialID"):If SpecialID = "" Then SpecialID = "0"
		
		  
		 Select Case KS.G("Action")
		 Case "SpecialList" GetTop : Call SpecialMainList()
		 Case "Add","Edit"  GetTop : Call SpecialAddOrEdit()
		 Case "AddSave" GetTop : Call SpecialAddSave()
		 Case "EditSave"  GetTop : Call SpecialEditSave()
		 Case "SpecialDel" GetTop : Call SpecialDel()
		 Case "SpecialInfoDel" GetTop : Call SpecialInfoDel()
		 Case "AddClass","EditClass" GetTop : Call SpecialClassAdd()
		 Case "DoClassSave" GetTop : Call DoClassSave()
		 Case "DelClass" GetTop : Call DelSpecialClass()
		 Case "ShowInfo" GetTop :  Call ShowInfo()
		 Case "SpecialClassList" GetTop : Call SpecialClassList()
		 Case "Select" Call SpecialSelect()
		 Case ELSE  GetTop : Call SpecialMainList()
		 End Select
		End Sub
		
		Sub SpecialSelect()
		  ChannelID=KS.S("channelid")
     %>
	<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<META HTTP-EQUIV="pragma" CONTENT="no-cache"> 
	<META HTTP-EQUIV="Cache-Control" CONTENT="no-cache, must-revalidate"> 
	<META HTTP-EQUIV="expires" CONTENT="Wed, 26 Feb 1997 08:21:57 GMT">
	<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>
	<title>选择专题</title>
    <script language="javascript" src="../KS_Inc/common.js"></script>
    <script language="javascript" src="../KS_Inc/jquery.js"></script>
	 <style type="text/css">
	  body{margin:0px;padding:0px;font-size:12px;COLOR: #454545; text-decoration: none;}
	  td{font0-size:12px;}
	  a{text-decoration: none;COLOR: #454545; }
	 </style>
	    <script language="javascript">
		function SelectFolder(TypeID){
		   $("#sub"+TypeID).toggle();
		   $("#sub"+TypeID).html("<img src='images/loading.gif'>");
		   $.get("../plus/ajaxs.asp",{action:"SpecialSubList",classid:TypeID},function(d){
		    $("#sub"+TypeID).html(unescape(d));
		   });
		}
		
      function set(specialid,specialname)
	  { 
	    top.frames["MainFrame"].UpdateSpecial(specialid+'@@@'+specialname);
		top.frames["MainFrame"].closeWindow();
	  }
    </script>
	  <body bgcolor="E9F6FE">
	    <table border="0" cellpadding="0" cellspacing="0" width="100%">
		 <tr>
		  <td>
	   <%
	  With KS 
		 Dim Node,K,SQL,ID,RS,Xml
		 Set RS=Conn.Execute("select ClassID,ClassName from KS_SpecialClass Order By OrderID ASC")
		 If Not RS.Eof Then
		   Set Xml=.RsToXml(RS,"row","xmlroot")
		 End If
		   RS.Close
		   Set RS=Nothing
		 If IsOBject(Xml) Then
		    For Each Node In Xml.DocumentElement.SelectNodes("row")
				ID=Node.SelectSingleNode("@classid").text
		          .echo "<table style=""margin:14px"" width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
				  .echo "<tr>" & vbcrlf
				  .echo " <td><img src='images/folder/folder.gif' align='absmiddle'><span onClick='SelectFolder(" & ID &");return false;'><a href='#'><strong>" & Node.SelectSingleNode("@classname").text & "</strong></a></span>"
				  .echo "</td>"&vbcrlf
				  .echo "</tr>" & vbcrlf
				  .echo "<tr>" & vbcrlf
                  .echo " <td style=""padding-left:20px"" ID=""sub"& ID &""" style=""display:none"">" & vbcrlf
                  .echo " </td>" & vbcrlf
                  .echo " </tr>" & vbcrlf
	  			  .echo "</table>"
			Next
	   Else
		     .echo "请先添加专题!"
	   End If
	   End With
	   		%>
		  </td>
		  </tr>
		 </table>
		</body>
		</html>
		<%
		End Sub
		
		Sub SpecialClassList()
			With KS
			.echo "</ul><table width='100%' border='0' cellspacing='0' cellpadding='0'>"
			.echo "        <tr>"
			.echo "          <td class=""sort"" width='65' align='center'>分类ID</td>"
			.echo "          <td class='sort' align='center'>专题分类名称</td>"
			.echo "          <td width='19%' class='sort' align='center'>专题数</td>"
			.echo "          <td width='10%' align='center' class='sort'>排序号</td>"
			.echo "          <td width='35%' align='center' class='sort'>管理操作</td>"
			.echo "  </tr>"
			MaxPerPage=15
			  Dim RS:Set RS = Server.CreateObject("ADODB.RecordSet")
			   RS.Open "SELECT ClassID,ClassName,OrderID FROM [KS_SpecialClass] order by OrderID", conn, 1, 1
				If Not RS.EOF Then
						totalPut = RS.RecordCount
						If CurrentPage < 1 Then CurrentPage = 1
						If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
							RS.Move (CurrentPage - 1) * MaxPerPage
						End If
						Dim SQL:SQL=RS.GetRows(MaxPerPage)
						Call showSpecialClass(SQL)
				Else
				  .echo "<tr><td class='splittd' align='center' height='25' colspan=5>您还没有添加专题分类，请添加!</td></tr>"
				End If
			.echo "</table>"
			.echo ("<table border='0' width='100%' cellspacing='0' cellpadding='0' align='center'>")
	        .echo ("<tr><td width='180'> </div>")
	        .echo ("</td>")
	        .echo ("<td></td>")
	        .echo ("</form><td align='right'>")
	         Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
	        .echo ("</td></tr></table>")
			.echo "</div>"
          End With
		End Sub
		
		Sub GetTop
			CurrentPage = KS.ChkClng(Request("page"))
			If CurrentPage=0 Then  CurrentPage = 1
		  With KS
			.echo "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd"">"
			.echo "<html xmlns=""http://www.w3.org/1999/xhtml"">"
			.echo "<head>"
			.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			.echo "<title>专题中心</title>"
			.echo "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.echo "<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>"
			.echo "<script language=""JavaScript"" src=""../KS_Inc/Jquery.js""></script>"
			.echo "<script language=""JavaScript"">" & vbCrLf
			.echo "var Page='" & CurrentPage & "';        //当前页码" & vbCrLf
			.echo "var ClassID='" & ClassID & "';       //频道ID" & vbCrLf
			.echo "var SpecialID=" & SpecialID & ";       //专题ID" & vbCrLf
			.echo "var KeyWord='" & KeyWord & "';         //搜索关键字" & vbCrLf
			.echo "var SearchParam='" & SearchParam & "'; //搜索参数集合" & vbCrLf
			.echo "</script>" & vbCrLf
			%>
			<script language="javascript">
			$(document).ready(function(){
				 // $(parent.frames["BottomFrame"].document).find("#Button1").attr("disabled",true);
				 // $(parent.frames["BottomFrame"].document).find("#Button2").attr("disabled",true);
				});
				function CreateHtml(SpecialID)
				{   if (SpecialID=='') SpecialID=get_Ids(document.myform);
					if (SpecialID!='')
					{
					   new parent.KesionPopup().PopupCenterIframe('发布专题','include/RefreshspecialSave.asp?Types=Special&id='+SpecialID+'&RefreshFlag=ID',530,110,'no')
					}
					else 
					alert('请选择要发布的专题!');
				}
				
				function ChangeUp()
				{
				 location.href='KS.Special.asp';
				 $(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?Opstr='+escape("所有专题")+'&ButtonSymbol=Disabled&ClassID='+ClassID;
				}
				
		
				function View(SpecialName,SpecialID)
				{ if (SpecialID=='') SpecialID=get_Ids(document.myform);
					 if (SpecialID!=''){
						 if (SpecialID.indexOf(',')==-1){
						 new parent.KesionPopup().PopupCenterIframe('查看专题<font color=red>['+SpecialName+']</font>下的文档','KS.Special.asp?Action=ShowInfo&SpecialID='+SpecialID+'',750,430,'auto')
							 } else alert('一次只能够编辑一个专题');
						}
					else{
					alert('请选择要编辑的专题');
					}
				}
				function Delete(SpecialID)
				{  
				  if (SpecialID=='') SpecialID=get_Ids(document.myform);
						if (SpecialID!='')
						{ 
						if (confirm('确定删除选中的专题吗?'))location="KS.Special.asp?Action=SpecialDel&Page="+Page+"&"+SearchParam+"&SpecialID="+SpecialID+'&ClassID='+ClassID;
						}
						else alert('请选择要删除的专题!');
					
				}
				function SpecialInfoDel(ID)
				{
				 if (confirm('确定将选中的文档从专题中移除吗?')) 
				 {
				   $("input[type=checkbox][value="+ID+"]").attr("checked",true);
				   $("#myform").submit();
				 }
				}
			  function showInfo(channelid,id)
			  {
				 window.open('../item/show.asp?m='+channelid+'&d='+id);
			   }
			function CreateSpecialClass()
			{
			 location.href='KS.Special.asp?Action=AddClass';
			 $(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?Opstr='+escape("专题管理 >> <font color=red>添加专题分类</font>")+'&ButtonSymbol=Go';
			}
			function EditSpecialClass(classid)
			{
			 location.href='KS.Special.asp?Action=EditClass&ClassID='+classid+'&Page='+Page;
			 $(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?Opstr='+escape("专题管理 >> <font color=red>修改专题分类</font>")+'&ButtonSymbol=GoSave';
			}
			function AddSpecial(ClassID)
			{
			 location.href='KS.Special.asp?Action=Add&ClassID='+ClassID;
			 $(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?Opstr='+escape("专题管理 >> <font color=red>添加专题</font>")+'&ButtonSymbol=Go';
			}
			function Edit(SpecialID)
			{  
			 if (SpecialID=='') SpecialID=get_Ids(document.myform);
			 if (SpecialID!=''){
				 if (SpecialID.indexOf(',')==-1){
				   location.href='KS.Special.asp?Action=Edit&SpecialID='+SpecialID;
				   $(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?Opstr='+escape("专题管理 >> <font color=red>编辑专题</font>")+'&ButtonSymbol=GoSave&ClassID='+ClassID;
					 } else alert('一次只能够编辑一个专题');
				}
			else{
			alert('请选择要编辑的专题');
			}
			
			}
			function ClassToggle(f)
			{
			  setCookie("SpecialclassExtStatus",f)
			  $('#classNav').toggle('slow');
			  $('#classOpen').toggle('show');
			}
			</script>
			<body>
			<%
		  If KS.G("Action")<>"ShowInfo" And KS.G("Action")<>"AddClass" And KS.G("Action")<>"EditClass" And KS.G("Action")<>"Add" And KS.G("Action")<>"Edit" Then
		 	.echo "<ul id='menu_top'>"
			.echo "<li class='parent' onclick=""location.href='?action=SpecialClassList'""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/addjs.gif' border='0' align='absmiddle'>分类管理</span></li>"
			.echo "<li class='parent' onclick='javascript:CreateSpecialClass();'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/addfolder.gif' border='0' align='absmiddle'>添加分类</span></li>"

			
			.echo "<li class='parent' onclick='javascript:AddSpecial(" & KS.G("ClassID") & ");'"
			If SpecialID <> "0" Then .echo (" Disabled=true")
			.echo "><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/addfolder.gif' border='0' align='absmiddle'>添加专题</span></li>"
			.echo "<li class='parent' onClick=""parent.initializeSearch('专题中心',0,'Special');""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/s.gif' border='0' align='absmiddle'>搜索助理</span></li>"
			.echo "<li class='parent' onClick=""ChangeUp();"""
			If ClassID="0" Then .echo " Disabled"
			.echo "><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/back.gif' border='0' align='absmiddle'>回上一级</span></li>"
		   End If
          End With
		End Sub
		
		Sub showSpecialClass(SQL)
		  Dim K
		  With KS
		  For K=0 To Ubound(SQL,2)
		     .echo "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
		     .echo "<td class='splittd' height='30' align='center'>" & SQL(0,K) & "</td>"
			 .echo "<td class='splittd'><img src='images/folder/folder.gif' align='absmiddle'><a href='?Action=SpecialList&ClassID=" & SQL(0,K) & "'>" & SQL(1,K) & "</a></td>"
			 .echo "<td class='splittd' align='center'>" & conn.execute("select count(*) from ks_special where classid=" & SQL(0,K))(0) & "</td>"
			 .echo "<td class='splittd' align='center'>" & SQL(2,K) &"</td>"
			 .echo "<td class='splittd' align='center'><a href='javascript:AddSpecial(" & SQL(0,K) & ");'>添加专题</a> | <a href='?Action=SpecialList&ClassID=" & SQL(0,K) & "'>查看该分类下的专题</a> | <a href='javascript:EditSpecialClass(" & SQL(0,K) & ");'>修改</a> | <a href='?Action=DelClass&ClassID=" &SQL(0,K) & "' onclick=""return(confirm('删除分类将同时删除该分类下的所有专题，确定删除吗？'))"">删除</a></td>"
			 .echo "</tr>"
		  Next
		  End With
		End Sub
		
		
		
		Sub SpecialMainList()
			With KS
			.echo "</head>"
		 If KeyWord = "" Then
			GetChannelList()
		 Else
			.echo "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" class=""sortbutton"">"
			.echo "  <tr>"
			.echo "    <td height=""23"" align=""left"">"
					   .echo ("<img src='Images/home.gif' align='absmiddle'><span style='cursor:pointer' onclick=""SendFrameInfo('KS.Special.asp','Special_Left.asp','KS.Split.asp?ButtonSymbol=Disabled&OpStr=专题管理 >> <font color=red>管理首页</font>')"">专题首页</span>")
				   .echo (">>> 搜索结果: ")
					 If StartDate <> "" And EndDate <> "" Then
						.echo ("专题更新日期在 <font color=red>" & StartDate & "</font> 至 <font color=red> " & EndDate & "</font>&nbsp;&nbsp;&nbsp;&nbsp;")
					 End If
					Select Case SearchType
					 Case 0
					  .echo ("名称含有 <font color=red>" & KeyWord & "</font> 的专题")
					 Case 1
					  .echo ("简要说明中含有 <font color=red>" & KeyWord & "</font> 的专题")
					 End Select
		 End If
				  
			.echo "    </td>"
			.echo "  </tr>"
			.echo "</table>"
			
			 '============分类显示,带记忆功能=======================================
			 Dim ExtStatus,CloseDisplayStr,ShowDisplayStr,classExtStatus
			 classExtStatus=request.cookies("SpecialclassExtStatus")
			 if classExtStatus="" Then classExtStatus=1
			 If classExtStatus=1 Then 
			  ExtStatus=2 :CloseDisplayStr="display:none;":ShowDisplayStr=""
			 Else 
			  ExtStatus=1 :CloseDisplayStr="":ShowDisplayStr="display:none;"
			 End If

			Dim RS,ClassXML,Node
			Set RS=Conn.Execute("Select ClassID,ClassName From KS_SpecialClass Order by OrderID")
			If Not RS.Eof Then Set ClassXML=KS.RsToXml(RS,"row","classxml")
			RS.Close:Set RS=Nothing
			If IsObject(ClassXML) Then
			.echo "<div id='classOpen' onclick=""ClassToggle("& ExtStatus& ")"" style='" & CloseDisplayStr &"cursor:pointer;text-align:center;position:absolute; z-index:2; left: 0px; top: 38px;' ><img src='images/kszk.gif' align='absmiddle'></div>"
		    .echo "<div id='classNav' style='" & ShowDisplayStr &"position:relative;height:auto;_height:30px;top:4px;line-height:30px;margin:8px 1px;border:1px solid #DEEFFA;background:#F7FBFE'>"
		    .echo "<div style='padding-top:2px;cursor:pointer;text-align:center;position:absolute; z-index:1; right: 0px; top: 2px;'  onclick=""ClassToggle(" & ExtStatus &")""> <img src='images/close.gif' align='absmiddle'></div>"
			 For Each Node In ClassXML.DocumentElement.SelectNodes("row")
			   .echo "<li style='margin:5px;float:left;width:100px'><img src='images/folder/folderopen.gif' align='absmiddle'><a href='?classid=" & Node.SelectSingleNode("@classid").text & "' title='" & Node.SelectSingleNode("@classid").text & "'>" & KS.Gottopic(Node.SelectSingleNode("@classname").text,10) & "</a></li>"
			 Next
			 .echo "</div>"
			End If
			 '=============================================================

			
			
		    .echo ("<table width=""100%"" align='center' border=""0"" cellpadding=""0"" cellspacing=""0"">")
			.echo ("<form name='myform' id='myform' action='KS.Special.asp' method='post'>")
		    .echo ("<tr class='sort'>")
			.echo ("<td>选择</td><td>专题名称</td><td>分类</td><td>添加时间</td><td>管理操作</td>")
			.echo ("</tr>")
	
	
	  If KeyWord <> "" Then
		  Dim Param:Param = " Where 1=1"
		  Select Case SearchType
			Case 0
			Param = Param & " And SpecialName like '%" & KeyWord & "%'"
			Case 1
			Param = Param & " And SpecialNote like '%" & KeyWord & "%'"
		  End Select
			If StartDate <> "" And EndDate <> "" Then
				Param = Param & " And (SpecialAddDate>=#" & StartDate & "# And SpecialAddDate<=#" & DateAdd("d", 1, EndDate) & "#)"
		   End If
		  Param = Param & " Order BY SpecialAddDate desc"
		  SqlStr = "Select SpecialID,a.ClassID,b.ClassName,SpecialName,Creater,SpecialAddDate,SpecialNote from KS_Special a Inner Join KS_SpecialClass B on a.classid=b.classid " & Param
	  Else
		  If ClassID<>"0" Then
		   SqlStr = "Select SpecialID,a.ClassID,b.ClassName,SpecialName,Creater,SpecialAddDate,SpecialNote from KS_Special a Inner Join KS_SpecialClass B on a.classid=b.classid Where a.ClassID=" & ClassID & " Order BY SpecialAddDate desc"
		  Else
		   SqlStr = "Select SpecialID,a.ClassID,b.ClassName,SpecialName,Creater,SpecialAddDate,SpecialNote from KS_Special a Inner Join KS_SpecialClass B on a.classid=b.classid Order BY SpecialAddDate desc"
		  End If
	  End If
	 Set SpecialRS = Server.CreateObject("AdoDb.RecordSet")
	 SpecialRS.Open SqlStr, Conn, 1, 1
	 If SpecialRS.EOF Then
	    .echo "<tr><td class='splittd' onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"" colspan='5' align='center'>找不到专题!</td></tr>"
	 Else
				totalPut = SpecialRS.RecordCount
	
						If CurrentPage > 1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
								SpecialRS.Move (CurrentPage - 1) * MaxPerPage
						End If
						
						Dim XML:Set XML=KS.ArrayToXml(SpecialRS.GetRows(MaxPerPage),SpecialRS,"row","xmlroot")
						showSpecialList XML
						Set XML=Nothing
						
		End If
			 .echo " <tr>"
			 .echo " <td colspan='3'><div style='margin:5px'><b>选择：</b><a href='javascript:void(0)' onclick='Select(0)'>全选</a> -  <a href='javascript:void(0)' onclick='Select(1)'>反选</a> - <a href='javascript:void(0)' onclick='Select(2)'>不选</a> <input type='button' class='button' value='删 除' onclick=""Delete('');""> &nbsp;&nbsp;<input type='button' class='button' value='生 成' onclick=""CreateHtml('');""></td></form>"
			 .echo "   <td align=""right"" colspan=5>"
			 
			 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
			.echo "   </td>"
			.echo "  </tr>"
	.echo "</table>"
	.echo "</body>"
	.echo "</html>"
	 End With
	End Sub
	
	 Sub showSpecialList(XML)
	  Dim Node,SpecialID,SpecialName
	  If Not IsObject(XML) Then Exit Sub
	  With KS
			For Each Node In XML.DocumentElement.SelectNodes("row")
			    SpecialID=Node.SelectSingleNode("@specialid").text
				SpecialName=Node.SelectSingleNode("@specialname").text
				  .echo ("<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"" id='u" &SpecialID & "' onclick=""chk_iddiv('" & SpecialID & "')"">")
				  .echo ("<td class='splittd' align=center><input name='id'  onclick=""chk_iddiv('" & SpecialID & "')"" type='checkbox' id='c"& SpecialID & "' value='" & SpecialID & "'></td>")
				  .echo ("<td class='splittd' TITLE='名 称:" & SpecialName & "'>")
				  .echo ("<span onmousedown=""mousedown(this);"" style=""POSITION:relative;"" SpecialID=""" &SpecialID & """ SpecialName=""" & SpecialName & """>")
				  .echo ("<img src=""Images/Folder/Special.gif""> ")
				  .echo ("<a href=""javascript:View('" & SpecialName & "','" & SpecialID & "')"">" & SpecialName & "</a>")
				  .echo ("</td>")
				  .echo ("<td class='splittd' align='center'>" & Node.SelectSingleNode("@classname").text & "</td>")
				  .echo ("<td class='splittd' align='center'>" & Node.SelectSingleNode("@specialadddate").text & "</td>")
				  .echo ("<td class='splittd' align='center'><a href='javascript:Edit(""" & SpecialID & """);'>编辑</a> | <a href='javascript:Delete(" & SpecialID & ")'>删除</a> | <a href='javascript:CreateHtml(""" & SpecialID & """);'>生成</a> | <a href=""javascript:View('" & SpecialName & "','" & SpecialID & "')"">查看</a> | <a href='../item/special.asp?id=" &SpecialID & "' target='_blank'>浏览</a></td>")
			     .echo " </tr>"
			Next
					   
			End With
		End Sub
		
		'显示专题下的信息
		Sub ShowInfo()
		    MaxPerPage=10
		 	With KS
			 .echo ("<table width=""100%"" align='center' border=""0"" cellpadding=""0"" cellspacing=""0"">")
			 .echo ("<form name='myform' id='myform' action='KS.Special.asp' method='post'>")
			 .echo ("<input type='hidden' name='action' value='SpecialInfoDel'>")
		     .echo ("<tr class='sort'>")
			 .echo ("<td>选择</td><td>文档名称</td><td>分类</td><td>添加时间</td><td>管理操作</td>")
			 .echo ("</tr>")

			 Dim SQLStr
			 Dim RS:Set RS=Server.CreateoBject("ADODB.RECORDSET")
			 SQLStr="Select R.ID,I.ChannelID,I.InfoID,I.Title,I.Tid,I.AddDate From KS_ItemInfo I Inner Join KS_SpecialR R On I.InfoID=R.InfoID Where R.SpecialID=" & SpecialID & " and i.channelid=r.channelid Order by i.id Desc"
			 RS.Open SQLStr,Conn,1,1
			 If RS.EOF Then
			  .echo ("<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">")
			  .echo "<td class='splittd' colspan='6' align='center'>该专题下没有添加文档!</td>"
			  .echo "</tr>"
			 Else
					      totalPut = RS.RecordCount
		
							If CurrentPage > 1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
									RS.Move (CurrentPage - 1) * MaxPerPage
							End If
							
							Dim XML,Node,InfoID,RID
							Set XML=KS.ArrayToXml(RS.GetRows(MaxPerPage),RS,"row","xmlroot")
							If IsObject(XML) Then
								For Each Node In XML.DocumentElement.SelectNodes("row")
								      RID=Node.SelectSingleNode("@id").text
									  InfoID=Node.SelectSingleNode("@infoid").text
									  .echo ("<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"" id='u" &RID & "' onclick=""chk_iddiv('" & RID & "')"">")
									  .echo ("<td class='splittd' align=center><input name='id'  onclick=""chk_iddiv('" & RID & "')"" type='checkbox' id='c"& RID & "' value='" & RID & "'></td>")
									  .echo ("<td class='splittd' TITLE='名 称:" & Node.SelectSingleNode("@title").text & "'>")
									  .echo ("<a href='javascript:void(0)' onclick=""showInfo('" & Node.SelectSingleNode("@channelid").text & "','" & InfoID & "')"">" & KS.Gottopic(Node.SelectSingleNode("@title").text,30) & "</a>")
									  .echo ("</td>")
									  .echo ("<td class='splittd' align='center'>" & KS.C_C(Node.SelectSingleNode("@tid").text,1) & "</td>")
									  .echo ("<td class='splittd' align='center'>" & Node.SelectSingleNode("@adddate").text & "</td>")
									  .echo ("<td class='splittd' align='center'> <a href=""javascript:SpecialInfoDel('" & RID & "')"">删除</a> | <a href=""javascript:showInfo(" & Node.SelectSingleNode("@channelid").text & "," & InfoID & ")"">查看</a></td>")
									 .echo " </tr>"
								Next
							End If
							Set XML=Nothing
							
			End If
			RS.Close:Set RS=Nothing
			 .echo " <tr>"
			 .echo " <td colspan='3'><div style='margin:5px'><b>选择：</b><a href='javascript:void(0)' onclick='Select(0)'>全选</a> -  <a href='javascript:void(0)' onclick='Select(1)'>反选</a> - <a href='javascript:void(0)' onclick='Select(2)'>不选</a> <input type='submit' class='button' value='删 除' onclick=""return(confirm('确定移除选中的文档吗?'))""> </td></form>"
			 .echo "   <td align=""right"" colspan=5>"
				Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
			.echo "   </td>"
			.echo "  </tr>"
		  End With
		End Sub
		
		'添加专题分类
		Sub SpecialClassAdd()
		   Dim ClassID,Action,ClassName,ClassEName,TemplateID,FsoIndex,AddDate,Descript,TopTitle,PhotoUrl,OrderID
		   Dim CurrPath:CurrPath = KS.GetUpFilesDir()
			If KS.G("Action")="EditClass" Then
			  ClassID=KS.G("ClassID")
			  TopTitle="编辑"
			  Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
			  RSObj.Open "Select * From KS_SpecialClass Where ClassID=" & ClassID,Conn,1,1
			  If Not RSObj.Eof Then
				ClassName   = RSObj("ClassName")
				ClassEName  = RSObj("ClassEName")
				TemplateID    = RSObj("TemplateID")
				FsoIndex=RSObj("FsoIndex")
				AddDate       = RSObj("AddDate")
				Descript   = RSObj("Descript")
				OrderID = RSOBj("Orderid")
			  End If
			Else
			  TopTitle="添加":AddDate=Now:FsoIndex="Index.html"
			  OrderID=KS.ChkClng(conn.execute("select max(OrderID) from ks_specialclass")(0))+1
			End If
			With KS
			.echo "<div class='topdashed sort'>" & TopTitle &"专题分类</div>" & vbCrLf

			 
			.echo "  <table width='100%' border='0' align='center' clcass='border' cellpadding='0' cellspacing='0'>" & vbCrLf
			.echo "  <form action='KS.Special.asp?Action=DoClassSave' name='SpecialForm' method='post'>" & vbCrLf
			.echo "  <input name='ClassID' type='hidden' id='ClassID' value='" & ClassID &"'>" & vbCrLf
			.echo "  <input name='Page' type='hidden' value='" & KS.G("Page") &"'>" & vbCrLf
			.echo "    <tr>" & vbCrLf
			.echo "      <td>" & vbCrLf
            .echo "        <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class='ctable'>" & vbCrLf
			.echo "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">" & vbCrLf
			.echo "      <td width='179' height='35'class='clefttitle'> <div align='right'><strong>专题类别名称：</strong></div></td>" & vbCrLf
			.echo "      <td> <input name='ClassName' value='" & ClassName & "' type='text' size='30' class='textbox'>"
			.echo "              概况性说明文字 </td>"
			.echo "    </tr>" & vbCrLf
			.echo "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.echo "      <td height='35' class='clefttitle'> <div align='right'><strong>专题类别目录名称：</strong></div></td>" & vbCrLf
			.echo "      <td>"
			.echo "<input"
				If KS.G("Action")="EditClass" Then .echo " Disabled"
			.echo " name='ClassEName' type='text' value='" & ClassEName & "'  size='30' class='textbox'>"
			.echo "        只能是字母，数字或下划线的组合  </td>"
			.echo "    </tr>"
			.echo "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.echo "      <td height='35' class='clefttitle'> <div align='right'><strong>专题列表页模板：</strong></div></td>" & vbCrLf
			.echo "      <td><input type='text' size='30' name='TemplateID' id='TemplateID' value='" & TemplateID & "' class='textbox'>&nbsp;" & KSCls.Get_KS_T_C("$('#TemplateID')[0]") 
			.echo "      </td>" & vbCrLf
			.echo "    </tr>" & vbCrLf
				
			.echo "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">" & vbCrLf
			.echo "      <td height='35' class='clefttitle'> <div align='right'><strong>生成专题列表页的文件名</strong></td>" & vbCrLf
			.echo "      <td><select name='FsoIndex' class='textbox'>"
			.echo "          <option value='index.html' selected>index.html</option>"
			.echo "          <option value='index.htm'>index.htm</option>"
			.echo "          <option value='index.shtm'>index.shtm</option>"
			.echo "          <option value='index.shtml'>index.shtml</option>"
			.echo "          <option value='default.html'>default.html</option>"
			.echo "          <option value='default.htm'>default.htm</option>"
			.echo "          <option value='default.shtml'>default.shtml</option>"
			.echo "          <option value='default.shtm'>default.shtm</option>"
			.echo "          <option value='index.asp'>index.asp</option>"
			.echo "         <option value='" & FsoIndex & "' selected>" & FsoIndex & "</option>"
			.echo "        </select></td>"
			.echo "    </tr>" & vbCrLf
			
			.echo "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.echo "      <td & vbCrLfheight='35' class='clefttitle'> <div align='right'><strong>添加时间：</strong></div></td>"
			.echo "      <td><input name='AddDate' type='text' value='" & AddDate & "' size='30' readonly class='textbox'>"
			.echo "      </td>" & vbCrLf
			.echo "    </tr>" & vbCrLf
			.echo "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.echo "      <td height='35' class='clefttitle'> <div align='right'><strong>简要说明：</strong></div></td>"
			.echo "      <td><textarea name='Descript' rows='8' style='width:80%;border-style: solid; border-width: 1'>" &Descript & "</textarea></td>"
			.echo "    </tr>"
			.echo "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.echo "      <td height='35' class='clefttitle'> <div align='right'><strong>分类序号：</strong></div></td>"
			.echo "      <td><input name='OrderID' size='5' value='" & OrderID & "' Class='textbox' style='text-align:center'> 数字越小排在越前面</td>"
			.echo "    </tr>"
			.echo "  </table>"
			.echo "       </td>"
			.echo "    </tr>"
			.echo "    </table>"
			.echo "  </form>"
			.echo "</body>"
			.echo "</html>"
			.echo "<Script Language='javascript'>" & vbCrLf
			.echo "<!--" & vbCrLf
			.echo "function CheckForm()" & vbCrLf
			.echo "{ var form=document.SpecialForm;" & vbCrLf
			.echo "   if (form.ClassName.value=='')"
			.echo "    {" & vbCrLf
			.echo "     alert('请输入专题分类名称!');" & vbCrLf
			.echo "     form.ClassName.focus();" & vbCrLf
			.echo "     return false;" & vbCrLf
			.echo "    }" & vbCrLf
			.echo "    if (form.ClassEName.value=='')" & vbCrLf
			.echo "    {"
			.echo "     alert('请输入专题分类的英文名称!');" & vbCrLf
			.echo "     form.ClassEName.focus();" & vbCrLf
			.echo "    return false;" & vbCrLf
			.echo "    }"
			.echo "    if (form.TemplateID.value=='')" & vbCrLf
			.echo "    {"
			.echo "     alert('请绑定专题列表页模板!');" & vbCrLf
			.echo "     form.TemplateID.focus();" & vbCrLf
			.echo "    return false;" & vbCrLf
			.echo "    }"

			.echo "    if (CheckEnglishStr(form.ClassEName,'目录的英文名称')==false)" & vbCrLf
			.echo "     return false;" & vbCrLf
			.echo "    form.submit();" & vbCrLf
			.echo "    return true;" & vbCrLf
			.echo "}" & vbCrLf
			.echo "//-->" & vbCrLf
			.echo "</Script>"
		  End With
		End Sub
		
		Sub DoClassSave()
			Dim RS, Sql,ClassName, ClassEName, TemplateID, FsoIndex, AddDate, Descript,OrderID,ClassID
			ClassID    = KS.ChkClng(KS.G("ClassID"))
			ClassName  = KS.G("ClassName")
			ClassEName = KS.G("ClassEName")
			TemplateID = KS.G("TemplateID")
			FsoIndex   = KS.G("FsoIndex")
			AddDate    = KS.G("AddDate")
			Descript   = KS.G("Descript")
			OrderID    = KS.ChkClng(KS.G("OrderID"))
			With KS		 
				 If ClassName <> "" Then
					If Len(ClassName) >= 100 Then
						Call KS.AlertHistory("专题分类名称不能超过50个字符!", -1):Exit Sub
					End If
				 Else
					Call KS.AlertHistory("请输入专题分类名称!", -1):Exit Sub
				 End If
				 If ClassEName <> "" and  ClassID=0 Then
					If Len(ClassEName) >= 50 Then
						Call KS.AlertHistory("专题分类英文名称不能超过50个字符!", -1):Exit Sub
					End If
					If Not Conn.Execute("Select ClassEName,ClassName from KS_SpecialClass where ClassID<>" & ClassID & " and ClassName='" & ClassName & "'").eof  Then Call KS.alertHistory("数据库中已存在该专题分类名称!", -1)
					If Not Conn.Execute("Select ClassEName,ClassName from KS_SpecialClass where ClassID<>" &ClassID & " and ClassEName='" & ClassEName & "'").eof  Then Call KS.alertHistory("数据库中已存在该专题分类英文名称!", -1)
				 ElseIf ClassID=0 Then
					Call KS.alertHistory("请输入专题分类英文名称!", -1)
					.End
				 End If
				 If ClassID=0 Then
				  Conn.Execute("Insert Into KS_SpecialClass(ClassName,ClassEname,Descript,FsoIndex,AddDate,TemplateID,OrderID) Values('" & ClassName & "','" & ClassEname & "','" & Descript & "','" & FsoIndex & "','" & AddDate & "','" & TemplateID & "'," & OrderID &")")
				  .echo ("<script>if (confirm('添加专题分类成功,继续添加吗?')==true){location.href='KS.Special.asp?action=AddClass';}else{location.href='KS.Special.asp?action=SpecialClassList';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?Opstr=" & server.URLEncode("内容管理 >> 所有专题分类") & "&ButtonSymbol=Disabled&ClassID=" & ClassID & "';}</script>")     
				 Else
				  Conn.Execute("Update KS_SpecialClass Set ClassName='" & ClassName & "',Descript='" & Descript & "',FsoIndex='" & FsoIndex & "',AddDate='" & AddDate & "',TemplateID='" & TemplateID & "',OrderID=" & Orderid & " Where ClassID=" & ClassID)
				  .echo ("<script>alert('专题分类修改成功');location.href='KS.Special.asp?action=SpecialClassList&Page=" & KS.G("Page") &"';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?Opstr=" & server.URLEncode("内容管理 >> 所有专题分类") & "&ButtonSymbol=Disabled';</script>")     
				 End If
			End With
		End Sub
		
		'删除专题分类
		Sub DelSpecialClass()
		  Dim ClassID:ClassID=KS.ChkClng(KS.S("ClassID"))
		  Conn.Execute("Delete From KS_SpecialR Where SpecialID in(select specialid from ks_special where classid=" & ClassID & ")")
		  Conn.Execute("Delete From KS_Special Where ClassID=" & ClassID)
		  Conn.Execute("Delete From KS_SpecialClass Where ClassID=" & ClassID)
		  Response.Redirect Request.ServerVariables("HTTP_REFERER")
		End Sub
		
		'添加或编辑专题
		Sub SpecialAddOrEdit()
		   Dim SpecialID,Action,SpecialName,SpecialEName,TemplateID,FsoSpecialIndex,AddDate,SpecialNote,TopTitle,PhotoUrl,ClassID,MetaKey,MetaDescript
		   Dim CurrPath:CurrPath = KS.GetUpFilesDir()
			If KS.G("Action")="Edit" Then
			  SpecialID=KS.G("SpecialID")
			  Action="EditSave":TopTitle="编辑"
			  Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
			  RSObj.Open "Select top 1 * From KS_Special Where SpecialID=" & SpecialID,Conn,1,1
			  If Not RSObj.Eof Then
			    ClassID       = RSObj("ClassID")
				SpecialName   = RSObj("SpecialName")
				SpecialEName  = RSObj("SpecialEName")
				TemplateID    = RSObj("TemplateID")
				FsoSpecialIndex=RSObj("FsoSpecialIndex")
				AddDate       = RSObj("SpecialAddDate")
				SpecialNote   = RSObj("SpecialNote")
				PhotoUrl      = RSObj("PhotoUrl")
				MetaKey       = RSObj("MetaKey")
				MetaDescript  = RSObj("MetaDescript")
			  End If
			Else
			  ClassID=KS.G("ClassID"):TopTitle="添加":Action="AddSave":AddDate=Now:FsoSpecialIndex="Index.html"
			End If
			If KS.IsNul(SpecialNote) Then SpecialNote=" "
			With KS
			.echo "<div class='topdashed sort'>" & TopTitle &"专题</div>" & vbCrLf
			.echo "  <table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
			.echo "  <form action='KS.Special.asp?Action=" & Action & "' name='SpecialForm' method='post'>" & vbCrLf
			.echo "  <input name='SpecialID' type='hidden' id='SpecialID' value='" & SpecialID &"'>" & vbCrLf
			.echo "    <tr>" & vbCrLf
			.echo "      <td>" & vbCrLf
            .echo "        <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class='ctable'>" & vbCrLf
			.echo "       <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">" & vbCrLf
			.echo "        <td height='35' class='clefttitle'> <div align='right'><strong>所属专题分类：</strong></div></td>" & vbCrLf
			.echo "         <td width='542'>" & vbCrLf
			.echo "         <select name='ClassID' class='textbox'>" & vbCrLf
					  
					  Dim FolderName, TempStr, FolderRS
						Set FolderRS = Server.CreateObject("ADODB.Recordset")
						TempStr = "<option value=0>--请选择专题分类--</option>"
					  FolderRS.Open "Select ClassID,ClassName From KS_SpecialClass Order BY OrderID", Conn, 1, 1
					If Not FolderRS.EOF Then
					  Do While Not FolderRS.EOF
						 FolderName = Trim(FolderRS(1))
						 If trim(ClassID) = Trim(FolderRS(0)) Then
						   TempStr = TempStr & "<option value=" & FolderRS(0) & " Selected>" & FolderName & "</option>"
						 Else
						   TempStr = TempStr & "<option value=" & FolderRS(0) & ">" & FolderName & "</option>"
						 End If
						 FolderRS.MoveNext
					  Loop
					End If
					FolderRS.Close:Set FolderRS = Nothing
					.echo TempStr
					
			.echo "        </select>" & vbCrLf
			.echo "            </td>" & vbCrLf
			.echo "    </tr>" & vbCrLf
			.echo "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">" & vbCrLf
			.echo "      <td width='179' height='35'class='clefttitle'> <div align='right'><strong>专题名称：</strong></div></td>" & vbCrLf
			.echo "      <td> <input name='SpecialName' value='" & SpecialName & "' type='text' id='SpecialName' size='30' class='textbox'>"
			.echo "              概况性说明文字 </td>"
			.echo "    </tr>" & vbCrLf
			.echo "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.echo "      <td height='35' class='clefttitle'> <div align='right'><strong>专题目录：</strong></div></td>" & vbCrLf
			.echo "      <td>"
			.echo "<input"
				If KS.G("Action")="Edit" Then .echo " Disabled"
			.echo " name='SpecialEName' type='text' value='" & SpecialEName & "' id='SpecialEName' size='30' class='textbox'>"
			.echo "        不能带\/：*？“ < > | 等特殊符号,并且一旦设定就不能改，请慎重  </td>"
			.echo "    </tr>"
			.echo "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.echo "      <td height='35' class='clefttitle'> <div align='right'><strong>专题页模板：</strong></div></td>" & vbCrLf
			.echo "      <td><input type='text' size='30' name='TemplateID' id='TemplateID' value='" & TemplateID & "' class='textbox'>&nbsp;" & KSCls.Get_KS_T_C("$('#TemplateID')[0]") 
			.echo "      </td>" & vbCrLf
			.echo "    </tr>" & vbCrLf
				 .echo "          <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">" & vbCrLf
				.echo "           <td height='40' nowrap class='clefttitle'> <div align='right'><strong>专题图片地址：</strong></td>" & vbCrLf
				.echo "            <td height='28' nowrap>" & vbCrLf
				.echo "             <INPUT NAME='PhotoUrl' value='" & PhotoUrl &"' TYPE='text' id='PhotoUrl' class='textbox' size=30>"
				.echo "                  <input class=""button""  type='button' name='Submit' value='选择图片...' onClick=""OpenThenSetValue('Include/SelectPic.asp?CurrPath=" & CurrPath & "',550,290,window,document.SpecialForm.PhotoUrl);"">  <input class=""button"" type='button' name='Submit' value='远程抓取图片...' onClick=""OpenThenSetValue('Include/Frame.asp?FileName=SaveBeyondfile.asp&PageTitle='+escape('抓取远程图片')+'&ItemName=图片&CurrPath=" & CurrPath & "',300,100,window,document.SpecialForm.PhotoUrl);"">"
				.echo "              </td>" & vbCrLf
				.echo "          </tr>" & vbCrLf
			.echo "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">" & vbCrLf
			.echo "      <td height='35' class='clefttitle'> <div align='right'><strong>生成专题页的文件名</strong></td>" & vbCrLf
			.echo "      <td><select name='FsoSpecialIndex' class='textbox'>"
			.echo "          <option value='index.html' selected>index.html</option>"
			.echo "          <option value='index.htm'>index.htm</option>"
			.echo "          <option value='index.shtm'>index.shtm</option>"
			.echo "          <option value='index.shtml'>index.shtml</option>"
			.echo "          <option value='default.html'>default.html</option>"
			.echo "          <option value='default.htm'>default.htm</option>"
			.echo "          <option value='default.shtml'>default.shtml</option>"
			.echo "          <option value='default.shtm'>default.shtm</option>"
			.echo "          <option value='index.asp'>index.asp</option>"
			.echo "         <option value='" & FsoSpecialIndex & "' selected>" & FsoSpecialIndex & "</option>"
			.echo "        </select></td>"
			.echo "    </tr>" & vbCrLf
			
			.echo "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.echo "      <td & vbCrLfheight='35' class='clefttitle'> <div align='right'><strong>添加时间：</strong></div></td>"
			.echo "      <td><input name='SpecialAddDate' type='text' id='SpecialAddDate' value='" & AddDate & "' size='30' class='textbox'>"
			.echo "      </td>" & vbCrLf
			.echo "    </tr>" & vbCrLf
			.echo "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.echo "      <td height='35' class='clefttitle'> <div align='right'><strong>简要说明：</strong></div></td>"
			.echo "      <td><textarea name='SpecialNote' rows='8' id='SpecialNote' style='display:none;width:80%;border-style: solid; border-width: 1'>" & Server.HTMLEncode(SpecialNote) & "</textarea>"
			.echo "		<script type=""text/javascript"" src=""../editor/ckeditor.js"" mce_src=""../editor/ckeditor.js""></script>"
			.echo "		   <script type=""text/javascript"">"
			.echo "                CKEDITOR.replace('SpecialNote', {width:'98%',height:'180px',toolbar:'Basic',filebrowserBrowseUrl :'Include/SelectPic.asp?from=ckeditor&Currpath=" & KS.GetUpFilesDir() & "',filebrowserWindowWidth:650,filebrowserWindowHeight:290});"
			.echo "	</script>"
			
			.echo "    </tr>"
			.echo  " <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.echo  "           <td height='50' align='right' width='200' class='clefttitle'><strong>栏目META关键词：</strong><br>"
			.echo "      <font color='#0000FF'>用于设置针对搜索引擎的关键词<br>可在对应的栏目模板页使用标签<br><font color=red>""{$GetSpecialMetaKey}""</font> 进行调用</font></td>"
			.echo "         <td height='28'>"
            .echo " <textarea name='MetaKey' id='MetaKey' class='upfile' cols='70' rows='5'>" & MetaKey & "</textarea>             </td></tr>"
			.echo "      <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">           <td height='50' align='right' width='200' class='clefttitle'><strong>栏目META网页描述：</strong><br>"
			.echo "<font color='#0000FF'>用于设置针对搜索引擎的网页描述<br>可在对应的栏目模板页使用标签<br><font color=red>""{$GetSpecialMetaDescript}""</font> 进行调用</font></font></td>"
		    .echo " <td height='28'><textarea name='MetaDescript' id='MetaDescript' class='upfile' cols='70' rows='5'>" & MetaDescript & "</textarea>             </td></tr>"
			.echo "  </table>"
			.echo "       </td>"
			.echo "    </tr>"
			.echo "    </table>"
			.echo "  </form>"
			.echo "</body>"
			.echo "</html>"
			.echo "<Script Language='javascript'>" & vbCrLf
			.echo "<!--" & vbCrLf
			.echo "function CheckForm()" & vbCrLf
			.echo "{ var form=document.SpecialForm;" & vbCrLf
			.echo "    if (form.ClassID.value==0)" & vbCrLf
			.echo "    {"
			.echo "     alert('请选择所属专题分类!');" & vbCrLf
			.echo "    form.ClassID.focus();" & vbCrLf
			.echo "    return false;" & vbCrLf
			.echo "    }"
			.echo "   if (form.SpecialName.value=='')"
			.echo "    {" & vbCrLf
			.echo "     alert('请输入专题名称!');" & vbCrLf
			.echo "     form.SpecialName.focus();" & vbCrLf
			.echo "     return false;" & vbCrLf
			.echo "    }" & vbCrLf
			.echo "    if (form.SpecialEName.value=='')" & vbCrLf
			.echo "    {"
			.echo "     alert('请输入专题的英文名称!');" & vbCrLf
			.echo "     form.SpecialEName.focus();" & vbCrLf
			.echo "    return false;" & vbCrLf
			.echo "    }"
			.echo "    if (form.TemplateID.value=='')" & vbCrLf
			.echo "    {"
			.echo "     alert('请绑定专题模板!');" & vbCrLf
			.echo "     form.TemplateID.focus();" & vbCrLf
			.echo "    return false;" & vbCrLf
			.echo "    }"

			.echo "    if (CheckEnglishStr(form.SpecialEName,'目录的英文名称')==false)" & vbCrLf
			.echo "     return false;" & vbCrLf
			.echo "    form.submit();" & vbCrLf
			.echo "}" & vbCrLf
			.echo "//-->" & vbCrLf
			.echo "</Script>"
		  End With
		End Sub
		
		'保存添加
		Sub SpecialAddSave()
		  Dim TemplateRS, TemplateSql,TempObj, SpecialRS, SpecialSql,SpecialName, SpecialEName, TemplateID, FsoSpecialIndex, SpecialAddDate, SpecialNote,PhotoUrl,ClassID,MetaKey,MetaDescript
					 SpecialName = KS.G("SpecialName")
					 SpecialEName = KS.G("SpecialEName")
					 TemplateID = KS.G("TemplateID")
					 FsoSpecialIndex = KS.G("FsoSpecialIndex")
					 SpecialAddDate = KS.G("SpecialAddDate")
					 SpecialNote = Request.Form("SpecialNote")
					 PhotoUrl = KS.G("PhotoUrl")
					 MetaKey=Request.Form("MetaKey")
					 MetaDescript=Request.Form("MetaDescript")
					 ClassID = KS.ChkClng(KS.G("ClassID"))
			With KS 
				 If SpecialName <> "" Then
					If Len(SpecialName) >= 100 Then
						Call KS.AlertHistory("专题名称不能超过50个字符!", -1):Exit Sub
					End If
				 Else
					Call KS.AlertHistory("请输入专题名称!", -1):Exit Sub
				 End If
				 If SpecialEName <> "" Then
					If Len(SpecialEName) >= 50 Then
						Call KS.AlertHistory("专题英文名称不能超过50个字符!", -1):Exit Sub
					End If
					Set TempObj = Conn.Execute("Select SpecialEName,SpecialName from KS_Special where SpecialName='" & SpecialName & "' OR SpecialEName='" & SpecialEName & "'")
					If Not TempObj.EOF Then
						 If Trim(TempObj(0)) = SpecialEName Then
						   Call KS.alertHistory("数据库中已存在该专题英文名称!", -1)
						 Else
						   Call KS.alertHistory("数据库中已存在该专题名称!", -1)
						 End If
						.End
					End If
				 Else
					Call KS.alert("请输入专题英文名称!", "Special_Add.asp?ClassID=" & ClassID)
					.End
				 End If
				 If TemplateID = "" Then
					Call KS.alert("请选择专题模板", "Special_Add.asp?ClassID=" & ClassID)
					.End
				 End If
				
				  Set SpecialRS = Server.CreateObject("adodb.recordset")
				  SpecialSql = "select top 1 * from [KS_Special] Where (ID IS NULL)"
				  SpecialRS.Open SpecialSql, Conn, 1, 3
				  SpecialRS.AddNew
				  SpecialRS("ID") = Year(Now) & Month(Now) & Day(Now) & KS.MakeRandom(5)
				  SpecialRS("ClassID") = ClassID
				  SpecialRS("SpecialName") = SpecialName
				  SpecialRS("SpecialEName") = SpecialEName
				  SpecialRS("TemplateID") = TemplateID
				  SpecialRS("FsoSpecialIndex") = FsoSpecialIndex
				  SpecialRS("SpecialAddDate") = SpecialAddDate
				  SpecialRS("SpecialNote") = SpecialNote
				  SpecialRS("PhotoUrl") = PhotoUrl
				  SpecialRS("Creater") = KS.C("AdminName")
				  SpecialRS("MetaKey") = MetaKey
				  SpecialRS("MetaDescript") = MetaDescript
				  SpecialRS.Update
				  SpecialRS.MoveLast
				  Call KS.FileAssociation(1001,SpecialRS("SpecialID"),PhotoUrl&SpecialNote ,0)
				  SpecialRS.Close:Set SpecialRS = Nothing
				  .echo ("<script>if (confirm('添加专题成功,继续添加吗?')==true){location.href='KS.Special.asp?action=Add&ClassID=" & ClassID & "';}else{location.href='KS.Special.asp?Action=SpecialList&ClassID=" & ClassID & "';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?Opstr=" & server.URLEncode("内容管理 >> 专题管理") & "&ButtonSymbol=Disabled';}</script>")     
			End With
		End Sub
		'保存修改
		Sub SpecialEditSave()
			Dim TemplateRS, TemplateSql,TempObj, SpecialRS, SpecialSql,SpecialName, SpecialEName, TemplateID, FsoSpecialIndex, SpecialAddDate, SpecialNote,PhotoUrl,MetaKey,MetaDescript
					 SpecialName = KS.G("SpecialName")
					 TemplateID = KS.G("TemplateID")
					 FsoSpecialIndex = KS.G("FsoSpecialIndex")
					 SpecialAddDate = KS.G("SpecialAddDate")
					 SpecialNote = Request.Form("SpecialNote")
					 SpecialID   = KS.G("SpecialID")
					 PhotoUrl    = KS.G("PhotoUrl")
					 MetaKey=Request.Form("MetaKey")
					 MetaDescript=Request.Form("MetaDescript")
			With KS	 
				 If SpecialName <> "" Then
					If Len(SpecialName) >= 100 Then
						Call KS.AlertHistory("专题名称不能超过50个字符!", -1)
						.End
					End If
				 Else
					Call KS.AlertHistory("请输入专题名称!", -1)
					.End
				 End If
					
					Set TempObj = Conn.Execute("Select SpecialEName,SpecialName from KS_Special where SpecialName='" & SpecialName & "' And SpecialID<>" & SpecialID)
					If Not TempObj.EOF Then Call KS.alertHistory("数据库中已存在该专题名称!", -1): Exit Sub
				
				    If TemplateID = "" Then	Call KS.alertHistory("请选择专题模板",-1):Exit Sub

				
				  Set SpecialRS = Server.CreateObject("adodb.recordset")
				  SpecialSql = "select * from [KS_Special] Where SpecialID=" & SpecialID
				  SpecialRS.Open SpecialSql, Conn, 1, 3
				  SpecialRS("ClassID") = ClassID
				  SpecialRS("SpecialName") = SpecialName
				  SpecialRS("TemplateID") = TemplateID
				  SpecialRS("FsoSpecialIndex") = FsoSpecialIndex
				  SpecialRS("SpecialAddDate") = SpecialAddDate
				  SpecialRS("SpecialNote") = SpecialNote
				  SpecialRS("PhotoUrl") = PhotoUrl
				  SpecialRS("MetaKey") = MetaKey
				  SpecialRS("MetaDescript") = MetaDescript
				  SpecialRS.Update
				  Call KS.FileAssociation(1001,SpecialID,PhotoUrl&SpecialNote ,1)
				  SpecialRS.Close:Set SpecialRS = Nothing
				  .echo ("<script>alert('专题信息修改成功');location.href='KS.Special.asp?Action=SpecialList&ClassID=" & ClassID & "';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?Opstr=" & Server.URLEncode("所有专题") & "&ButtonSymbol=Disabled&ClassID=" & ClassID & "';</script>")     
			End With
		End Sub
		
		Sub GetChannelList()
		  With KS
		  	.echo (" <div class=""quicktz"" >")
		    Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		    RSObj.Open "Select ClassID,ClassName From KS_SpecialClass ",conn,1,1
		    If Not RSObj.Eof Then
			 .echo "<select OnChange=""location.href=this.value;"">"
			 .echo "<option value='KS.Special.asp?Action=SpecialList'>--按分类检索专题--</option>"
			 Do While Not RSObj.Eof
			   If ClassID=Trim(RSObj(0)) Then
			   	 .echo "<option value='KS.Special.asp?Action=SpecialList&ClassID=" & RSObj(0) &"' selected>" & RSObj(1) &"</option>"
			   Else
			   .echo "<option value='KS.Special.asp?Action=SpecialList&ClassID=" & RSObj(0) &"'>" & RSObj(1) &"</option>"
			   End If
			   RSObj.MoveNext
			 Loop
			  .echo "</select>"
			Else
			 .echo "<select style=""margin:-2px;"" OnChange=""location.href=this.value;"">"
			 .echo "<option value='KS.Special.asp'>--还没有添加任何分类--</option>"
			 .echo "</select>"
			End If
			 .echo "</div>"
			.echo "</ul>"
		  End With  
		End Sub
		
		'删除专题
		Sub SpecialDel()
			Dim K, ID, SpecialRS, FolderPath,Page
			Set SpecialRS = Server.CreateObject("Adodb.RecordSet")
			ID = Trim(KS.G("SpecialID"))
			Page = KS.G("Page")
			If ID="" Then KS.AlertHintScript "您没有选择专题" : Exit Sub
			ID = Split(ID, ",")
			For K = LBound(ID) To UBound(ID)
				 SpecialRS.Open "Select * FROM KS_Special Where SpecialID=" & KS.ChkClng(ID(K)), Conn, 1, 1
			  If SpecialRS.EOF And SpecialRS.BOF Then
				Call KS.AlertHistory("参数传递出错!", -1):Exit Sub
			  Else
				   If KS.Setting(95) = "/" Or KS.Setting(95) = "\" Then
					   FolderPath = KS.Setting(3) & SpecialRS("SpecialEName")
				   Else
					   FolderPath = KS.Setting(3) & KS.Setting(95) & SpecialRS("SpecialEName")
				   End If
			       If KS.DeleteFolder(FolderPath) = False Then  Call KS.AlertHistory("error!", -1):Exit Sub
				   Conn.Execute("Delete From KS_SpecialR Where SpecialID=" & ID(K))
				   Conn.Execute("Delete From KS_UploadFiles Where ChannelID=1001 and infoid=" & ID(K))
			  SpecialRS.Close
			  Conn.Execute("delete from ks_special Where SpecialID=" & KS.ChkClng(ID(K)))
			  End If
			Next
			If KeyWord = "" Then
			  Response.Write ("<script>location.href='KS.Special.asp?Action=SpecialList&Page=" & Page & "&ClassID=" & ClassID & "';</script>")
			Else
			  Response.Write ("<script>location.href='KS.Special.asp?Action=SpecialList&Page=" & Page & "&KeyWord=" & KeyWord & "&SearchType=" & SearchType & "&StartDate=" & StartDate & "&EndDate=" & EndDate & "';</script>")
			End If
		End Sub
		
		'从专题移出文章，图片，下载等
		Sub SpecialInfoDel()
		  Dim ID:ID = Replace(KS.G("ID")," ","")
		  ID=KS.FilterIDs(ID)
		  If ID="" Then
		   KS.AlertHintScript "出错!"
		  Else
		  Conn.Execute("Delete From KS_SpecialR Where ID in (" & ID & ")")
		  KS.AlertHintScript "恭喜,操作成功!"
		  End If
		End Sub
End Class
%> 
